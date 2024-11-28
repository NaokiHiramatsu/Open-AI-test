import openai
import os
from flask import Flask, request, jsonify, render_template, session, send_file, url_for
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd
from fpdf import FPDF
from docx import Document
from pptx import Presentation
import tempfile
import requests

app = Flask(__name__)
app.secret_key = os.urandom(24)  # セッション管理の秘密鍵

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# Azure Cognitive Search の設定
search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
search_service_key = os.getenv("AZURE_SEARCH_KEY")
index_name = "vector-1730110777868"

search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# Computer Vision API の設定
vision_subscription_key = os.getenv("VISION_API_KEY")
vision_endpoint = os.getenv("VISION_ENDPOINT")

# ファイル生成用関数
def generate_file(file_type, content):
    try:
        temp_file_path = tempfile.mktemp(suffix=f".{file_type}")
        if file_type == 'xlsx':  # Excelファイル生成
            if isinstance(content, list):
                df = pd.DataFrame(content)
            else:
                df = pd.DataFrame({"Content": [content]})
            df.to_excel(temp_file_path, index=False)
        elif file_type == 'pdf':  # PDFファイル生成
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, content)
            pdf.output(temp_file_path)
        elif file_type == 'docx':  # Wordファイル生成
            doc = Document()
            doc.add_heading('Generated Content', level=1)
            doc.add_paragraph(content)
            doc.save(temp_file_path)
        return temp_file_path
    except Exception as e:
        print(f"File generation failed ({file_type}): {e}")
        return None

# OCR機能
def ocr_image(image_url):
    try:
        ocr_url = vision_endpoint + "/vision/v3.2/ocr"
        headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key}
        params = {"language": "ja", "detectOrientation": "true"}
        data = {"url": image_url}
        response = requests.post(ocr_url, headers=headers, params=params, json=data)
        response.raise_for_status()

        ocr_results = response.json()
        text_results = []
        for region in ocr_results.get("regions", []):
            for line in region.get("lines", []):
                line_text = " ".join([word["text"] for word in line["words"]])
                text_results.append(line_text)
        return "\n".join(text_results)
    except Exception as e:
        print(f"OCR failed: {e}")
        return None

def ocr_pdf(file):
    try:
        ocr_url = vision_endpoint + "/vision/v3.2/read/analyze"
        headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key, "Content-Type": "application/pdf"}
        response = requests.post(ocr_url, headers=headers, data=file.read())
        response.raise_for_status()

        operation_url = response.headers["Operation-Location"]
        analysis = {}
        while not "analyzeResult" in analysis:
            response_final = requests.get(operation_url, headers={"Ocp-Apim-Subscription-Key": vision_subscription_key})
            analysis = response_final.json()

        text_results = []
        for read_result in analysis["analyzeResult"]["readResults"]:
            for line in read_result["lines"]:
                text_results.append(line["text"])
        return "\n".join(text_results)
    except Exception as e:
        print(f"PDF OCR failed: {e}")
        return None

@app.route('/')
def index():
    session.clear()  # セッションをリセット
    return render_template('index.html', chat_history=[])

@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    if 'chat_history' not in session:
        session['chat_history'] = []

    # チャット履歴を構成
    messages = [{"role": "system", "content": "あなたは有能なアシスタントです。"}]
    for entry in session['chat_history']:
        messages.append({"role": "user", "content": entry['user']})
        messages.append({"role": "assistant", "content": entry['assistant']})

    # アップロードされたファイルの内容を処理
    file_data_text = []
    for file in files:
        if file.filename.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
            file_data_text.append(df.to_csv(index=False))
        elif file.filename.endswith('.pdf'):
            file_data_text.append(ocr_pdf(file))
        elif file.filename.endswith('.docx'):
            doc = Document(file)
            file_data_text.append("\n".join([para.text for para in doc.paragraphs]))
        elif file.filename.endswith('.pptx'):
            ppt = Presentation(file)
            ppt_text = []
            for slide in ppt.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        ppt_text.append(shape.text)
            file_data_text.append("\n".join(ppt_text))

    # ファイル内容とプロンプトを結合
    file_content_combined = "\n\n".join(file_data_text)
    input_data = f"アップロードされたファイルの内容:\n{file_content_combined}\nプロンプト: {prompt}"
    messages.append({"role": "user", "content": input_data})

    # Azure Cognitive Search
    search_results = search_client.search(search_text=prompt, top=3)
    relevant_docs = "\n".join([doc['chunk'] for doc in search_results])
    messages.append({"role": "user", "content": f"以下に基づいて回答してください:\n{relevant_docs}"})

    # OpenAI API呼び出し
    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=messages,
        max_tokens=2000
    )
    response_content = response['choices'][0]['message']['content']

    # ファイル生成とダウンロードリンクの追加
    file_path = generate_file('xlsx', response_content)
    if file_path:
        download_url = url_for('download_file', file_path=file_path, _external=True)
        response_content += f"\n\n<a href='{download_url}' download>こちらからダウンロードできます</a>"

    session['chat_history'].append({'user': input_data, 'assistant': response_content})
    return render_template('index.html', chat_history=session['chat_history'])

@app.route('/download_file')
def download_file():
    file_path = request.args.get('file_path')
    if file_path and os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "File not found.", 404

if __name__ == '__main__':
    app.run(debug=True)
