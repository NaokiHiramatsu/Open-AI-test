import openai
import os
from flask import Flask, request, render_template, session, send_file, url_for, safe_join
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd
from docx import Document
from pptx import Presentation
import requests
import tempfile
import re  # 正規表現を利用してダウンロード意図を判別

app = Flask(__name__)
app.secret_key = os.urandom(24)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
search_service_key = os.getenv("AZURE_SEARCH_KEY")
index_name = "vector-1730110777868"

vision_subscription_key = os.getenv("VISION_API_KEY")
vision_endpoint = os.getenv("VISION_ENDPOINT")

search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# ファイル保存用ディレクトリ
output_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# ファイル生成関数
def generate_file(content, file_type="xlsx"):
    try:
        temp_file_path = tempfile.mktemp(suffix=f".{file_type}")
        if file_type == "xlsx":
            df = pd.DataFrame({"Content": [content]})
            df.to_excel(temp_file_path, index=False)
        elif file_type == "pdf":
            from fpdf import FPDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, content)
            pdf.output(temp_file_path)
        elif file_type == "docx":
            doc = Document()
            doc.add_paragraph(content)
            doc.save(temp_file_path)
        return temp_file_path
    except Exception as e:
        print(f"File generation failed: {e}")
        return None

# OCR機能
def ocr_image(image_url):
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

# ダウンロード希望を判定する関数
def is_download_requested(prompt):
    """
    プロンプト内にダウンロードを希望する意図があるかを判定。
    """
    download_keywords = [
        "ダウンロード", "保存", "エクスポート", "ファイルとして出力", 
        "Excelで", "ファイルが欲しい", "出力したい", "保存したい"
    ]
    pattern = "|".join(download_keywords)
    return bool(re.search(pattern, prompt, re.IGNORECASE))

@app.route('/')
def index():
    session.clear()
    return render_template('index.html', chat_history=[])

@app.route('/ocr', methods=['POST'])
def ocr_endpoint():
    image_url = request.json.get("image_url")
    if not image_url:
        return {"error": "No image URL provided"}, 400

    try:
        text = ocr_image(image_url)
        return {"text": text}
    except Exception as e:
        return {"error": str(e)}, 500

@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    # ダウンロード希望の判定
    download_requested = is_download_requested(prompt)

    if 'chat_history' not in session:
        session['chat_history'] = []

    messages = [{"role": "system", "content": "あなたは有能なアシスタントです。"}]
    for entry in session['chat_history']:
        messages.append({"role": "user", "content": entry['user']})
        messages.append({"role": "assistant", "content": entry['assistant']})

    file_data_text = []
    for file in files:
        if file.filename.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
            file_data_text.append(df.to_csv(index=False))
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

    file_content_combined = "\n\n".join(file_data_text)
    input_data = f"アップロードされたファイルの内容:\n{file_content_combined}\nプロンプト: {prompt}"
    messages.append({"role": "user", "content": input_data})

    search_results = search_client.search(search_text=prompt, top=3)
    relevant_docs = "\n".join([doc['chunk'] for doc in search_results])
    messages.append({"role": "user", "content": f"以下に基づいて回答してください:\n{relevant_docs}"})

    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=messages,
        max_tokens=2000
    )
    response_content = response['choices'][0]['message']['content']

    download_url = None
    if download_requested:
        file_path = generate_file(response_content)
        if file_path:
            download_url = url_for('download_file', filename=os.path.basename(file_path), _external=True)

    session['chat_history'].append({'user': input_data, 'assistant': response_content, 'download_url': download_url})
    return render_template('index.html', chat_history=session['chat_history'])

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = safe_join(output_dir, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return "File not found.", 404
    except Exception as e:
        return f"Error while accessing file: {e}", 500

if __name__ == '__main__':
    app.run(debug=True)
