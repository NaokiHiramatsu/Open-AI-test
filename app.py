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
import logging

# ログ設定
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Azure OpenAI設定
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# Azure Cognitive Search 設定
search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
search_service_key = os.getenv("AZURE_SEARCH_KEY")
index_name = "vector-1730110777868"  # インデックス名
search_client = None
if search_service_endpoint and search_service_key:
    search_client = SearchClient(
        endpoint=search_service_endpoint,
        index_name=index_name,
        credential=AzureKeyCredential(search_service_key)
    )

# Azure Computer Vision 設定
vision_subscription_key = os.getenv("VISION_API_KEY")
vision_endpoint = os.getenv("VISION_ENDPOINT")

# ホームページ表示
@app.route('/')
def index():
    session.clear()
    return render_template('index.html', chat_history=[])

# OCR（画像）機能
@app.route('/ocr', methods=['POST'])
def ocr_endpoint():
    try:
        image_url = request.json.get("image_url")
        if not image_url:
            return jsonify({"error": "No image URL provided"}), 400
        text = ocr_image(image_url)
        return jsonify({"text": text})
    except Exception as e:
        logging.error(f"OCR error: {e}")
        return jsonify({"error": str(e)}), 500

def ocr_image(image_url):
    ocr_url = f"{vision_endpoint}/vision/v3.2/ocr"
    headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key}
    params = {"language": "ja", "detectOrientation": "true"}
    data = {"url": image_url}
    response = requests.post(ocr_url, headers=headers, params=params, json=data)
    response.raise_for_status()
    ocr_results = response.json()
    text_results = [
        " ".join([word["text"] for word in line["words"]])
        for region in ocr_results.get("regions", [])
        for line in region.get("lines", [])
    ]
    return "\n".join(text_results)

def ocr_pdf(file):
    ocr_url = f"{vision_endpoint}/vision/v3.2/read/analyze"
    headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key, "Content-Type": "application/pdf"}
    response = requests.post(ocr_url, headers=headers, data=file.read())
    response.raise_for_status()
    operation_url = response.headers["Operation-Location"]
    analysis = {}
    while "analyzeResult" not in analysis:
        response_final = requests.get(operation_url, headers=headers)
        analysis = response_final.json()
    text_results = [
        line["text"]
        for read_result in analysis["analyzeResult"]["readResults"]
        for line in read_result["lines"]
    ]
    return "\n".join(text_results)

# ファイルとプロンプト処理
@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    try:
        files = request.files.getlist('files')
        prompt = request.form.get('prompt', '')
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
            elif file.filename.endswith('.pdf'):
                file_data_text.append(ocr_pdf(file))
            elif file.filename.endswith('.docx'):
                doc = Document(file)
                file_data_text.append("\n".join([para.text for para in doc.paragraphs]))
            elif file.filename.endswith('.pptx'):
                ppt = Presentation(file)
                for slide in ppt.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            file_data_text.append(shape.text)
        file_content_combined = "\n\n".join(file_data_text)
        input_data = f"プロンプト: {prompt}\nアップロードされたファイル内容:\n{file_content_combined}"
        relevant_docs = ""
        if search_client:
            try:
                search_results = search_client.search(search_text=prompt, top=3)
                relevant_docs = "\n".join([doc['content'] for doc in search_results])
            except Exception as e:
                logging.error(f"Search error: {e}")
        input_data += f"\n関連ドキュメント:\n{relevant_docs}"
        messages.append({"role": "user", "content": input_data})
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )
        response_content = response['choices'][0]['message']['content']
        download_url = None
        if "ダウンロード" in prompt:
            file_path = generate_file(response_content, file_type="xlsx")
            if file_path:
                download_url = url_for('download_file', filename=os.path.basename(file_path), _external=True)
        session['chat_history'].append({'user': prompt, 'assistant': response_content, 'download_url': download_url})
        return render_template('index.html', chat_history=session['chat_history'])
    except Exception as e:
        logging.error(f"Error: {e}")
        return "内部エラーが発生しました。", 500

def generate_file(content, file_type):
    try:
        if file_type == "xlsx":
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            df = pd.DataFrame({"Generated Content": [content]})
            df.to_excel(temp_file.name, index=False)
            return temp_file.name
    except Exception as e:
        logging.error(f"File generation error: {e}")
        return None

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(tempfile.gettempdir(), filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return "ファイルが見つかりません。", 404
    except Exception as e:
        logging.error(f"Download error: {e}")
        return "エラーが発生しました。", 500

if __name__ == '__main__':
    app.run(debug=True)
