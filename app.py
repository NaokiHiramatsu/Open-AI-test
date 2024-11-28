import openai
import os
from flask import Flask, request, jsonify, render_template, session, send_file, url_for
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import requests
import pandas as pd
from fpdf import FPDF
from docx import Document
from pptx import Presentation
import tempfile
import logging
import shutil

app = Flask(__name__)
app.secret_key = os.urandom(24)  # セッションを管理するための秘密鍵

# ログ設定
logging.basicConfig(level=logging.DEBUG)

# 環境変数からAPIキーやエンドポイントを取得
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY")
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME")

# Azure Cognitive Search の設定
search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT")
search_service_key = os.getenv("AZURE_SEARCH_KEY")
index_name = "vector-1730110777868"  # 使用するインデックス名を指定

search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# Computer Vision API の設定
vision_subscription_key = os.getenv("VISION_API_KEY")
vision_endpoint = os.getenv("VISION_ENDPOINT")

@app.before_request
def initialize_session():
    if 'chat_history' not in session:
        session['chat_history'] = []

@app.after_request
def cleanup_temp_files(response):
    temp_dir = tempfile.gettempdir()
    shutil.rmtree(temp_dir, ignore_errors=True)
    return response

@app.route('/')
def index():
    session.clear()
    return render_template('index.html', chat_history=[])

@app.route('/ocr', methods=['POST'])
def ocr_endpoint():
    image_url = request.json.get("image_url")
    if not image_url:
        return jsonify({"error": "No image URL provided"}), 400

    try:
        text = ocr_image_with_retry(image_url)
        return jsonify({"text": text})
    except Exception as e:
        logging.debug(f"OCR Error: {str(e)}")
        return jsonify({"error": str(e)}), 500

def ocr_image_with_retry(image_url, retries=3):
    ocr_url = vision_endpoint + "/vision/v3.2/ocr"
    headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key}
    params = {"language": "ja", "detectOrientation": "true"}
    data = {"url": image_url}

    for attempt in range(retries):
        try:
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
            if attempt < retries - 1:
                continue
            raise e

@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    try:
        file_data_text = [process_file(file) for file in files]
        file_contents = "\n\n".join(file_data_text) if file_data_text else "アップロードされたファイルはありません。"

        input_data = f"アップロードされたファイルの内容は次の通りです:\n{file_contents}\nプロンプト: {prompt}"
        search_results = search_client.search(search_text=prompt, top=3)
        relevant_docs = "\n".join([doc['chunk'] for doc in search_results])
        messages = [{"role": "system", "content": "あなたは有能なアシスタントです。"}]
        messages.append({"role": "user", "content": f"以下のドキュメントに基づいて質問に答えてください：\n{relevant_docs}\n質問: {input_data}"})

        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )

        response_content = response['choices'][0]['message']['content']
        session['chat_history'].append({'user': input_data, 'assistant': response_content})

        download_link = determine_file_type_and_generate(response_content)
        if download_link:
            response_content += f"\n\n[こちらからダウンロード]({download_link}) できます。"

        session['generated_content'] = response_content
        return render_template('index.html', chat_history=session['chat_history'])
    except Exception as e:
        logging.debug(f"File Processing Error: {str(e)}")
        return f"エラーが発生しました: {str(e)}"

def process_file(file):
    if file.filename.endswith('.xlsx'):
        df = pd.read_excel(file, engine='openpyxl')
        columns_text = " | ".join(df.columns)
        rows_text = [" | ".join(map(str, row)) for _, row in df.iterrows()]
        return f"ファイル名: {file.filename}\n{columns_text}\n" + "\n".join(rows_text)
    elif file.filename.endswith('.pdf'):
        return f"ファイル名: {file.filename}\n{ocr_pdf(file)}"
    elif file.filename.endswith('.docx'):
        doc = Document(file)
        return f"ファイル名: {file.filename}\n" + "\n".join([p.text for p in doc.paragraphs])
    elif file.filename.endswith('.pptx'):
        ppt = Presentation(file)
        slides_text = [
            shape.text for slide in ppt.slides for shape in slide.shapes if hasattr(shape, "text")
        ]
        return f"ファイル名: {file.filename}\n" + "\n".join(slides_text)
    else:
        return f"ファイル形式が対応していません: {file.filename}"

def determine_file_type_and_generate(response_content):
    if "Excel" in response_content:
        return generate_file('excel', response_content)
    elif "PDF" in response_content:
        return generate_file('pdf', response_content)
    elif "Word" in response_content:
        return generate_file('word', response_content)
    elif "テキスト" in response_content or "txt" in response_content:
        return generate_file('txt', response_content)
    return None

def generate_file(file_type, content):
    try:
        temp_dir = tempfile.mkdtemp()
        if file_type == 'txt':
            temp_path = os.path.join(temp_dir, "output.txt")
            with open(temp_path, 'w', encoding='utf-8') as f:
                f.write(content)
        elif file_type == 'pdf':
            temp_path = os.path.join(temp_dir, "output.pdf")
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, content)
            pdf.output(temp_path)
        elif file_type == 'excel':
            temp_path = os.path.join(temp_dir, "output.xlsx")
            df = pd.DataFrame({"Content": [content]})
            df.to_excel(temp_path, index=False)
        elif file_type == 'word':
            temp_path = os.path.join(temp_dir, "output.docx")
            doc = Document()
            doc.add_heading('Generated Content', level=1)
            doc.add_paragraph(content)
            doc.save(temp_path)
        return url_for('download_file', file_path=temp_path)
    except Exception as e:
        logging.debug(f"File Generation Error: {str(e)}")
        return None

@app.route('/download_file')
def download_file():
    file_path = request.args.get('file_path', None)
    if file_path and os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found.", 404

if __name__ == '__main__':
    app.run(debug=True)
