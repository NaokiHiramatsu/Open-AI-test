import os
import uuid
from flask import Flask, request, jsonify, render_template, session, send_file, url_for
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import openai
import pandas as pd
from io import BytesIO
from flask_session import Session
import requests
from fpdf import FPDF
from docx import Document

# Flaskアプリの設定
app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['SESSION_TYPE'] = 'filesystem'
Session(app)

# 環境変数の設定
openai.api_type = "azure"
openai.api_base = os.getenv("OPENAI_API_BASE", "https://example.openai.azure.com")
openai.api_version = "2024-08-01-preview"
openai.api_key = os.getenv("OPENAI_API_KEY", "your-api-key")
response_model = os.getenv("OPENAI_RESPONSE_MODEL", "response-model")

search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT", "https://search-service.azure.com")
search_service_key = os.getenv("AZURE_SEARCH_KEY", "your-search-key")
index_name = os.getenv("AZURE_SEARCH_INDEX_NAME", "your-index-name")

vision_subscription_key = os.getenv("VISION_API_KEY", "your-vision-key")
vision_endpoint = os.getenv("VISION_ENDPOINT", "https://vision.azure.com")

# Azure Search クライアントの設定
try:
    if not all([search_service_endpoint, search_service_key, index_name]):
        raise ValueError("Azure Search の環境変数が正しく設定されていません。")

    search_client = SearchClient(
        endpoint=search_service_endpoint,
        index_name=index_name,
        credential=AzureKeyCredential(search_service_key)
    )
    print("Search client initialized successfully.")
except Exception as e:
    search_client = None
    print(f"SearchClient initialization failed: {e}")

# 一時ファイル保存ディレクトリを作成
SAVE_DIR = "generated_files"
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

@app.route('/')
def index():
    session.clear()
    return render_template('index.html', chat_history=[])

@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    if 'chat_history' not in session:
        session['chat_history'] = []

    if search_client is None:
        error_message = "SearchClient が初期化されていません。サーバー設定を確認してください。"
        print(error_message)
        return jsonify({"error": error_message}), 500

    try:
        # ファイル処理
        file_data_text = []
        for file in files:
            if file and file.filename.endswith('.xlsx'):
                df = pd.read_excel(file, engine='openpyxl')
                columns = df.columns.tolist()
                rows_text = df.to_string(index=False)
                file_data_text.append(f"ファイル名: {file.filename}\n列: {columns}\n内容:\n{rows_text}")
            elif file and file.filename.endswith(('.png', '.jpg', '.jpeg')):
                image_data = file.read()
                image_filename = save_image_to_temp(image_data)
                file_data_text.append(f"ファイル名: {file.filename}\nOCR抽出内容:\n{ocr_image(image_filename)}")

        file_contents = "\n\n".join(file_data_text) if file_data_text else "なし"

        # Azure Search 呼び出し
        search_results = search_client.search(search_text=prompt, top=3)
        relevant_docs = [doc.get('chunk', "該当するデータがありません") for doc in search_results]
        relevant_docs_text = "\n".join(relevant_docs)

        # AI応答生成と出力形式判断
        input_data = f"アップロードされたファイル内容:\n{file_contents}\n\n関連ドキュメント:\n{relevant_docs_text}\n\nプロンプト:\n{prompt}"
        response_content, output_format = generate_ai_response_and_format(input_data, response_model)

        # ファイル生成と保存（正しい拡張子設定）
        file_data, mime_type, file_format = generate_file(response_content, output_format)
        if file_format == "excel":
            temp_filename = f"{uuid.uuid4()}.xlsx"
        elif file_format == "pdf":
            temp_filename = f"{uuid.uuid4()}.pdf"
        elif file_format == "word":
            temp_filename = f"{uuid.uuid4()}.docx"
        else:
            temp_filename = f"{uuid.uuid4()}.txt"

        file_path = os.path.join(SAVE_DIR, temp_filename)
        with open(file_path, 'wb') as f:
            file_data.seek(0)
            f.write(file_data.read())

        download_url = url_for('download_file', filename=temp_filename, _external=True)
        session['chat_history'].append({
            'user': input_data,
            'assistant': f"<a href='{download_url}' target='_blank'>生成されたファイルをダウンロード</a>"
        })
        return render_template('index.html', chat_history=session['chat_history'])

    except Exception as e:
        print(f"Error during file processing: {e}")
        return jsonify({"error": f"エラーが発生しました: {e}"}), 500

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.abspath(os.path.join(SAVE_DIR, filename))
    if not file_path.startswith(os.path.abspath(SAVE_DIR)):
        print(f"Invalid file path: {file_path}")
        return "Invalid file path", 400
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        try:
            return send_file(file_path, as_attachment=True)
        except Exception as e:
            print(f"Error sending file: {e}")
            return "ファイル送信中にエラーが発生しました。", 500
    return "File not found or file is empty", 404

def save_image_to_temp(image_data):
    temp_filename = f"{uuid.uuid4()}.png"
    temp_path = os.path.join(SAVE_DIR, temp_filename)
    with open(temp_path, 'wb') as f:
        f.write(image_data)
    return temp_path

def ocr_image(image_path):
    ocr_url = f"{vision_endpoint}/vision/v3.2/ocr"
    headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key}
    with open(image_path, "rb") as img:
        response = requests.post(ocr_url, headers=headers, files={"file": img})
    response.raise_for_status()
    ocr_results = response.json()
    return "\n".join([" ".join(word['text'] for word in line['words']) for region in ocr_results.get('regions', []) for line in region.get('lines', [])])

def generate_ai_response_and_format(input_data, deployment_name):
    messages = [
        {"role": "system", "content": (
            "あなたは、システム内で直接ファイルを生成し、適切な形式（text, Excel, PDF, Word）を判断し生成します。"
            "Flaskの/downloadエンドポイントを使用してリンクをHTML <a>タグで提供してください。"
        )},
        {"role": "user", "content": input_data}
    ]
    response = openai.ChatCompletion.create(engine=deployment_name, messages=messages, max_tokens=2000)
    response_text = response['choices'][0]['message']['content']
    output_format = determine_output_format_from_response(response_text)
    return response_text, output_format

def determine_output_format_from_response(response_content):
    if "excel" in response_content.lower(): return "excel"
    if "pdf" in response_content.lower(): return "pdf"
    if "word" in response_content.lower(): return "word"
    return "text"

def generate_file(content, file_format):
    output = BytesIO()
    if file_format == "excel":
        pd.DataFrame([[content]]).to_excel(output, index=False, engine='xlsxwriter')
    elif file_format == "pdf":
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, content)
        pdf.output(output)
    elif file_format == "word":
        doc = Document()
        doc.add_paragraph(content)
        doc.save(output)
    else:
        output.write(content.encode("utf-8"))
    output.seek(0)
    return output, "application/octet-stream", file_format

if __name__ == '__main__':
    app.run(debug=True)
