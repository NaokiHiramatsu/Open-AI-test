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
if not os.path.exists('generated_files'):
    os.makedirs('generated_files')

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
                image_url = save_image_to_temp(image_data)
                ocr_text = ocr_image(image_url)
                file_data_text.append(f"ファイル名: {file.filename}\nOCR抽出内容:\n{ocr_text}")

        file_contents = "\n\n".join(file_data_text) if file_data_text else "なし"

        # Azure Search 呼び出し
        search_results = search_client.search(search_text=prompt, top=3)
        relevant_docs = []
        for doc in search_results:
            if 'chunk' in doc:
                relevant_docs.append(doc['chunk'])
            else:
                relevant_docs.append("該当するデータがありません")
        relevant_docs_text = "\n".join(relevant_docs)

        # AIへの入力データ生成と応答
        input_data_with_context = (
            f"アップロードされたファイル内容:\n{file_contents}\n\n関連ドキュメント:\n{relevant_docs_text}\n\nプロンプト:\n{prompt}"
        )
        response_content, output_decision = generate_ai_response_and_format(input_data_with_context, response_model)

        if output_decision not in ["excel", "pdf", "word"]:
            output_decision = "text"

        if output_decision == "text":
            session['chat_history'].append({
                'user': input_data_with_context,
                'assistant': response_content
            })
            return render_template('index.html', chat_history=session['chat_history'])
        else:
            # ファイル生成と保存
            file_data, mimetype, filename = generate_file(response_content, output_decision)
            temp_filename = f"{uuid.uuid4()}.{output_decision}"
            file_path = os.path.join('generated_files', temp_filename)

            # ファイル保存
            with open(file_path, 'wb') as f:
                file_data.seek(0)
                f.write(file_data.read())

            # ダウンロードリンクを生成
            download_url = url_for('download_file', filename=temp_filename, _external=True)
            print(f"Generated download URL: {download_url}")  # デバッグ用

            # セッションにリンクを格納
            session['chat_history'].append({
                'user': input_data_with_context,
                'assistant': f"<a class='download-link' href='{download_url}' target='_blank'>生成されたファイルをダウンロード</a>"
            })

            return render_template('index.html', chat_history=session['chat_history'])

    except Exception as e:
        print(f"Error in process_files_and_prompt: {e}")
        return jsonify({"error": f"エラーが発生しました: {str(e)}"}), 500

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join('generated_files', filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        print(f"File not found: {file_path}")
        return "File not found", 404

def save_image_to_temp(image_data):
    temp_filename = f"{uuid.uuid4()}.png"
    temp_path = os.path.join('generated_files', temp_filename)
    with open(temp_path, 'wb') as f:
        f.write(image_data)
    return temp_path

def ocr_image(image_url):
    ocr_url = f"{vision_endpoint}/vision/v3.2/ocr"
    headers = {"Ocp-Apim-Subscription-Key": vision_subscription_key}
    params = {"language": "ja", "detectOrientation": "true"}
    data = {"url": image_url}

    response = requests.post(ocr_url, headers=headers, params=params, json=data)
    response.raise_for_status()

    ocr_results = response.json()
    text_results = []
    for region in ocr_results.get("regions", []):
        for line in region.get("lines", []):
            text_results.append(" ".join(word["text"] for word in line["words"]))
    return "\n".join(text_results)

def generate_ai_response_and_format(input_data, deployment_name):
    """応答生成と出力形式判断をAIに依頼"""
    messages = [
        {
            "role": "system",
            "content": (
                "あなたは、システム内で直接ファイルを生成し、適切な形式（text, Excel, PDF, Word）を判断し生成します。"
                "Flaskの/downloadエンドポイントを使用してリンクをHTML <a>タグで提供してください。"
            )
        },
        {"role": "user", "content": input_data}
    ]
    try:
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )
        response_text = response['choices'][0]['message']['content']
        output_format = determine_output_format_from_response(response_text)
        return response_text, output_format
    except Exception as e:
        return f"ChatGPT 呼び出し中にエラーが発生しました: {str(e)}", "text"

def determine_output_format_from_response(response_content):
    """AI応答から出力形式を簡易的に判断"""
    if "excel" in response_content.lower():
        return "excel"
    elif "pdf" in response_content.lower():
        return "pdf"
    elif "word" in response_content.lower():
        return "word"
    else:
        return "text"

def generate_file(content, file_format):
    output = BytesIO()

    if file_format == "excel":
        rows = [row.split("\t") for row in content.split("\n") if row]
        df = pd.DataFrame(rows[1:], columns=rows[0])
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
    elif file_format == "pdf":
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for line in content.split("\n"):
            pdf.cell(200, 10, txt=line, ln=True, align='L')
        pdf.output(output)
    elif file_format == "word":
        doc = Document()
        for line in content.split("\n"):
            doc.add_paragraph(line)
        doc.save(output)
    else:
        output.write(content.encode('utf-8'))

    output.seek(0)
    return output, f"application/{file_format}", f"output.{file_format}"

if __name__ == '__main__':
    app.run(debug=True)
