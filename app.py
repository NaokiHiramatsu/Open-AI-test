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

# Azure Search接続確認関数
def check_search_connection():
    try:
        test_response = requests.get(
            f"{search_service_endpoint}/indexes?api-version=2021-04-30-Preview",
            headers={"api-key": search_service_key}
        )
        test_response.raise_for_status()
        print("Azure Search connection is successful.")
    except Exception as e:
        print(f"Failed to connect to Azure Search: {e}")
        return False
    return True

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

    try:
        file_data_text = []
        excel_dataframes = []  # Excelデータを格納するリスト

        for file in files:
            if file and file.filename.endswith('.xlsx'):
                df = pd.read_excel(file, engine='openpyxl')
                excel_dataframes.append(df)
                columns = df.columns.tolist()
                rows_text = df.to_string(index=False)
                file_data_text.append(f"ファイル名: {file.filename}\n列: {columns}\n内容:\n{rows_text}")
            elif file and file.filename.endswith(('.png', '.jpg', '.jpeg')):
                image_data = file.read()
                image_filename = save_image_to_temp(image_data)
                file_data_text.append(f"ファイル名: {file.filename}\nOCR抽出内容:\n{ocr_image(image_filename)}")

        file_contents = "\n\n".join(file_data_text) if file_data_text else "なし"

        # Azure Search 呼び出し
        if search_client and check_search_connection():
            try:
                search_results = search_client.search(search_text=prompt, top=3)
                relevant_docs = []
                for result in search_results:
                    print(f"Search result: {result}")  # デバッグ用に検索結果をログ出力
                    headers = result.get("chunk_headers", [])
                    rows = result.get("chunk_rows", [])
                    if headers and rows:
                        df = pd.DataFrame(data=rows, columns=headers)
                        excel_dataframes.append(df)
                        relevant_docs.append(f"データ取得: {headers} {rows}")
                relevant_docs_text = "\n".join(relevant_docs)
            except Exception as e:
                relevant_docs_text = f"Azure Search クエリ実行中にエラーが発生しました: {e}"
                print(relevant_docs_text)
        else:
            relevant_docs_text = "Azure Search クライアントが初期化されていないか、接続エラーが発生しています。"
            print(relevant_docs_text)

        # AI応答生成と出力形式判断
        input_data = f"アップロードされたファイル内容:\n{file_contents}\n\n関連ドキュメント:\n{relevant_docs_text}\n\nプロンプト:\n{prompt}"
        response_content, output_format = generate_ai_response_and_format(input_data, response_model)

        # 応答を分割して処理
        chat_output, file_output = parse_response_content(response_content)

        # Excelへの統合処理
        if excel_dataframes:
            combined_df = pd.concat(excel_dataframes, ignore_index=True)
            output_format = "xlsx"
            excel_buffer = BytesIO()  # BytesIOオブジェクトを作成
            combined_df.to_excel(excel_buffer, index=False, engine='openpyxl')  # Excelに書き込む
            excel_buffer.seek(0)  # ファイルの先頭にポインタを戻す
            file_output = excel_buffer  # ファイル出力用に渡す

        # チャット履歴用
        session['chat_history'].append({
            'user': input_data,
            'assistant': chat_output
        })

        # ファイル生成と保存
        file_data, mime_type, file_format = generate_file(file_output, output_format)
        temp_filename = f"{uuid.uuid4()}.{file_format}"
        file_path = os.path.join(SAVE_DIR, temp_filename)
        with open(file_path, 'wb') as f:
            file_data.seek(0)
            f.write(file_data.read())

        download_url = url_for('download_file', filename=temp_filename, _external=True)
        session['chat_history'][-1]['assistant'] += f" <a href='{download_url}' target='_blank'>生成されたファイルをダウンロード</a>"

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
            mimetype_map = {
                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'pdf': 'application/pdf',
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'txt': 'text/plain',
                'png': 'image/png',
                'jpg': 'image/jpeg'
            }
            ext = filename.split('.')[-1]
            mimetype = mimetype_map.get(ext, 'application/octet-stream')

            return send_file(file_path, mimetype=mimetype, as_attachment=True, download_name=filename)
        except Exception as e:
            print(f"Error sending file: {e}")
            return "ファイル送信中にエラーが発生しました。", 500

    print(f"File not found or empty: {file_path}")
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
            "生成するExcelファイルには、1行目に列名、2行目以降にデータ行を含める必要があります。"
            "ファイルで出力すべき内容以外の文章はテキスト形式で返してください。"
            "必ずファイル形式と内容を判断し、必要に応じて表形式を正しく出力してください。"
            "Flaskの/downloadエンドポイントを使用してリンクをHTML <a>タグで提供してください。"
        )},
        {"role": "user", "content": input_data}
    ]
    response = openai.ChatCompletion.create(engine=deployment_name, messages=messages, max_tokens=2000)
    response_text = response['choices'][0]['message']['content']
    output_format = determine_output_format_from_response(response_text)
    return response_text, output_format

def determine_output_format_from_response(response_content):
    if "excel" in response_content.lower(): return "xlsx"
    if "pdf" in response_content.lower(): return "pdf"
    if "word" in response_content.lower(): return "docx"
    return "txt"

def generate_file(content, file_format):
    output = BytesIO()
    if file_format == "xlsx":
        if isinstance(content, BytesIO):  # BytesIOオブジェクトの場合
            content.seek(0)  # ポインタを先頭に戻す
            return content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", file_format
        else:  # 通常の文字列コンテンツの場合
            rows = [row.split("\t") for row in content.split("\n") if row]
            df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
    elif file_format == "pdf":
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, content)
        pdf.output(output)
    elif file_format == "docx":
        doc = Document()
        doc.add_paragraph(content)
        doc.save(output)
    else:
        output.write(content.encode("utf-8"))
    output.seek(0)
    return output, "application/octet-stream", file_format

def parse_response_content(response_content):
    if "ファイル内容:" in response_content:
        parts = response_content.split("ファイル内容:", 1)
        return parts[0].strip(), parts[1].strip()
    return response_content, ""

if __name__ == '__main__':
    app.run(debug=True)
