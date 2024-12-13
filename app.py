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
format_model = os.getenv("OPENAI_FORMAT_MODEL", "format-model")

search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT", "https://search-service.azure.com")
search_service_key = os.getenv("AZURE_SEARCH_KEY", "your-search-key")
index_name = os.getenv("AZURE_SEARCH_INDEX_NAME", "your-index-name")

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

        # AIへの入力データ生成
        input_data_with_context = (
            f"アップロードされたファイル内容:\n{file_contents}\n\n関連ドキュメント:\n{relevant_docs_text}\n\nプロンプト:\n{prompt}"
        )

        # 応答生成
        response_content = generate_ai_response(input_data_with_context, response_model)

        # 出力形式判断
        output_decision = determine_file_format(response_content, format_model)

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
            try:
                with open(file_path, 'wb') as f:
                    file_data.seek(0)
                    f.write(file_data.read())
                print(f"File saved at: {file_path}")
            except Exception as save_error:
                print(f"Error saving file: {save_error}")
                raise

            # ダウンロードリンクを生成
            download_url = url_for('download_file', filename=temp_filename, _external=True)
            print(f"Generated download URL: {download_url}")

            # セッションにリンクを格納
            session['chat_history'].append({
                'user': input_data_with_context,
                'assistant': f"以下のリンクから生成されたファイルをダウンロードできます： "
                             f"<a href='{download_url}' target='_blank'>{filename}</a>"
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

def generate_ai_response(input_data, deployment_name):
    """応答生成を担当するAI"""
    messages = [
        {
            "role": "system",
            "content": (
                "あなたは、システム内で直接ファイルを生成することが可能な有能なアシスタントです。"
                "ユーザーから提供されたデータやプロンプトに基づき、以下の形式でファイルを生成してください:"
                " - 'Excel'（Excelファイルで出力）"
                " - 'PDF'（PDFファイルで出力）"
                " - 'Word'（Wordファイルで出力）"
                "出力可能なファイル形式がない場合でも、スクリプトの例ではなく明確な理由をユーザーに伝えてください。"
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
        return response['choices'][0]['message']['content']
    except Exception as e:
        return f"ChatGPT 呼び出し中にエラーが発生しました: {str(e)}"

def determine_file_format(response_content, deployment_name):
    """出力形式の判断を担当するAI"""
    try:
        prompt = f"""
        以下の応答を基に、どの出力形式が適切か選択してください。必ず1つを選んでください。
        - "text" （テキスト形式で返す場合に選択）
        - "Excel" （Excelファイルが適切な場合に選択）
        - "PDF" （PDFファイルが適切な場合に選択）
        - "Word" （Wordファイルが適切な場合に選択）

        応答内容:
        {response_content}
        """
        response = openai.Completion.create(
            engine=deployment_name,
            prompt=prompt,
            max_tokens=50,
            temperature=0.3
        )
        return response['choices'][0]['text'].strip().lower()
    except Exception:
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
    if file_format == "excel":
        return output, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx"
    elif file_format == "pdf":
        return output, "application/pdf", "output.pdf"
    elif file_format == "word":
        return output, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx"
    else:
        return output, "text/plain", "output.txt"

if __name__ == '__main__':
    app.run(debug=True)
