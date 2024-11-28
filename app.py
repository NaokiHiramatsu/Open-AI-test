import openai
import os
from flask import Flask, request, jsonify, render_template, session, send_file, url_for, safe_join
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
import pandas as pd
from fpdf import FPDF
from docx import Document
from pptx import Presentation
import shutil

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

search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

# ファイル保存用ディレクトリを定義
output_dir = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# ファイル生成関数
def generate_file(file_type, content, filename="output"):
    try:
        file_path = os.path.join(output_dir, f"{filename}.{file_type}")
        if file_type == 'xlsx':  # Excelファイル生成
            df = pd.DataFrame({"Content": [content]})
            df.to_excel(file_path, index=False)
        elif file_type == 'pdf':  # PDFファイル生成
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, content)
            pdf.output(file_path)
        elif file_type == 'docx':  # Wordファイル生成
            doc = Document()
            doc.add_heading('Generated Content', level=1)
            doc.add_paragraph(content)
            doc.save(file_path)
        return file_path
    except Exception as e:
        print(f"File generation failed: {e}")
        return None

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
            file_data_text.append("PDF処理未実装")  # 例として記載
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

    # ファイル生成とリンクの生成
    file_path = generate_file('xlsx', response_content, filename="response")
    if file_path:
        download_url = url_for('download_file', filename="response.xlsx", _external=True)
        response_content += f"\n\n<a href='{download_url}' download>こちらからダウンロードできます</a>"

    session['chat_history'].append({'user': input_data, 'assistant': response_content})
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
