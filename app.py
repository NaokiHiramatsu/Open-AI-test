import os
from flask import Flask, request, jsonify, render_template, session, send_file
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
deployment_name = os.getenv("OPENAI_DEPLOYMENT_NAME", "default-deployment")

search_service_endpoint = os.getenv("AZURE_SEARCH_ENDPOINT", "https://search-service.azure.com")
search_service_key = os.getenv("AZURE_SEARCH_KEY", "your-search-key")
index_name = os.getenv("AZURE_SEARCH_INDEX_NAME", "your-index-name")

vision_subscription_key = os.getenv("VISION_API_KEY", "your-vision-key")
vision_endpoint = os.getenv("VISION_ENDPOINT", "https://vision.azure.com")

# Azure Search クライアントの設定
search_client = SearchClient(
    endpoint=search_service_endpoint,
    index_name=index_name,
    credential=AzureKeyCredential(search_service_key)
)

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
        text = ocr_image(image_url)
        return jsonify({"text": text})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

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

@app.route('/process_files_and_prompt', methods=['POST'])
def process_files_and_prompt():
    files = request.files.getlist('files')
    prompt = request.form.get('prompt', '')

    if 'chat_history' not in session:
        session['chat_history'] = []

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

        # 生成AIによる判断
        response_content, output_decision = generate_ai_response_with_decision(input_data_with_context)

        if output_decision == "text":
            session['chat_history'].append({
                'user': input_data_with_context,
                'assistant': response_content
            })
            return render_template('index.html', chat_history=session['chat_history'])
        else:
            file_data, mimetype, filename = generate_file(response_content, output_decision)
            return send_file(file_data, as_attachment=True, download_name=filename, mimetype=mimetype)

    except Exception as e:
        return jsonify({"error": f"エラーが発生しました: {str(e)}"}), 500

def generate_ai_response_with_decision(input_data):
    """生成AIで応答と出力形式の判断を生成"""
    messages = [
        {"role": "system", "content": "あなたは有能なアシスタントです。"},
        {"role": "user", "content": input_data}
    ]

    try:
        # 応答生成
        response = openai.ChatCompletion.create(
            engine=deployment_name,
            messages=messages,
            max_tokens=2000
        )
        response_content = response['choices'][0]['message']['content']

        # 出力形式の判断
        format_prompt = f"""
        以下の応答を基に、どの出力形式が適切か選択してください。
        - "text" （テキストで返す）
        - "Excel" （Excelファイルで出力）
        - "PDF" （PDFファイルで出力）
        - "Word" （Wordファイルで出力）

        応答内容:
        {response_content}
        """
        format_response = openai.Completion.create(
            engine=deployment_name,
            prompt=format_prompt,
            max_tokens=50,
            temperature=0.3
        )
        output_decision = format_response['choices'][0]['text'].strip().lower()
        return response_content, output_decision
    except Exception as e:
        return f"ChatGPT 呼び出し中にエラーが発生しました: {str(e)}", "text"

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
