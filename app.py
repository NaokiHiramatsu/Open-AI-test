import os
import openai
from flask import Flask, request, jsonify

app = Flask(__name__)

# Azure OpenAI APIの設定
openai.api_type = "azure"
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_key = os.getenv("AZURE_OPENAI_API_KEY")
openai.api_version = "2023-03-15-preview"

@app.route('/ask', methods=['POST'])
def ask_openai():
    try:
        prompt = request.json.get('prompt')
        response = openai.Completion.create(
            engine="gpt-35-turbo",
            prompt=prompt,
            max_tokens=100
        )
        return jsonify(response.choices[0].text.strip())
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
