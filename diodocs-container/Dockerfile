# ベースイメージ
FROM python:3.9-slim

# 作業ディレクトリを設定
WORKDIR /app

# 必要なファイルをコピー
COPY requirements.txt .
COPY diodocs_api.py .

# 必要なライブラリをインストール
RUN pip install --no-cache-dir -r requirements.txt

# FastAPIを起動するコマンド
CMD ["uvicorn", "diodocs_api:app", "--host", "0.0.0.0", "--port", "8000"]
