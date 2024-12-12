from fastapi import FastAPI, UploadFile, File
import io
from diodocs import Workbook

app = FastAPI()

@app.post("/generate-excel/")
async def generate_excel(data: list[dict]):
    """
    JSONデータをExcelファイルに変換するエンドポイント
    """
    try:
        # DioDocsを使ってExcelファイルを作成
        workbook = Workbook()
        sheet = workbook.worksheets.add("Sheet1")
        
        # データ構造からヘッダーと内容を取得
        if not data or not isinstance(data, list):
            return {"status": "error", "message": "データが不正です。"}
        headers = list(data[0].keys())
        sheet.append(headers)  # ヘッダーを追加
        for row in data:
            sheet.append([row.get(header, "") for header in headers])  # 行データを追加

        # 作成したExcelをバイト形式に変換
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        # ファイルを返す
        return {
            "status": "success",
            "message": "Excelファイルが生成されました。",
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}

