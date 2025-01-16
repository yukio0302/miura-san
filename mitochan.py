import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import io
import openpyxl

# PDF生成関数
def generate_pdf(data):
    PAGE_WIDTH, PAGE_HEIGHT = A4
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    # 店名と基本情報
    c.setFont("Helvetica", 12)
    c.drawString(10 * 2.83, PAGE_HEIGHT - 20 * 2.83, f"店名: {data['店名']}")
    c.drawString(150 * 2.83, PAGE_HEIGHT - 20 * 2.83, f"伝票番号: {data['伝票番号']}")
    c.drawString(10 * 2.83, PAGE_HEIGHT - 35 * 2.83, f"日付: {data['日付']}")

    # 商品一覧
    y_position = PAGE_HEIGHT - 60 * 2.83
    for item in data['商品一覧']:
        c.drawString(20 * 2.83, y_position, item['商品コード'])
        c.drawString(100 * 2.83, y_position, str(item['数量']))
        c.drawString(160 * 2.83, y_position, f"{item['金額']}円")
        y_position -= 10 * 2.83

    # 合計金額
    c.drawString(10 * 2.83, (y_position - 10 * 2.83), f"原価金額合計: {data['原価合計']}円")
    c.drawString(100 * 2.83, (y_position - 10 * 2.83), f"売価金額合計: {data['売価合計']}円")

    c.save()
    buffer.seek(0)
    return buffer

# Streamlitアプリ
st.title("専用伝票印刷アプリ")
st.write("アップロードしたExcelデータからPDFを生成します。")

# ファイルアップロード
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=["xlsx"])

if uploaded_file is not None:
    # Excelファイルを読み込む
    wb = openpyxl.load_workbook(uploaded_file)
    sheet = wb.active

    # データを取得
    data = {
        "店名": sheet["A1"].value,  # セルA1に店名があると仮定
        "伝票番号": sheet["B1"].value,
        "日付": sheet["C1"].value,
        "商品一覧": [],
        "原価合計": sheet["D10"].value,  # 合計金額がD10セルにあると仮定
        "売価合計": sheet["E10"].value,
    }

    # 商品データを取得
    for row in sheet.iter_rows(min_row=2, max_row=6, values_only=True):
        if row[0]:  # 商品コードが空でない場合
            data["商品一覧"].append({
                "商品コード": row[0],
                "数量": row[1],
                "金額": row[2],
            })

    # データのプレビュー
    st.write("アップロードされたデータ:")
    st.json(data)

    # PDF生成ボタン
    if st.button("PDFを生成"):
        pdf_buffer = generate_pdf(data)
        st.download_button(
            label="PDFをダウンロード",
            data=pdf_buffer,
            file_name="伝票.pdf",
            mime="application/pdf"
        )
