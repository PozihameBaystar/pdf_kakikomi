import sys
import os
import io
import yaml
import PyPDF2
from PyPDF2 import PdfMerger, PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import B5
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# 日本語フォントファイルのパスを正しく指定
#font_path = "C:/Windows/Fonts/meiryo.ttc"  # 例: メイリオフォント
#pdfmetrics.registerFont(TTFont("Japanese", font_path))

# 空白のページを作成する関数
def create_blank_pdf():
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=B5)
    can.showPage()
    can.save()
    packet.seek(0)
    return PdfReader(packet)

# 文字を追加する関数
def add_text_to_pdf(input_pdf,text):
    # バイトストリーム（一時メモリ）を作成
    packet = io.BytesIO()

    # 透明なpdfデータをバイトストリーム上に作成
    can = canvas.Canvas(packet,pagesize=B5)
    # フォントを日本語にセット
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    can.setFont("HeiseiKakuGo-W5",12)
    # 透明なpdfデータに文字を書き込む
    can.drawString(B5[0]-20*mm, B5[1]-10*mm, text)

    # 変更をメモリストリームに保存
    can.save()
    # ストリームの読み取りポインタを先頭に戻す
    packet.seek(0)

    # メモリストリームから文字だけが書かれたpdfを読み取る
    watermark = PdfReader(packet)
    # 出力pdfの準備
    output = PdfWriter()

    # 文字が掛かれた透明なpdfと元のpdfを重ねる
    page = input_pdf.pages[0]
    page.merge_page(watermark.pages[0])
    output.add_page(page)

    # PdfWriterオブジェクトをバイトストリームに変換
    output_stream = io.BytesIO()
    output.write(output_stream)
    output_stream.seek(0)

    return output_stream

# pdf同士を結合する関数
def merge_pdfs_with_blank_page(pdf1_path, pdf2_path, output_path, text):
    merger = PdfMerger()

    # 最初のPDF（文字付き）を追加
    pdf1 = PdfReader(pdf1_path)
    pdf1_with_text = add_text_to_pdf(pdf1,text)
    merger.append(pdf1_with_text)

    # 空白ページを追加
    blank_page = create_blank_pdf()
    merger.append(blank_page)

    # 2つ目のPDFを追加
    merger.append(pdf2_path)

    # 出力ファイルに保存
    with open(output_path, "wb") as output_file:
        merger.write(output_file)

    merger.close()


if __name__ == "__main__":

    test_id = "2023本試"

    # yamlファイルを開く
    with open('kakikomi.yaml') as file:
        config = yaml.safe_load(file.read())

    # yamlファイルの設定からフォルダーのパスを取得
    input_folder = config['input_name']
    output_folder = config['output_name']
    op_folder = config['op_name']

    # 組み合わせてパスを作成
    input_path = os.path.join(input_folder,test_id)
    op_path = op_folder
    output_path = os.path.join(output_folder,test_id)

    # 問題冊子（pdfファイル）の一覧を取得
    file_list = [f for f in os.listdir(input_path) if f[-3:] == "pdf"]
    # 表紙の一覧を取得
    op_list = [f for f in os.listdir(op_path)]

    for i in range(len(file_list)):
        f_path = os.path.join(input_path,file_list[i])
        o_path = os.path.join(op_path,op_list[i])
        out_path = os.path.join(output_path,file_list[i])
        merge_pdfs_with_blank_page(o_path,f_path,out_path,test_id)
