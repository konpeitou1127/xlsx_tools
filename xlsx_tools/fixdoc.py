import openpyxl
from openpyxl.styles import Font
from openpyxl.styles import colors
import glob
import pathlib
import shutil

def convert_tab_color(work_sheet):
    if work_sheet.sheet_properties.tabColor is not None:
        if work_sheet.sheet_properties.tabColor.rgb == colors.RED:
            work_sheet.sheet_properties.tabColor = None 
   

def convert_cell_font_color(work_sheet):
    # セルのフォントカラーを赤から黒にする
    # 一括で変更する何かいい方法があるはず
    for i in work_sheet.rows:
        for j in i:
            if j.font.color.rgb == colors.RED:
                j.font.color.rgb = colors.BLACK


def delete_strike_cell(work_sheet):
    # 取り消し線が引いてある行の削除
    for i, x in enumerate(work_sheet.rows):
        strike_status_list = [j.font.strike for j in x]
        if all(strike_status_list):
            work_sheet.delete_rows(i+1)    


# 指定したディレクトリから.xlsxファイルをすべて見つける
target_dir = "./hogehoge"
target_files = pathlib.Path(target_dir).glob("./**/*.xlsx")
output_dir = "./result"

# 出力ファイルを格納するフォルダを作成
if not pathlib.Path(output_dir).exists:
    shutil.copytree(target_dir, output_dir, ignore=shutil.ignore_patterns("*.xlsx"))

for target_file in target_files:
    work_book = openpyxl.load_workbook(target_file.absolute())

    for work_sheet in work_book:

        # タブの色を赤から色なしに変更
        convert_tab_color(work_sheet)

        # セルのフォントカラーを赤から黒にする
        convert_cell_font_color(work_sheet)

        # 取り消し線が引いてある行の削除
        delete_strike_cell(work_sheet)

    # output_filename = str(target_file.absolute()).replace(target_dir, output_dir)
    output_filename = pathlib.Path(output_dir) / target_file.relative_to(target_dir)
    work_book.save(output_filename.absolute())

