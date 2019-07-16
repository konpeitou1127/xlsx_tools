import openpyxl

def create_dict(work_sheet, top_left, bottom_right):
        select_cells = work_sheet[top_left:bottom_right]
        if isinstance(select_cells, tuple) == False:
            assert "error 1"
        elif isinstance(select_cells[0], tuple) == False:
            assert "error 2"
        else:
            assert "error 3"
        
        return {i[0].value:i[1].value for i in select_cells}


def hogehoge(work_sheet):

    # 赤文字のセルを黒文字に変更するメソッド


def fugafuga(work_sheet):
    # 取り消し線のセルを削除するメソッド

    