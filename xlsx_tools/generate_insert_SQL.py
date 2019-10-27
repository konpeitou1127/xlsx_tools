import openpyxl

def generate_values(cells):
    def hoge(cell):
        if cell.value == "":
            return ""
        elif cell.value == "NULL":
            return "NULL"
        else:
            return "'{}'".format(str(cell.value))

    return ",".join([hoge(cell) for cell in cells]) + ")\n"

def main(ws, file_name):
    #table_name = ws[""].value
    table_name = "hogehoge"
    row_list = [i for i in list(ws.rows) if i[0].font.color.rgb == "FFFF0000"]
    values= ["insert into {} values(".format(table_name) + generate_values(cells) for cells in row_list]

    with open(file_name, "w") as sql_file:
        sql_file.writelines(values)

if __name__ == "__main__":
    filename = ""
    wb = openpyxl.load_workbook(filename)
    ws = wb[""]
    main(ws, "")