import openpyxl

def fuga(cell_dict):
    if cell_dict["add_column"]: 
        return "\t/************/\n"
    else:
        return "\t{},\n".format(cell_dict["name"])

def generate_cell_dict(cells, keys):
    cell_dict = {key:cell.value for key, cell in zip(keys, cells)}
    cell_dict["NOT NULL"] = "NOT NULL" if cell_dict["NOT NULL"] == "○" else ""
    cell_dict["add_column"] = True if cells[0].font.color.rgb == "FFFF0000" else False
    return cell_dict

def generate_columns_define(cell_dict):
    return "\t{0} {1} ({2}) {3},\n".format(cell_dict["name"], cell_dict["type"], cell_dict["range"], cell_dict["NOT NULL"]) 
def generate_new_table_SQL(table_name, colmuns_define, primary_keys):
    sql_query = "create table {} (\n".format(table_name)
    sql_query += "".join(colmuns_define)
    sql_query += ")\nprimary_key(\n"
    sql_query += "".join(primary_keys)
    sql_query += ")\n\n"
    return sql_query

def generate_change_table_SQL(table_name, colmuns_define, primary_keys, table_define_list):
    table_name_bk = table_name + "_BK"
    sql_query = "EXEC sp_rename '{0}', '{1}', 'OBJECT' \n\n".format(table_name, table_name_bk)
    sql_query += generate_new_table_SQL(table_name, colmuns_define, primary_keys)
    sql_query += "insert into {} (\n".format(table_name)
    s = ["\t{},\n".format(i["name"]) for i in table_define_list]
    sql_query += "".join(s)
    sql_query += ")\nselect\n"
    sql_query += "".join([fuga(i) for i in table_define_list])
    sql_query += "from \n\t{0}\n\ndrop table {0}".format(table_name_bk)
    return sql_query

def main(ws, file_name):
    ##table_name = ws[""].value
    table_name = "hoge"

    keys = list(ws.values)[10]
    row_list = [i for i in list(ws.rows)[11:] if i[0].value is not None]
    table_define_list = [generate_cell_dict(cells, keys) for cells in row_list]

    colmuns_define = [generate_columns_define(i) for i in table_define_list]
    primary_keys = ["\t{0},\n".format(i["name"]) for i in table_define_list if i["primary"] == "○"]

    with open(file_name, "w") as sql_file:
        if len([i for i in table_define_list if i["add_column"]]) > 0:
            sql_file.write(generate_change_table_SQL(table_name, colmuns_define, primary_keys, table_define_list))
        else:
            sql_file.write(generate_new_table_SQL(table_name, colmuns_define, primary_keys))

if __name__ == "__main__":

    filename = ""
    wb = openpyxl.load_workbook(filename)
    ws = wb[""]

    main(ws, "")
