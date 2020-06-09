import os
import sys
import xlwt
import xlrd
import optparse
from xlutils.copy import copy


def read_from_xlsx(xls_file_name, sheet_name_in_xls, header="T"):
    # 读取xlsx
    global data_sh
    data_info = xlrd.open_workbook(xls_file_name)
    try:
        data_sh = data_info.sheet_by_name(sheet_name_in_xls)
    except Exception as e:
        print("No sheet in %s named %s, %s" % (xls_file_name, sheet_name_in_xls, e))

    data_nrows = data_sh.nrows
    data_ncols = data_sh.ncols

    result_data = []
    if header == "T":
        for i in range(3, data_nrows):  # 表头为第3行
            result_data.append(dict(zip(data_sh.row_values(2), data_sh.row_values(i))))
    elif header == "F":
        for i in range(data_nrows):
            result_data.append(data_sh.row_values(i))
    else:
        print("header = ", header)
        print("The parameter 'header' is undefined, please check!")
        os._exit()

    return result_data


def info_to_xlsx(list_head_names, list_info, output_file_name, sheet_name):
    # 写入xlsx
    if list_head_names == '':
        print("Please set the list of head names of the output .xls file.")
        print("If you don't set the names, default value is empty.")
    if not list_info:
        print("The content to be filled is empty, ")
        print("please check the input parameters of function info_to_xlsx")
        return
        # os._exit()
    if output_file_name == '':
        print("Please assign the names of the output xls file's name.")
        os._exit()

    # 若无汇总表则新建，有则追加
    if os.path.exists(output_file_name):
        exist_wb = xlrd.open_workbook(output_file_name)
        try:
            ws = exist_wb.sheet_by_name(sheet_name)
        except Exception as e:
            print("No sheet in %s named %s, %s" % (output_file_name, sheet_name, e))
            ws = exist_wb.add_sheet(sheet_name)
        n_rows = ws.nrows
        wb = copy(exist_wb)
        ws = wb.get_sheet(sheet_name)
    else:
        print("Summary file not existing, now creating at %s" % output_file_name)
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet(sheet_name)
        n_rows = 0

    if isinstance(list_info[0], list):
        if list_head_names != '':
            for i in range(len(list_head_names)):
                ws.write(0, i, list_head_names[i])
            for i in range(len(list_info)):
                for j in range(len(list_info[i])):
                    ws.write((n_rows + i + 1), j, (list_info[i])[j])
        else:
            for i in range(len(list_info)):
                for j in range(len(list_info[i])):
                    ws.write(n_rows + i + 1, j, (list_info[i])[j])
    elif isinstance(list_info[0], str):
        if list_head_names != '':
            for i in range(len(list_head_names)):
                ws.write(0, i, list_head_names[i])
            for i in range(len(list_info)):
                ws.write(n_rows, i, list_info[i])
        else:
            for i in range(len(list_info)):
                ws.write(n_rows, i, list_info[i])
    wb.save(output_file_name)


def read_file(path):
    try:
        with open(path, 'r') as file:
            rows = file.read().splitlines()
            return rows
    except Exception as e:
        print("Read from: %s, ERROR: %s" % (path, e))


def append_file(path, data):
    try:
        with open(path, 'a+') as file:
            for row in data:
                file.write(row + '\n')
    except Exception as e:
        print("Write to: %s, ERROR: %s" % (path, e))


def write_file(path, data):
    try:
        with open(path, 'w') as file:
            for row in data:
                file.write(row + '\n')
    except Exception as e:
        print("Write to: %s, ERROR: %s" % (path, e))


def quality_control(table):
    # 质控：修正大小写
    table = [i for i in table if i['检测编号'] != '' and i['样本姓名'] != '']
    for sample in table:
        if sample['检测编号'] != '':
            sample['检测编号'] = str(sample['检测编号']).upper()
    return table


if __name__ == '__main__':
    parser = optparse.OptionParser()
    parser.add_option('-i', '--input', dest='xlsx_path', default='')
    parser.add_option('-o', '--output', dest='output', default='./summary.xlsx')
    parser.add_option('-l', '--list', dest='list_path', default='./update.list.txt')
    parser.add_option('-g', '--log', dest='log', default='./updated.log')
    (options, args) = parser.parse_args()

    # 读取路径，默认./update.list.txt
    update_list = []
    error_list = []
    if options.xlsx_path != '' and os.path.exists(options.xlsx_path):
        update_list.append(options.xlsx_path)
    elif options.xlsx_path != '':
        print('%s not exist, check please.' % options.xlsx_path)
    for xlsx_path in read_file(options.list_path):
        if xlsx_path != '' and os.path.exists(xlsx_path):
            update_list.append(xlsx_path)
        elif xlsx_path != '':
            print('%s not exist, check please' % xlsx_path)
            error_list.append(xlsx_path)
    if not update_list:
        print('No valid xlsx found')
        os._exit()

    # 更新数据库，并将路径记录到log
    if not os.path.exists(options.log):
        write_file(options.log, '')
    for update_path in update_list:
        if update_path not in read_file(options.log):
            xlsx_table = read_from_xlsx(update_path, 'Sheet1', header='T')
            xlsx_table = quality_control(xlsx_table)
            info_to_xlsx(list(xlsx_table[0].keys()), [list(i.values()) for i in xlsx_table],
                         options.output, 'Sheet1')
            append_file(options.log, [update_path])
        else:
            print('%s had already been updated' % update_path)

    # 未成功更新的路径覆盖输入文件，无则清空
    write_file(options.list_path, error_list)
