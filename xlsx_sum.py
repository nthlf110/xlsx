import os
import sys
import xlwt
import xlrd
import optparse


def read_from_xlsx(xls_file_name, sheet_name_in_xls, header="T"):
    data_info = xlrd.open_workbook(xls_file_name)
    try:
        data_sh = data_info.sheet_by_name(sheet_name_in_xls)
    except:
        print("no sheet in %s named data" % xls_file_name)

    data_nrows = data_sh.nrows
    data_ncols = data_sh.ncols

    result_data = []
    if header == "T":
        for i in range(3,data_nrows):
            result_data.append(dict(zip(data_sh.row_values(2), data_sh.row_values(i) )))
    elif header == "F":
        for i in range(data_nrows):
            result_data.append(data_sh.row_values(i))
    else:
        print("header = ", header)
        print("The parameter 'header' is undefined, please check!")
        os._exit()

    return result_data


def info_to_xlsx(list_head_names, list_info, output_file_name, sheet_name):
    if list_head_names == '':
        print("Please set the list of head names of the output .xls file.")
        print("If you don't set the names, default value is empty.")
    if list_info == []:
        print("The content to be filled is empty, ")
        print("please check the input parameters of function info_to_xlsx")
        return
        #os._exit()
    if output_file_name == '':
        print("Please assign the names of the output xls file's name.")
        os._exit()

    wb = xlwt.Workbook(encoding='utf-8')
    #wb = xlrd.open_workbook(output_file_name)
    ws = wb.add_sheet(sheet_name)
    if isinstance(list_info[0],(list)) == True:
        if list_head_names != '':
            for i in range(len(list_head_names)):
                ws.write(0, i, list_head_names[i])
            for i in range(len(list_info)):
                for j in range(len(list_info[i])):
                    ws.write((i+1),j,(list_info[i])[j])
        else:
            for i in range(len(list_info)):
                for j in range(len(list_info[i])):
                    ws.write(i,j,(list_info[i])[j])
    elif isinstance(list_info[0],(str)) == True:
        if list_head_names != '':
            for i in range(len(list_head_names)):
                ws.write(0, i, list_head_names[i])
            for i in range(len(list_info)):
                ws.write(1, i, list_info[i])
        else:
            for i in range(len(list_info)):
                ws.write(0, i, list_info[i])
    #output_name = output_file_name + ".xls"
    wb.save(output_file_name)


def read_file(path):
    try:
        with open(path, 'r') as file:
            rows = [row for row in file.readlines()]
            return rows
    except Exception as e:
        print("Read from: %s, ERROR: %s" %(path, e))


def write_file(path, data):
    try:
        with open(path, 'w') as file:
            for row in data:
                file.write(row+'\n')
    except Exception as e:
        print("Write to: %s, ERROR: %s" %(path, e))


if __name__ == '__main__':
    # read update list
    parser = optparse.OptionParser()
    parser.add_option('-i', '--input', dest='xlsx_path')
    parser.add_option('-o', '--output', dest='output', default='./summary.xlsx')
    parser.add_option('-l', '--list', dest='list_path', default='./update.list.txt')
    parser.add_option('-g', '--log', dest='log', default='./output.log')
    (options, args) = parser.parse_args()
    update_list = []
    update_list.append(options.xlsx_path)
    update_list += read_file(options.list_path)

    for update_path in update_list:
        if update_path not in read_file(options.log):
            update_xlsx = read_from_xlsx(update_path, 'sheet1', header='T')