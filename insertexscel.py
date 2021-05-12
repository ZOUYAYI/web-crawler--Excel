import xlwt
import xlrd
from xlutils.copy import copy
# 创建一个workbook 设置编码
# workbook = xlrd.open_workbook("Excel_test.xls")
# workbook = xlwt.Workbook(encoding = 'utf-8')
# 创建一个worksheet
# worksheet = workbook.add_sheet('My Worksheet')

# 写入excel
# 参数对应 行, 列, 值
# worksheet.write(1,2, label = 'this is test3')
#
# # 保存
# workbook.save('Excel_test.xls')

def append_to_excel(words, filename):
    '''
    追加数据到excel
    :param words: 【item】 [{},{}]格式
    :param filename: 文件名
    :return:
    '''
    try:
        # 打开excel
        word_book = xlrd.open_workbook(filename)
        # 获取所有的sheet表单。
        sheets = word_book.sheet_names()
        print('sheets:',sheets)
        # 获取第一个表单
        work_sheet = word_book.sheet_by_name(sheets[0])
        print('work_sheet:', work_sheet)
        # 获取已经写入的行数
        old_rows = work_sheet.nrows
        print('old_rows',old_rows)
        # 获取表头信息
        heads = work_sheet.row_values(0)
        print('heads:', heads)
        # 将xlrd对象变成xlwt
        new_work_book = copy(word_book)
        print("copy ok")
        # 添加内容
        new_sheet = new_work_book.get_sheet(0)
        print('new_sheet',new_sheet)
        i = old_rows
        for item in words:
            for j in range(len(heads)):
                new_sheet.write(i, j, item[heads[j]])
            i += 1
        print("ok")
        new_work_book.save(filename)
        print('追加成功！')
    except Exception as e:
        print('追加失败！', e)

if __name__ == '__main__':
    # 样例
    words1 = [
        {'name': 'aki', 'age': 18, 'gender': '女'},
        {'name': 'zed', 'age': 20, 'gender': '男'}
    ]

    words2 = [
        {'name': 'leblance', 'age': 19, 'gender': '女'},
        {'name': 'yasuo', 'age': 20, 'gender': '男'}
    ]
    # 写入内容
    # write_to_excel(words=words1, filename='demo.xls', )
    # 追加内容
    append_to_excel(words=words2, filename='Excel_test.xls')