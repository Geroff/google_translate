#!/usr/bin/python

import xlrd
import xlwt
import os
import threading
import time

from translate_google import get_translate

# 英文的列，此处翻译都是基于英文
TRANSLATE_BASE_FIELD = "English"

# 是否使用多线程翻译，多线程容易导致超时，适合翻译少量数据
IS_MULTITHREADING = False

# 是否显示日志
IS_DEBUG = True

# 翻译后数据存放的Excel文件
translate_file = "data/translate_result_fields.xls"

# 待翻译的文件
translate_source_file = 'data/test_fields.xls'

# 需要翻译的Excel文件
book = xlrd.open_workbook(translate_source_file)
# 默认获取第一张表
sheet = book.sheet_by_index(0)

file = xlwt.Workbook(encoding='utf-8')


def print_log(text):
    """
    打印日志
    """
    if IS_DEBUG:
        print(str(text))


def get_english_column():
    """
    获取英文字段的列
    """
    tmp_english_col = 0
    for col in range(1, sheet.ncols):
        field_name = sheet.cell_value(0, col)
        if field_name == TRANSLATE_BASE_FIELD:
            print_log(field_name)
            tmp_english_col = col
            break

    return tmp_english_col


# 需要翻译的语言
translate_dict = {
    # '中文': 'zh-CN',
    # # 挪威语
    # 'Norwegian': 'nb',
    # '中文繁体': 'zh-TW',
    # # 德语
    # 'German': 'de',
    # '韩语': 'ko',
    # 'Japanese': 'ja',
    # '法语': 'fr',
    # '西班牙语': 'es',
    # # 波兰语
    # 'Polski': 'pl',
    # # 意大利语
    # 'Italian': 'it',
    # # 希伯来语
    # 'Hebrew (עברית(': 'iw',
    # # 荷兰语
    # 'Dutch': 'nl',
    # # 印度尼西亚（印尼）
    # 'Indonesian': 'id',
    '捷克语': 'cs',
    # '芬兰语': 'fi',
    #  # 葡萄牙语(巴西),葡萄牙语(葡萄牙)
    # '葡萄牙语': 'pt',
    # '罗马尼亚语': 'ro',
    # '俄语': 'ru',
    # '瑞典语': 'sv',
    # '土耳其语': 'tr',
}
dict_len = len(translate_dict)


class ExcelUtil:
    save_count = 0

    def __init__(self, title, tl, english_col):
        self.tl = tl
        self.title = title
        self.english_col = english_col
        self.translate_sheet = file.add_sheet(tl)
        language_list = []
        language_list.append("Language")
        for row in range(1, sheet.nrows):
            english_field_name = sheet.cell_value(row,  self.english_col)
            if english_field_name is None or english_field_name == "":
                print_log("empty filed!")
                break
            language_list.append(english_field_name)

        self.write_to_sheet(language_list, 0)

    def write_to_sheet(self, field_list, column=0):
        """
        将内容写到Excel的工作表
        :param field_list:
        :param column:
        :return:
        """
        print_log("write_to_sheet")
        row = 0
        for item in field_list:
            print(item)
            self.translate_sheet.write(row, column, item)
            row += 1

    def write_to_excel(self):
        language_list = self.get_translate_list()
        print_log("===========================")
        self.write_to_sheet(language_list, 1)

        if IS_MULTITHREADING:
            ExcelUtil.save_count += 1
            dict_len = len(translate_dict)
            print_log("ExcelUtil.save_count=%d" % ExcelUtil.save_count)
            if dict_len == ExcelUtil.save_count:
                print_log("ExcelUtil.save_count==dict_len")
                file.save(translate_file)

            print_log("===========================")

    def get_translate_list(self):
        print_log("title: %s, col: %s, tl: %s\n" % (self.title, self.english_col, self.tl))
        language_list = []
        language_list.append(self.title)
        translate_text = ""
        total_row = sheet.nrows
        print_log("sheet total row=%d" % total_row)
        row_count = 0

        for row in range(1, sheet.nrows):
            english_field_name = sheet.cell_value(row, self.english_col)
            if '\n' in english_field_name:
                print_log("row==%d, exist \\n" % row)
                # 替换内容中带换行符的，否则google翻译会返回两个结果
                english_field_name = str(english_field_name).replace("\n", "")
            print_log("row=%d field name = %s" % (row, english_field_name))

            next_english_field_name = ""
            if row + 1 < sheet.nrows:
                # 获取下一行内容
                next_english_field_name = sheet.cell_value(row + 1, self.english_col)

            translate_text += (english_field_name + "\n")
            translate_length = len(translate_text)
            row_count += 1
            # 要翻译的内容不能带有英文的句号或者问号，字段中存在多句的需要单独翻译
            if row == total_row - 1 or '.' in next_english_field_name or '?' in next_english_field_name or '.' in english_field_name or '?' in english_field_name or translate_length >= 512:
                print_log("translate_text==" + translate_text)
                self.translate_text(language_list, translate_text, row_count)

                # 重置数据，准备下次翻译
                row_count = 0
                translate_text = ""

        return language_list

    def translate_text(self, language_list, translate_text, row_count):
        translate_result = ""
        this_translate_list = []
        for i in range(0, 5):
            translate_result = get_translate(translate_text, self.tl)
            if len(translate_result):
                try_again = False
                if '.' in translate_text or '?' in translate_text:
                    duplicate_count = 0
                    result_text = ""
                    for results in translate_result:
                        if results is None or results[0] is None:
                            break
                        result_text += results[0]
                        if results[0] in translate_text:
                            duplicate_count += 1

                    # 判断是否是真的进行翻译了
                    if duplicate_count >= 2:
                        try_again = True
                        print_log("include '.' and '？',no translate, result_text==" + result_text + ",duplicate_count==" + str(duplicate_count))

                else:
                    duplicate_count = 0
                    for results in translate_result:
                        if results[0] is None:
                            continue
                        if results[0] in translate_text:
                            duplicate_count += 1

                    print_log("duplicate_count==%d, row_count=%d" % (duplicate_count, row_count))
                    if duplicate_count == row_count:
                        try_again = True
                        print_log("no translate duplicate_count==row_count")

                if not try_again:
                    break

            time.sleep(1)

        if len(translate_result):
            if '.' in translate_text or '?' in translate_text:
                temp_result = ""
                for results in translate_result:
                    if results is None or results[0] is None:
                        break
                    temp_result += results[0]

                temp_result = replace_text(temp_result)
                language_list.append(temp_result)
                print_log(". exist")
                this_translate_list.append("")
            else:
                for results in translate_result:
                    if results[0] is None:
                        continue
                    result = replace_text(results[0])

                    language_list.append(result)
                    this_translate_list.append("")
                print_log(". not exist")
        else:
            # 如果是翻译失败，需要用空格替换对应的行，防止结果对不上
            print_log("result empty! row_count=%d" % row_count)
            for count in range(0, row_count):
                print_log("add empty! count=%d" % count)
                language_list.append(translate_text)
                this_translate_list.append("")

        this_len = len(this_translate_list)

        print_log("this_len==%d, row_count==%d" % (this_len, row_count))

        if this_len != row_count:
            raise Exception("translate count not match!")
        this_translate_list.clear()


def replace_text(text):
    """
    Android资源文件英文的双引号或单引号需要加斜杠，否则会报错，中文的双引号和单引号不需要
    :param text:
    :return:
    """
    temp_text = text.replace(r' \ "', r' \"').replace(r' / ', r'/').replace(r'% ', r' %') \
        .replace(r' $ ', r'$').replace(r'$ ', r'$').replace(r'¥  ', r'¥ ').replace(r'￥  ', r'￥ ').replace(r"'", r"\'").replace(r'"', r'\"').replace("\\\\", "\\")

    return temp_text


def del_file(_translate_file):
    """
    删除文件
    """
    is_exists = os.path.exists(_translate_file)
    if is_exists:
        os.remove(_translate_file)


def start_translate(title, tl, _english_col):
    ExcelUtil(title, tl, _english_col).write_to_excel()


if __name__ == '__main__':
    # 删除文件
    del_file(translate_file)
    english_col = get_english_column()
    print_log("dict size==" + str(len(translate_dict)))
    print_log("english column-->" + str(english_col))
    if IS_MULTITHREADING:
        for key, value in translate_dict.items():
            threading.Thread(target=start_translate, args=(key, value, english_col)).start()
    else:
        for key, value in translate_dict.items():
            start_translate(key, value, english_col)
        file.save(translate_file)

    # print(replace_text("This cellphone number doesn\\'t exsist, please input again"))
