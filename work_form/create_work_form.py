from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Pt
import datetime
import calendar
import pandas as pd

day_to_chinese = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五'}


class WeekFreeTime:
    def __init__(self, curriculum):
        self.free_time_list = []
        self.curriculum = curriculum
        for date in range(1, 6):
            column = TableGenerator.convert_day_to_chinese(date)
            free_time = WeekFreeTime.calculate_day_free_time(self.curriculum[column])
            self.free_time_list.append(free_time)

    def get_free_time(self, date):
        return self.free_time_list[date.weekday()]

    @staticmethod
    def calculate_day_free_time(curriculum_col):
        free_time = []
        start = 0
        i = 0
        while i != len(curriculum_col):
            if curriculum_col[i] == 1:
                end = i
                free_time.append([start, end])
                while i != len(curriculum_col) and curriculum_col[i] == 1:
                    i += 1
                start = i
            else:
                i += 1
        return free_time


class TableGenerator:
    def __init__(self, curriculum):
        self.curriculum = curriculum

    def generate_table(self, hour, docx, month):
        table = docx.tables[0]
        nowrow = 1
        week_free_times = WeekFreeTime(self.curriculum)
        for date in TableGenerator.iter_month(month):
            if date.weekday() < 5:  # 是否為平日
                day_free_times = week_free_times.get_free_time(date)
                for start, end in day_free_times:
                    hour_count = end - start
                    if hour < hour_count or hour_count >= 4:
                        if hour < hour_count:
                            end = start + hour
                        hour_record_count = \
                            TableGenerator.add_valid_hour(table, nowrow, date, start, end)
                        hour -= hour_record_count
                        if hour_record_count != 0:
                            nowrow += 1
                            if nowrow == 18:
                                return True
        return hour == 0

    @staticmethod
    def convert_hour_count_to_valid_hour_count(hour_count):
        valid_hour_count_list = [0, 1, 2, 4, 8]
        for i in range(len(valid_hour_count_list)-1, -1, -1):
            if hour_count >= valid_hour_count_list[i]:
                return valid_hour_count_list[i]

    @staticmethod
    def add_valid_hour(table, nowrow, date, start, end):
        hour_count = end - start
        valid_hour_count = TableGenerator.convert_hour_count_to_valid_hour_count(hour_count)
        end = start + valid_hour_count
        if start != end:
            return TableGenerator.add_hour_record(table, nowrow, date, start, end)
        else:
            return 0

    @staticmethod
    def add_hour_record(table, nowrow, date, start, end):
        hour_count = end - start
        if hour_count > 4:
            hour_count -= 1  # 休息的那一小時
        text_list = [
                     f'{date.year - 1911}/{date.month}/{date.day}',
                     TableGenerator.convert_day_to_chinese(date.isoweekday()),
                     TableGenerator.convert_free_time_interval_to_output_format(start, end),
                     str(hour_count),
                     '資料整理',
                    ]
        for i in range(5):
            TableGenerator.table_run_add_and_set(table, nowrow, i, text_list[i])
        return hour_count

    @staticmethod
    def convert_free_time_interval_to_output_format(start, end):
        s = TableGenerator.convert_index_to_curse_start_time(start)
        e = TableGenerator.convert_index_to_curse_start_time(end)
        hour_count = end-start
        output = s + '~' + e
        if hour_count > 4:
            rest_time = TableGenerator.convert_index_to_curse_start_time(start+4) + '-' + \
                        TableGenerator.convert_index_to_curse_start_time(start+5)
            output += '\n(' + rest_time + '休息)'
        return output

    @staticmethod
    def convert_index_to_curse_start_time(idx):
        return f'{idx+8}:00'

    @staticmethod
    def table_run_add_and_set(table, row, col, text):
        run = table.cell(row, col).paragraphs[0].add_run(text)
        table.cell(row, col).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 水平置中
        TableGenerator.process_chinese_setting(run)

    @staticmethod
    def text_run_add_and_set(docx, idx, pre, hour, suf):
        docx.paragraphs[idx].text = ''
        run = docx.paragraphs[idx].add_run(pre)
        TableGenerator.process_run_setting(run)
        run = docx.paragraphs[idx].add_run(str(hour))
        TableGenerator.process_run_setting(run)
        run.underline = True
        run = docx.paragraphs[idx].add_run(suf)
        TableGenerator.process_run_setting(run)

    @staticmethod
    def process_run_setting(run):
        TableGenerator.process_chinese_setting(run)
        run.font.size = Pt(14)
        run.bold = True

    @staticmethod
    def process_chinese_setting(run):
        run.font.name = u'標楷體'
        r = run._element.rPr.rFonts  # 中文特有的處理
        r.set(qn("w:eastAsia"), "標楷體")

    @staticmethod
    def iter_month(month):
        beg_date = datetime.date(datetime.datetime.now().date().year, month, 1)
        end_date = beg_date + datetime.timedelta(calendar.mdays[month])
        while beg_date != end_date:
            yield beg_date
            beg_date += datetime.timedelta(1)

    @staticmethod
    def convert_day_to_chinese(day):
        return day_to_chinese.get(day, 'None')
'''
def convert_day_to_chinese(day):
    return day_to_chinese.get(day, 'None')

def iter_month(month):  # 無考慮閏年 產生那一個月的迭代器
    beg_date = datetime.date(datetime.datetime.now().date().year, month, 1)
    end_date = beg_date + datetime.timedelta(calendar.mdays[month])
    while beg_date != end_date:
        yield beg_date
        beg_date += datetime.timedelta(1)
def process_chinese_setting(run):  # 使run正常顯示中文(標楷體)
    run.font.name = u'標楷體'
    r = run._element.rPr.rFonts  # 中文特有的處理
    r.set(qn("w:eastAsia"), "標楷體")

def table_run_add_and_set(table, row, col, text):  # add run and set some value in table
    run = table.cell(row, col).paragraphs[0].add_run(text)
    table.cell(row, col).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 水平置中
    process_chinese_setting(run)


def text_run_add_and_set(docx, idx, pre, hour, suf):  # add run and set some value in text
    docx.paragraphs[idx].text = ''
    run = docx.paragraphs[idx].add_run(pre)
    process_chinese_setting(run)
    run.font.size = Pt(14)
    run.bold = True
    run = docx.paragraphs[idx].add_run(str(hour))
    process_chinese_setting(run)
    run.font.size = Pt(14)
    run.bold = True
    run.underline = True
    run = docx.paragraphs[idx].add_run(suf)
    process_chinese_setting(run)
    run.font.size = Pt(14)
    run.bold = True


def generate_table(table, hour, days, month):  # hour is integer,days is a list contain integer,month is int
    nowrow = 1
    for date in iter_month(month):
        if date.isoweekday() in days:
            if hour > 7:
                table_run_add_and_set(table, nowrow, 0, f'{date.year - 1911}/{date.month}/{date.day}')
                table_run_add_and_set(table, nowrow, 1, convert_day_to_chinese(date.isoweekday()))
                table_run_add_and_set(table, nowrow, 2, '13:00~21:00\n(17:00-18:00休息)')
                table_run_add_and_set(table, nowrow, 3, str(7))
                table_run_add_and_set(table, nowrow, 4, '資料整理')
            else:
                table_run_add_and_set(table, nowrow, 0, f'{date.year - 1911}/{date.month}/{date.day}')
                table_run_add_and_set(table, nowrow, 1, convert_day_to_chinese(date.isoweekday()))
                if hour > 4:
                    table_run_add_and_set(table, nowrow, 2, '13:00~' + str(13 + hour + 1) + ':00\n(17:00-18:00休息)')
                else:
                    table_run_add_and_set(table, nowrow, 2, '13:00~' + str(13 + hour) + ':00')
                table_run_add_and_set(table, nowrow, 3, str(hour))
                table_run_add_and_set(table, nowrow, 4, '資料整理')
                return True
            hour -= 7
            nowrow += 1
    return False
'''

docx = Document("work.docx")
df = pd.read_excel("work.xlsx", engine='openpyxl')
df = df.append({"一": 1, "二": 1, "三": 1, "四": 1, "五": 1}, ignore_index=True)
table_generator = TableGenerator(df)
'''
table = docx.tables[0]
available_day = [3, 5]  # list(map(int,input("請輸入可行的工作日: ").split()))
'''
hour = 30
month = 12
#hour = int(input("請輸入工時: "))
#month = int(input("請輸入月份: "))
if table_generator.generate_table(hour, docx, month):
    print('產生word完成。')
    TableGenerator.text_run_add_and_set(docx, 2, '*正常工作時數：', f'      {hour}     ', '小時')
    TableGenerator.text_run_add_and_set(docx, 3, '*薪資小計：', f'      {hour*160}      ', '元(A)')
    TableGenerator.text_run_add_and_set(docx, 5, '*合計應領薪資：', f'      {hour*160}     ', '元(A+B) /　工讀生指導人：______________')
    docx.save(str(month) + '月份工作表.docx')
else:
    print('工作日不夠，產生word失敗。')
