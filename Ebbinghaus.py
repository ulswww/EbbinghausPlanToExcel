from datetime import datetime, timedelta
import xlwt
import argvsparser as parser


def parseDays(value, type):
    if type == "M" or type == "H":
        return 0
    if type == "d":
        return value
    if type == "m":
        return value * 30
    pass


class Ebbinghaus(object):
    def __init__(self, cycles):
        self.cycles = cycles

    def get_date_plan(self, day, studyCount, date, use_empty_date):

        lastDay = studyCount

        studiesForPlan = []

        studiesForPlan.append(day)

        if not use_empty_date:
            studiesForPlan.append(date.strftime("%Y-%m-%d"))
        else:
            studiesForPlan.append('    年  月  日')

        studiesForPlan.append('')

        cycleCount = len(self.cycles)

        for cycle in self.cycles:
            dayOffset = -parseDays(cycle. value, cycle.type)
            recallDay = day + dayOffset

            if recallDay >= 1 and recallDay <= lastDay:
                studiesForPlan.append(recallDay)
            else:
                cycleCount -= 1
                studiesForPlan.append("-")

        return (studiesForPlan, cycleCount == 0)

    def get_plans(self, startDate, studyCount, startNo=1, isToEnd=False,
                  use_empty_date=False,
                  use_holiday=False):

        plans = []
        lastCycle = self.cycles[-1]
        lastDays = parseDays(lastCycle.value, lastCycle.type) if isToEnd else 0

        if use_empty_date:
            lastDays = 0
        # startNo = 1
        date = startDate
        empty = []
        empty.append('')
        empty.append(date.strftime("休息"))
        empty.append("")

        for cycle in self.cycles:
            empty.append('')

        for i in range(studyCount+lastDays):

            if date.weekday() == 6 and not use_empty_date and use_holiday:
                date = date + timedelta(days=1)
                plans.append(empty)
            plan, isend = self.get_date_plan(i + 1 + startNo, studyCount, date,
                                             use_empty_date)
            date = date + timedelta(days=1)
            plans.append(plan)

        return plans


class Stage(object):
    def __init__(self, value, type):
        super().__init__()
        self.value = value
        self.type = type

    def __str__(self):
        return "(%d,%s)" % (self.value, self.type)

    __repr__ = __str__


def to_excel(plans, file_name):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('工作表1')
    startOffset = 4
    borders = xlwt.Borders()
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 0
    borders.right = 1
    borders.top = 0
    borders.bottom = 1

    style = xlwt.easyxf('align: vert centre, horiz center;')
    # font: height 350, bold on;
    style.borders = borders

    title_style = xlwt.easyxf('''font: height 350, bold on; align: vert centre,
                               horiz center;''')
    title_style.borders = borders
    name_style = xlwt.easyxf('''font: height 200, bold on; align: vert centre,
                              horiz left;''')
    name_style.borders = borders
    worksheet.row(0).height = 800
    worksheet.row(1).height = 400

    worksheet.col(1).width = 5800
    worksheet.col(2).width = 10000

    worksheet.write_merge(0, 0, 0, 13, '艾宾浩斯遗忘曲线复习计划表', title_style)
    worksheet.write_merge(1, 1, 0, 13, '项目:', name_style)
    worksheet.write_merge(2, 3, 0, 0, '序号', style)
    worksheet.write_merge(2, 3, 1, 1, '学习日期', style)
    worksheet.write_merge(2, 3, 2, 2, '学   习   内   容', style)
    worksheet.write_merge(2, 2, 3, 5, '短期记忆复习周期', style)
    worksheet.write_merge(2, 2, 6, 13, '长期记忆复习周期（复习后打钩）', style)
    worksheet.write(3, 3, '5  分钟', style)
    worksheet.write(3, 4, '30  分钟', style)
    worksheet.write(3, 5, '12  小时', style)
    worksheet.write(3, 6, '1天', style)
    worksheet.write(3, 7, '2天', style)
    worksheet.write(3, 8, '4天', style)
    worksheet.write(3, 9, '7天', style)
    worksheet.write(3, 10, '15天', style)
    worksheet.write(3, 11, '1  个月', style)
    worksheet.write(3, 12, '3  个月', style)
    worksheet.write(3, 13, '6  个月', style)

    worksheet.set_panes_frozen(True)
    worksheet.set_horz_split_pos(4)
    # worksheet.set_vert_split_pos(1)

    for i, plan in enumerate(plans):
        for k, planContent in enumerate(plan):
            worksheet.write(i + startOffset, k, planContent, style)

    workbook.save(file_name)


if __name__ == "__main__":

    ebbinghaus = Ebbinghaus([
                            Stage(5, "M"),
                            Stage(30, "M"),
                            Stage(12, "H"),
                            Stage(1, "d"),
                            Stage(2, "d"),
                            Stage(4, "d"),
                            Stage(7, "d"),
                            Stage(15, "d"),
                            Stage(1, "m"),
                            Stage(3, "m"),
                            Stage(6, "m")])

    d = parser.parse_func('-d', datetime.today())()
    c = parser.parse_func('-c', 365)()
    n = parser.parse_func('-s', 0)()
    file_name = parser.parse_func('-f', 'EbbinghausPlan.xls')()
    e = parser.parse_func('-e', False)()
    use_empty_date = parser.parse_func('--use_empty_date', False, True)()
    use_holiday = parser.parse_func('--use_holiday', False, True)()

    plans = ebbinghaus.get_plans(d, c, startNo=n, isToEnd=e,
                                 use_empty_date=use_empty_date,
                                 use_holiday=use_holiday)

    to_excel(plans, file_name)
