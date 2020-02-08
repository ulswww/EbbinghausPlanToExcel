import sys
from datetime import datetime, timedelta
import functools
import xlwt


argv = sys.argv[1:]


def get_arg_value(type):
    argv_len = len(argv)
    index = argv.index(type) if type in argv else -1
    if index < 0 or index >= argv_len - 1:
        return None
    else:
        value = argv[index+1]
        return value


def parse_date(value, default):
    if not value:
        return default
    else:
        return datetime.strptime(value, '%Y-%m-%d')


def parse_int(value, default):
    if not value:
        return default
    else:
        return int(value)


def parse_boolean(value, default):
    if not value:
        return default
    else:
        return bool(value)


def parse_func(type, default):
    valuestr = get_arg_value(type)
    getterfunc = argv_getter.get(type)
    return functools.partial(getterfunc, value=valuestr, default=default)


argv_getter = {'-d': parse_date,
               '-c': parse_int,
               '-s': parse_int,
               '-e': parse_boolean}


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

    def GetStudyPlan(self, day, studyCount, date):

        lastDay = studyCount

        studiesForPlan = []

        studiesForPlan.append(day)

        studiesForPlan.append(date.strftime("%Y-%m-%d"))

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

    def getPlans(self, startDate, studyCount, startNo=1, isToEnd=False):

        plans = []
        lastCycle = self.cycles[-1]
        lastDays = parseDays(lastCycle.value, lastCycle.type) if isToEnd else 0
        startNo -= 1
        date = startDate
        empty = []
        empty.append('')
        empty.append(date.strftime("休息"))
        empty.append("")

        for cycle in self.cycles:
            empty.append('')

        for i in range(studyCount+lastDays):

            if date.weekday() == 6:
                date = date + timedelta(days=1)
                plans.append(empty)
            plan, isend = self.GetStudyPlan(i + 1 + startNo, studyCount, date)
            date = date + timedelta(days=1)
            if not isend:
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
                        Stage(6, "m")
                 ])

    dt = datetime.today()

    d = parse_func('-d', dt)()

    c = parse_func('-c', 10)()

    n = parse_func('-s', 0)()

    e = parse_func('-e', False)()

    plans = ebbinghaus.getPlans(d, c, startNo=n, isToEnd=e)

    # for plan in plans:
    #     print(plan)

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

    workbook.save('EbbinghausPlan.xls')
