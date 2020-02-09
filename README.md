# EbbinghausPlanToExcel
需要 xlwt

# 使用
py Ebbinghaus.py -d 2019-9-9 -c 100 -s 0 -e False
* -d:开始日期，默认当日
* -c:学习数量，默认365
* -s:开始偏移量,默认0
* -e:是否直到学习的最后一日,默认False
* -f:输出文件名，默认./EbbinghausPlan.xls
* --use_empty_date:使用空日期，使得-d参数无效
* --use_holiday:星期日休息

EXCEL样式参考了 http://www.xuexili.com/jiyili/1351.html
