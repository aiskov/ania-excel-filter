import re 
import utils

# Откуда копировать
origin_filename = 'test.xlsx'
origin_sheetname = 'Arkusz1'

# Куда копировать
target_filename = 'result.xlsx'
target_sheetname = 'Sheet1'

def row_condition(ws, row):
	return True # re.compile(r".*(FCU).*").match(ws['B%s' % row].value)
 
# Интервал строк для копирования
rows_range = range(2, 45238)
# Колонки которые мы хотим скопировать
columns = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[0:15]

# Колонка по которой группируем
group_by = 'A'

# Условие выбор группы
def group_condition(ws, rows):
	# Указываем utils.sum_by_group(ws, rows, <>, <>)
	return utils.is_represented(ws, rows, 'O', 'ZIMA 2012')
