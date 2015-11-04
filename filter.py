from openpyxl import load_workbook, Workbook
import re 
import utils
import sys

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
	
############################################################################
# Open origin
wb_origin = load_workbook(filename = origin_filename)
ws_origin = wb_origin[origin_sheetname]

print('Source file opened')

# Open target
wb_target = Workbook()
ws_target = wb_target.active
ws_target.title = target_sheetname

# Check groups condition in case if exists
if not group_condition:
	selected_groups = None
else: 
	selected_groups = []

	# Group rows 
	groups = dict()

	print('Starting grouping')
	for row in rows_range:
		sys.stdout.write('.')
		group_value = ws_origin['%s%s' % (group_by, row)].value

		if group_value not in groups:
			groups[group_value] = []

		groups[group_value].append(row)

	print('File grouped to %s by column %s.' % (len(groups), group_by))

	# Check group condition
	print('Starting group filtering')
	for key, rows in groups.iteritems():
		sys.stdout.write('.')
		if group_condition(ws_origin, rows):
			selected_groups.append(key)

	print('%s groups found that match the conditions.' % len(selected_groups))

# Copy 
target_row = 1

print('Start copying')
for row in rows_range:
	sys.stdout.write('.')
	check = not row_condition or row_condition(ws_origin, row)
	check = check and (not group_condition or ws_origin['%s%s' % (group_by, row)].value in selected_groups)

	if check:
		for column in columns:
			ws_target['%s%s' % (column, target_row)] = ws_origin['%s%s' % (column, row)].value
		
		target_row += 1


# Save target file
wb_target.save(filename = target_filename)

# Print report and wait input
raw_input("Press Enter to continue...")
