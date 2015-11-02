def sum_by_group(ws, rows, group_column, sum_column):
	result = dict()

	for row in rows:
		key = ws['%s%s' % (group_column, row)].value
		value = ws['%s%s' % (sum_column, row)].value

		if key not in result:
			result[key] = value
		else:
			result[key] += value

	return result

def is_represented(ws, rows, group_column, value):
	for row in rows:
		if ws['%s%s' % (group_column, row)].value == value:
			return True

	return False

# def percent_of_group(ws, rows, group_column, sum_column, group)
