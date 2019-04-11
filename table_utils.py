#pdf_utils


def get_merged_cells(sheet):
	return sheet.merged_cells

def get_merged_cells_value(sheet, row_index, col_index):
	merged = get_merged_cells(sheet)
	for (rlow, rhigh, clow, chigh) in merged:
		if (row_index >= rlow and row_index < rhigh):
			if (col_index >= clow and col_index < chigh):
				cell_value = str(sheet.cell_value(rlow, clow)).strip()
				return cell_value
	return str(sheet.cell_value(row_index,col_index)).strip()

def get_cellspan(sheet,row_index,col_index):
	cv = get_merged_cells_value(sheet,row_index,col_index)
	cellspan = '<th style="border: 1px solid black;white-space: nowrap;"'
	for (rlow, rhigh, clow, chigh) in get_merged_cells(sheet):
		if (row_index >= rlow and row_index < rhigh):
			if (col_index >= clow and col_index < chigh):
				rowheight = rhigh - rlow
				colheight = chigh - clow
				if (chigh - col_index == colheight) and (rhigh - row_index == rowheight):
					cellspan += f' rowspan="{rowheight}"'
					cellspan += f' colspan="{colheight}"'
				else:
					return ''
				cellspan += f'>{cv}</th>'
				return cellspan

	return f'<th style="border: 1px solid black;white-space: nowrap;">{cv}</th>'
