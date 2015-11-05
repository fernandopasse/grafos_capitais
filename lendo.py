import xlrd
from graphviz import Graph
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path, formatting_info = True)
    # print number of sheets
    print book.nsheets
    # print sheet names
    print book.sheet_names()
    g = Graph(format='pdf')
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    num_cols = first_sheet.ncols
    for row_idx in range(1, first_sheet.nrows):    # Iterate through rows
    	for col_idx in range(1, num_cols):  # Iterate through columns
    		cel_nome_linha = first_sheet.cell(row_idx, 0)
    		cel_nome_coluna = first_sheet.cell(0, col_idx)
    		cel_valor = first_sheet.cell(row_idx, col_idx)
    		xfx = first_sheet.cell_xf_index(row_idx, col_idx)
    		xf = book.xf_list[xfx]
    		bgx = xf.background.pattern_colour_index
    		if cel_valor.value != '':
        		print('(%s, %s, %s, %s)' % (cel_nome_linha.value, cel_nome_coluna.value, cel_valor.value, bgx))
        		g.edge(cel_nome_linha.value, cel_nome_coluna.value, label = '%s' % (cel_valor.value))
    print(g.source)
    g.render()
    # read a row
#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "dados.xls"
    open_file(path)