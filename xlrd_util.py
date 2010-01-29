import re
from xlrd import open_workbook, XLRDError

CELL_TYPE_LOOKUP = ['EMPTY','TEXT','NUMBER','DATE','BOOLEAN','ERROR','BLANK']

re_fmt = re.compile('[\$0-9#\.\,]')
re_num = re.compile('[\%#0-9\.\,]')
re_alpha = re.compile('[A-z]')
re_year = re.compile('[1-2][0-9][0-9][0-9]')

def load_file(filename):
    """Attempt to open the specified file using xlrd.  
    
    Catch and suppress common exceptions, but output a warning.
    """
    # TODO: use real python warnings
    try:
        workbook = open_workbook(filename, formatting_info=True)
    except XLRDError, e:
        if 'Expected BOF' in str(e):
            print "Error reading XLS file (file extension may be wrong):", e
        else:
            raise
        return
    except IOError, e:    
        if 'Permission denied' in str(e):
            print "Error reading XLS file (check file permissions):", e
        else:
            raise
        return
    except EnvironmentError, e:
        if 'Errno 22' in str(e):
            print "Error reading XLS file (file may be empty):", e
        else:
            raise
        return
    except AssertionError:
        print 'Error in XLS parsing library, file could not be opened.'
        return
    return workbook


class CellBlock(object):
    """Represents a block of cells in an Excel Worksheet."""
        
    def __init__(self, workbook, sheet_id):

        self.cells = []
        self.row_list = []
        self.col_list = []
        
        sheet = workbook.sheet_by_index(sheet_id)
        
        self.get_cell_attributes(workbook, sheet)
        self.mark_merged_cells(sheet.merged_cells)
        self.update_cell_lists()
    
    def get_cell_attributes(self, workbook, sheet):
        
        cells = []

        xf_list = workbook.xf_list
        font_list = workbook.font_list
        format_map = workbook.format_map
        format_map_str = {}
        for k in format_map:
            format_map_str[k] = ('%s' % format_map[k].format_str).encode('utf8')
        
        for row_index in range(sheet.nrows):
            cells.append([None]*sheet.ncols)
            for col_index in range(sheet.ncols):
            
                c = sheet.cell(row_index, col_index)
                
                cell_value = c.value
                try:
                    cell_value_str = str(cell_value)
                except:
                    cell_value_str = ('%s' % cell_value).encode('utf8')
                    
                cell_type = CELL_TYPE_LOOKUP[c.ctype]
                cell_xf = xf_list[c.xf_index]
                cell_font = font_list[cell_xf.font_index]
                cell_format_str = format_map_str[cell_xf.format_key]
                
                cell_style = {}
                cell_style['bold'] = cell_font.bold
                cell_style['italic'] = cell_font.italic
                cell_style['color_idx'] = cell_font.colour_index
                cell_style['font'] = cell_font.name
                cell_style['underlined'] = cell_font.underlined
                            
                current_cell = {}
                current_cell['value'] = cell_value
                current_cell['value_str'] = cell_value_str
                current_cell['type'] = cell_type
                current_cell['style'] = cell_style
                current_cell['format_str'] = cell_format_str
                current_cell['merged'] = 0
                
                cells[row_index][col_index] = current_cell
                
        self.cells = cells
    
    def mark_merged_cells(self, merged_cells):
        """ Mark merged cells, copy value from top-left cell to 
        others in merged area
        """
        cells = self.cells
        for crange in merged_cells:
            rlo, rhi, clo, chi = crange
            for rowx in xrange(rlo, rhi):
                for colx in xrange(clo, chi):
                    cells[rowx][colx]['merged'] = 1
                    cells[rowx][colx]['value'] = cells[rlo][clo]['value']
            cells[rlo][clo]['merged'] = (rhi-rlo, chi-clo)
        

    def update_cell_lists(self):
        """Remove empty columns from column list"""
        self.row_list = range(len(self.cells))
        if len(self.row_list) == 0:
            return
        
        for col in range(len(self.cells[0])):
            for row in self.row_list:
                c = self.cells[row][col]
                if not check_for_empty_cell(c):
                    self.col_list.append(col)
                    break

def check_for_empty_cell(c):
    if (c['value_str'] == '' or
        c['type'] in ['BLANK','EMPTY']):
        return 1
    else:
        return 0
        
def check_for_empty_row(cellblock, row_index):
    for col in cellblock.col_list:
        c = cellblock.cells[row_index][col]
        if not check_for_empty_cell(c):
            return 0
    return 1
    
def check_for_numeric_cell(c):
    score = 0.0
    
    if check_for_empty_cell(c):
        return 1.0
    
    if (c['type'] != 'NUMBER' or
        re_year.match(c['value_str'])):
        score -= 1
        
    # add 2 if we see currencies, percentages, or numbers
    if (re_fmt.search(c['format_str']) and
        re_num.search(c['value_str']) and
        not re_alpha.search(c['value_str'])):
        score += 2
    
    # TODO
    # add 10 if the cell above and cell below are both numbers
    # and the three are in a sequence
    
    return score
