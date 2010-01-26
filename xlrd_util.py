from xlrd import open_workbook

def load_file(filename=self._filename):
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
    self._workbook = workbook


class CellBlock(object):
    """Represents a block of cells in an Excel Worksheet."""
    
    cells = []
    
    def __init__(self, workbook, sheet_id):
        
        sheet = workbook.sheet_by_idx(sheet_id)
        
        get_cell_attributes(workbook, sheet)
        mark_merged_cells(sheet.merged_cells)
    
    def get_cell_attributes(workbook, sheet):
        
        self.cells = []

        xf_list = workbook.xf_list
        font_list = workbook.font_list
        format_map = workbook.format_map
        
        for row_index in range(sheet.nrows):
            cells.append([None]*sheet.ncols)
            for col_index in range(sheet.ncols):
            
                c = sheet.cell(row_index, col_index)
                
                cell_value = c.value
                cell_value_str = ('%s' % cell_value).encode('utf8')
                cell_type = CELL_TYPE_LOOKUP[c.ctype]
                
                cell_xf_idx = c.xf_index
                cell_xf = xf_list[cell_xf_idx]
                
                cell_font_idx = cell_xf.font_index
                cell_font = font_list[cell_font_idx]
                
                cell_format_idx = cell_xf.format_key
                cell_format = format_map[cell_format_idx]
                cell_format_str = ('%s' % cell_format.format_str).encode('utf8')
                
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
    
    def mark_merged_cells(merged_cells):
        """ Mark merged cells, copy value from top-left cell to 
        others in merged area
        """
        for crange in merged_cells:
            rlo, rhi, clo, chi = crange
            for rowx in xrange(rlo, rhi):
                for colx in xrange(clo, chi):
                    cells[rowx][colx]['merged'] = 1
                    cells[rowx][colx]['value'] = cells[rlo][clo]['value']
            cells[rlo][clo]['merged'] = (rhi-rlo, chi-clo)

        
