#!/usr/bin/env python
import traceback, optparse
from xlrd import open_workbook, cellname, XLRDError
import xls_scores, xls_output

CELL_TYPE_LOOKUP = ['EMPTY','TEXT','NUMBER','DATE','BOOLEAN','ERROR','BLANK']
TEST_SPREADSHEET = ['./repo/test.xls']

filename_list = None

def main():
    init_settings()
    
    if len(filename_list) == 0:
        files = [TEST_SPREADSHEET]
    else:
        files = filename_list

    for filename in files:
        xls_output.reset_output()
#        print '*'*20
        print 'Parsing file: %s' % filename
        xls_output.redirect_output_to(filename, new_ext="html")
        parse_spreadsheet_schema(filename)
            #xls_output.reset_output()
            #print "Unexpected error:", traceback.print_exc()


def parse_spreadsheet_schema(filename):
    """Given a file, parse each worksheet.  Print and output the schema for 
    each block.
    """
    workbook = open_xls(filename)

    if workbook is None:
        return
    
    xls_output.print_header(filename, workbook.nsheets)

    for i, sheet in enumerate(workbook.sheets()):
        xls_output.print_sheet_header(sheet)
        
        # parse spreadsheet, save value and attributes for each cell.
        #
        # cell attributes include: type, font style, format string, and whether
        # the cell is merged with others
        
        cells = expand_sheet(workbook, i)
        
        block = get_blocks(cells)
        
        # compute score for each row.
        #
        # scores represent estimates of the probability that a row is a header
        # and a "similarity score" that measures how similar the row is to the
        # next row.
        
        scores = xls_scores.score_rows(block, cells)
        
        # use scores to classify rows.
        # 
        # rows are labeled: HEADER, DATA, METADATA, or EMPTY
        
        (label, row_score, header_score, schema_blocks) = xls_scores.classify_rows(block, scores, cells)

        xls_output.print_cells(sheet, cells, label, row_score, header_score, schema_blocks)

    xls_output.print_footer()

def open_xls(filename):

    try:
        workbook = open_workbook(filename, formatting_info=True)
    except XLRDError, e:
        xls_output.reset_output()
        if 'Expected BOF' in str(e):
            print "Error reading XLS file (file extension may be wrong):", e
        else:
            raise
        return
    except IOError, e:
        xls_output.reset_output()    
        if 'Permission denied' in str(e):
            print "Error reading XLS file (check file permissions):", e
        else:
            raise
        return
    except EnvironmentError, e:
        xls_output.reset_output()
        if 'Errno 22' in str(e):
            print "Error reading XLS file (file may be empty):", e
        else:
            raise
        return
        
    return workbook


def expand_sheet(workbook, sheet_idx):
    sheet = workbook.sheet_by_index(sheet_idx)
    nrows = sheet.nrows
    xf_list = workbook.xf_list
    font_list = workbook.font_list
    format_map = workbook.format_map
    
    cells = []
    
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
            
    # mark merged cells, copy value from top-left cell to
    # others in merged area
    for crange in sheet.merged_cells:
        rlo, rhi, clo, chi = crange
        for rowx in xrange(rlo, rhi):
            for colx in xrange(clo, chi):
                cells[rowx][colx]['merged'] = 1
                cells[rowx][colx]['value'] = cells[rlo][clo]['value']
        cells[rlo][clo]['merged'] = (rhi-rlo, chi-clo)

    return cells
    

def get_blocks(cells):

    b = {}
    b['rows'] = range(len(cells))
    if len(b['rows']) == 0:
        b['cols'] = []
    else:
        b['cols'] = range(len(cells[0]))
    
    #print b['cols']
    remove_empty_block_cols(b, cells)
    #print b['cols']
    
    return b

def remove_empty_block_cols(block, cells):
    for col in block['cols']:
        empty = True
        for row in block['rows']:
            c = cells[row][col]
            if (not xls_scores.check_for_empty_cell(c) and
                c['merged'] != 1):
                 empty = False
                 break
        if empty:
            block['cols'].remove(col)
            pass
    
def init_settings():
    """Fetch command line options and arguments"""
    
    global filename_list
    p = optparse.OptionParser()
    p.add_option('-v', '--verbose', action='store_true', dest='verbose')
    p.add_option('-w', '--web', action='store_true', dest='show_html')
    p.add_option('-o', '--output-dir', action='store', dest='output_dir')
    p.set_defaults(verbose=False)
    p.set_usage('%prog [options] input_file(s)')
    opts, args = p.parse_args()
    
    xls_output.verbose = opts.verbose
    xls_output.show_html = opts.show_html
    xls_output.output_dir = opts.output_dir
    
    filename_list = args

if __name__ == '__main__':
    main()
