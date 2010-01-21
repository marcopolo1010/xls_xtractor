import sys, os

show_html = False
verbose = False
output_dir = None

def reset_output():
    sys.stdout = sys.__stdout__

def redirect_output_to(filename, new_ext):
    """Update stdout to point to specified filename.
    
    If the original filename ended in .xls, that is removed.  Then the filename
    is given a .html extension.
    """
    if output_dir == None:
        return
    
    base_name = os.path.basename(filename) 
    if base_name[-4:] == '.xls':
        base_name = base_name[:-4]
    base_name = '%s.%s' % (base_name, new_ext)
    out_file = os.path.join(output_dir,base_name)
    sys.stdout = open(out_file, 'w')

def print_header(filename, num_sheets):
    if verbose: 
        print 'parsing %s' % filename
    elif show_html:
        print '<h1>%s (%d worksheets)</h1>' % (filename, num_sheets)
    
    if show_html:
        print_html_header()

def print_html_header():
    print """
<html>
<head>
    <title>Show</title>
    <style type="text/css">
        caption {font-size:x-large;}
        .METADATA, .SUBHEADER, .HEADER {text-align:center; font-weight:bold;}
        .METADATA {background-color:#fbb;}
        .SUBHEADER {background-color:#bfb;}
        .HEADER {background-color:#dfd;}
        .DATA {background-color:#ddf;}
        .GEOG {background-color:#ffb;}
        .ROWNUM {font-style:italic;}
    </style>
</head>
<body>"""

def print_sheet_header(sheet):
    header_line = "%s (%d x %d)" % (sheet.name, sheet.nrows, sheet.ncols)
    if show_html:
        print '<h2>%s</h2>' % header_line
    else:
        print header_line

def print_cells(sheet, cells, label, row_score, header_score, schema_blocks):
    if verbose:
        for row in range(sheet.nrows):
            if row < 30:
                #for col_index, c in enumerate(cells[row]):
                #    print cellname(row, col_index), '-',
                #    print c['type'], '\t-',
                #    print c['value'], '-',
                #    print c['format_str'], '\t-',
                #    print c['style'], '\t-',
                #    print c['merged']
                print '%-8s - %.2f (%.2f)\t\t' % (label[row], 
                                              row_score[row], 
                                              header_score[row]),
                for c in cells[row]:
                    print '%s\t' % c['value_str'],
                print
                #print 'similar to next row(s): %s' % similar.get(row,0)
        print
        
    if show_html:
        print '<table><tr><td style="vertical-align:top;">Annotated:<br />'
        print_html_simple(cells, label, row_score, header_score)
        print '</td><td style="vertical-align:top;">Schema:<br />'
        print_html(cells, label, row_score, header_score, schema_blocks)
        print '</td></tr></table>'


def print_html_simple(cells, label, row_score, header_score):
    print '<table>'
    for row in xrange(len(cells)):
        cell_type = 'td'
        if label[row] == 'HEADER':
            #print '<thead>'
            cell_type = 'th'
        print '<tr class="%s">' % label[row],
        print '<%s>' % cell_type
        print '%s - %.2f (%.2f)' % (label[row], row_score[row], header_score[row])
        print '</%s>' % cell_type
        for cell in cells[row]:
            if cell['merged'] == 1:
                continue
            elif cell['merged']:
                print '<%s rowspan="%d" colspan="%d">' % (cell_type, cell['merged'][0], cell['merged'][1]),
            else:
                print '<%s>' % cell_type,                
            print cell['value_str'],
            print '</%s>' % cell_type,
        print '</tr>'
    print '</table>'

def print_html(cells, label, row_score, header_score, schema_blocks):

    for b in schema_blocks:
    
        print '<table>'
        
        print '<tr class="HEADER">',
        for i, col in enumerate(b['cols']):
            print '<th>%s</th>' % b['header_text'][i],
        print '</tr>'
        
        print '<tr class="HEADER">',
        for i, col in enumerate(b['cols']):
            print '<th>%s</th>' % b['data_types'][i],
        print '</tr>'
        
        for row in b['data_rows']:
            print '<tr class="DATA">',
            for col in b['cols']:
                print '<td>%s</td>' % cells[row][col]['value_str'],
            print '</tr>'
            
        print '</table>'


def print_footer():
    if show_html:
        print '</body></html>'

