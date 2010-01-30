from xls_xtractor.xlrd_util import (check_for_empty_row,
                                    check_for_numeric_cell)

#***************************************
#
#
def compute_row_scores(sheetinfo):
    """Computes "probability" that each row is byline, header, or data."""
    
    cellblock = sheetinfo['cellblock']
    cells = cellblock.cells
    row_list = cellblock.row_list
            
    empty = {}
    metadata_scores = {}
    header_scores = {}
    similarity_scores = {}
    header_text = {}
    
    # initialize scores
    for row in cellblock.row_list:
        empty[row] = check_for_empty_row(cellblock, row)
        try:
            metadata_scores[row] = calc_metadata_score(cellblock, row)
        except:
            print row, cellblock.col_list
            raise
        similarity_scores[row] = 0
        header_scores[row] = 0

    for i, row in enumerate(row_list):

        if empty[row]:
            continue

        # find next row in row_list that is not empty or metadata
        next_row = None
        for n in row_list[i+1:]:
            if not empty[n] and metadata_scores[n] < 0.6:
                next_row = n
                break
        if next_row is None:
            continue
            
        gap = (next_row-row)
        if gap > 6:
            continue
        
        # calculate similarity and header scores
        s = calc_sim_score(cellblock, row, next_row)
        h = calc_header_score(cellblock, row, next_row)
        
        # penalize if there's a big gap
        if gap > 2:
            s -= 0.3
        elif gap > 4:
            s -= 0.7
        
        similarity_scores[row] = s
        (header_scores[row], header_text[row]) = h
    
    scores = {}        
    scores['empty'] = empty
    scores['metadata'] = metadata_scores
    scores['similarity'] = similarity_scores
    scores['header'] = header_scores
    
    sheetinfo['scores'] = scores
    
    sheetinfo['header_text'] = header_text

def calc_metadata_score(cellblock, row):
    score = 0.0
    
    for i, col in enumerate(cellblock.col_list):
        c = cellblock.cells[row][col]
        
        if c['empty']:
            score += 1
        elif c['merged']:
            score += 1
        elif i > 1:
            score -= .75
            if (c['type'] == 'TEXT' and len(c['value']) > 20):
                score += 1
            elif (c['style']['bold'] == 1 or c['style']['italic'] == 1):
                score += 1

    cell_count = len(cellblock.col_list)
    
    if cell_count == 0:
        return 0
                
    return score / cell_count
   
def calc_sim_score(cellblock, row, next_row):

    score = 0.0

    for col in cellblock.col_list:
        c1 = cellblock.cells[row][col]
        c2 = cellblock.cells[next_row][col]
        
        if c1['empty'] or c2['empty']:
            score += 2
        else:
            score -= .25

        # add 4 if numeric formatting is the same
        if c1['format_str'] == c2['format_str']:
            score += 4
        
        # add 2 if same type, o.w. subtract 4
        if c1['type'] == c2['type']:
            score += 2
        else:
            score -= 4
            
        # add 2 if same font, italics or color
        if c1['style']['font'] == c2['style']['font']:
            score += 2
        if c1['style']['italic'] == c2['style']['italic']:
            score += 2
        if c1['style']['color_idx'] == c2['style']['color_idx']:
            score += 2

        # add .5 if both are non-empty
        if not c1['empty'] or not c2['empty']:
            score += 2
        
        # add 2 if both merged, -6 if one is merged
        if c1['merged'] or c2['merged']:
            if c1['merged'] and c2['merged']:
                score += 2
            else:
                score -= 6
    
    col_count = len(cellblock.col_list)    
    return score / col_count / 12;

def calc_header_score(cellblock, row, next_row):

    score = 0.0
    col_count = len(cellblock.col_list)
    
    cells = cellblock.cells
    
    use_previous = 0
    header_text = {}
    
    for i, col in enumerate(cellblock.col_list):
        prev_col = next_col = -1
        
        if i > 0:
            prev_col = cellblock.col_list[i-1]
        if i + 1 < col_count:
            next_col = cellblock.col_list[i+1]
        
        c1 = cellblock.cells[row][col]
        c2 = cellblock.cells[next_row][col]

        # Key ideas: schema should be complete
        # An empty cell should be filled by rows above
        # Duplicate column names should have another column above to differentiate
        
        if c1['empty']:
            score -= 1
            
            # skip if both are empty
            if c2['empty']:
                continue
            
            # if this is empty, hope that the real header is in one of the
            # previous two rows
            elif (row - 1 >= 0 and not cells[row-1][col]['empty']):
                score -= check_for_numeric_cell(cells[row-1][col]) / 2        
                use_previous = 1
                
            elif (row - 2 >= 0 and not cells[row-2][col]['empty']):
                score -= check_for_numeric_cell(cells[row-2][col]) / 2      
                use_previous = 2
                
            else:
                score -= 4
            
        # if cell is identical to neighbor, hope that there is another
        # header row above
        elif (row - 1 >= 0 and prev_col >= 0 and
              cells[row][col]['value'] == cells[row][prev_col]['value']):
            
            if cells[row-1][col]['value'] != cells[row-1][prev_col]['value']:
                score -= check_for_numeric_cell(cells[row-1][prev_col])
                use_previous = 1
                
            elif cells[row-2][col]['value'] != cells[row-2][prev_col]['value']:
                score -= check_for_numeric_cell(cells[row-2][prev_col])
                use_previous = 2
                
            else:
                score -= 4
                
        elif (row == 0 and prev_col >= 0 and
              cells[row][col]['value'] == cells[row][prev_col]['value']):
            score -= 5
            
        else:
            score += 1
        
        if (c1['format_str'] != c2['format_str'] or
            c1['type'] != c2['type']):
            score += 1
        
        if c1['style']['bold'] != c2['style']['bold']:
            score += 1
        
        if (check_for_numeric_cell(c1) <=1 and check_for_numeric_cell(c2) > 1):
            score += 2
        
        score -= check_for_numeric_cell(c1)
        score += c1['style']['bold']
        score += c1['style']['italic']
        score += (c1['style']['font'] != c2['style']['font'])
        score += (c1['style']['color_idx'] != c2['style']['color_idx'])

        
    for col in range(len(cellblock.cells[row])):
        header_text[col] = ''.encode('utf8')
    
    for col in range(len(cellblock.cells[row])):
        if use_previous >= 2:
            header_text[col] += ' %s' % (cellblock.cells[row-2][col]['value_str'])
        if use_previous >= 1:
            header_text[col] += ' %s' % (cellblock.cells[row-1][col]['value_str'])

        header_text[col] += ' %s' % (cellblock.cells[row][col]['value_str'])
    
    score = score / col_count /  3.0; # Normalize it to 1
    
    return (score, header_text)

#**********************************************************
# The next two functions classify the rows as either:
#   - HEADER
#   - METADATA
#   - DATA
#   - UNKNOWN
#**********************************************************
def dumb_classify(sheetinfo):
    """Having computed metrics for each row, determine a label for each."""

    scores = sheetinfo['scores']
    empty = scores['empty']
    metadata_scores = scores['metadata']
    header_scores = scores['header']
    similarity_scores = scores['similarity']
    header_text = sheetinfo['header_text']

    labels = {}
    row_scores = {}
    
    in_block = False
    schema_blocks = []
    
    for row in sheetinfo['cellblock'].row_list:
        # classify rows based on header score, similarity score        
        row_scores[row] = 0.0
        labels[row] = 'UNKNOWN'
        
        if empty[row]:
            labels[row] = 'EMPTY'
        
        if not in_block:
            if empty[row]:
                continue
            if (metadata_scores[row] < 0.2 and header_scores[row] < 0.2):
                labels[row] = 'METADATA'
            elif (header_scores[row] > 0.5 and 
                  header_scores.get(row + 1, 0.0) < header_scores[row]):
                labels[row] = 'HEADER'
                row_scores[row] = header_scores[row]
                in_block = True
            else:
                labels[row] = 'METADATA'
                row_scores[row] = metadata_scores[row]
        else:
            #schema_score = calculate_schema_score(cells[row], current_block)
            if empty[row] and empty.get(row + 1, 0) and empty.get(row + 2, 0):
                in_block = False
            elif header_scores[row] > 1.0 and similarity_scores[row] < .8:
                row_scores[row] = header_scores[row]
                labels[row] = 'HEADER'
#                elif len(current_block['data_rows']) > 2: # and schema_score < .25:
#                    labels[row] = 'METADATA'
#                    row_scores[row] = metadata_scores[row]
            elif similarity_scores[row] > .4:
                labels[row] = 'DATA'
                row_scores[row] = similarity_scores[row]
#                elif (((row + 1 >= num_rows) or empty[row + 1]) and
#                      row - 1 > 0 and similarity_scores[row-1] > .4):
#                    labels[row] = 'DATA'
#                elif len(current_block['data_rows']) > 2 and schema_score > .8:
#                    labels[row] = 'DATA'
        
        # create new schema if we found a new header
#            if labels[row] == 'HEADER':
#                current_block = {}
#                current_block['cols'] = block['cols']
#                current_block['header_row'] = row
#                current_block['data_rows'] = []
#                current_block['header_text'] = header_text[row]
#                current_block['data_types'] = [{} for c in current_block['cols']]
#                schema_blocks.append(current_block)
#            elif labels[row] == 'DATA':
#                row_score[row] = schema_score
#                current_block['data_rows'].append(row)
#                d = current_block['data_types']
#                for i, col in enumerate(current_block['cols']):
#                    increment_dict(d[i],cells[row][col]['type'])
    for row in sheetinfo['cellblock'].row_list:    
        if labels[row] == 'UNKNOWN' and labels.get(row-1,'') == 'DATA':
            labels[row] = 'DATA'
    
    sheetinfo['labels'] = labels
    sheetinfo['scores']['row_scores'] = row_scores

def smart_classify(sheet_info):
    pass
