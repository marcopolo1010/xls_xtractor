import re

re_fmt = re.compile('[\$0-9#\.\,]')
re_num = re.compile('[\%#0-9\.\,]')
re_alpha = re.compile('[A-z]')
re_year = re.compile('[1-2][0-9][0-9][0-9]')

def check_for_empty_cell(c):
    if (c['value'] is not None and
        c['value'] != '' and
        c['type'] not in ['BLANK','EMPTY']):
        return 0
    return 1


def check_for_empty_row(cells, row_index):
    for c in cells[row_index]:
        if not check_for_empty_cell(c):
            return 0
    return 1

def check_for_by_line(cells, row_index):
    score = 0.0
    cell_count = 0
    
    for c in cells[row_index]:
        cell_count += 1
        
        if check_for_empty_cell(c):
            score += 1
        elif c['merged']:
            score += 1
        elif cell_count > 1:
            score -= .75
            if (c['type'] == 'TEXT' and len(c['value']) > 20):
                score += 1
            elif (c['style']['bold'] == 1 or c['style']['italic'] == 1):
                score += 1
    return score / cell_count

def compute_row_similarity (schema_block, cells, row1, row2):
    score = 0.0
    col_count = len(schema_block['cols'])
    
    for col in schema_block['cols']:
        
        c1 = cells[row1][col]
        c2 = cells[row2][col]
        
        if check_for_empty_cell(c1) or check_for_empty_cell(c2):
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

        # add .5 if both non-empty
        if not check_for_empty_cell(c1) or not check_for_empty_cell(c2):
            score += 2
        
        # add 2 if both merged, -6 if one is merged
        if c1['merged'] or c2['merged']:
            if c1['merged'] and c2['merged']:
                score += 2
            else:
                score -= 6
    
    return score / col_count / 12; # Normalize it to 1

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

def check_for_header_row(schema_block, cells, row1, row2):
    score = 0.0
    col_count = len(schema_block['cols'])
    
    use_previous = 0
    header_text = {}
    
    for i, col in enumerate(schema_block['cols']):
        prev_col = next_col = -1
        
        if i > 0:
            prev_col = schema_block['cols'][i-1]
        if i + 1 < col_count:
            next_col = schema_block['cols'][i+1]
        
        c1 = cells[row1][col]
        c2 = cells[row2][col]

        # Key ideas: schema should be complete
        # An empty cell should be filled by rows above
        # Duplicate column names should have another column above to differentiate
        
        if check_for_empty_cell(c1):
            score -= 1
            
            # skip if both are empty
            if check_for_empty_cell(c2):
                continue
            
            # if this is empty, hope that the real header is in one of the
            # previous two rows
            elif (row1 - 1 >= 0 and not check_for_empty_cell(cells[row1-1][col])):
                score -= check_for_numeric_cell(cells[row1-1][col]) / 2        
                use_previous = 1
                
            elif (row1 - 2 >= 0 and not check_for_empty_cell(cells[row1-2][col])):
                score -= check_for_numeric_cell(cells[row1-2][col]) / 2      
                use_previous = 2
                
            else:
                score -= 10
            
        # if cell is identical to neighbor, hope that there is another
        # header row above
        elif (row1 - 1 >= 0 and prev_col >= 0 and
              cells[row1][col]['value'] == cells[row1][prev_col]['value']):
            
            if cells[row1-1][col]['value'] != cells[row1-1][prev_col]['value']:
                score -= check_for_numeric_cell(cells[row1-1][prev_col])
                use_previous = 1
                
            elif cells[row1-2][col]['value'] != cells[row1-2][prev_col]['value']:
                score -= check_for_numeric_cell(cells[row1-2][prev_col])
                use_previous = 2
                
            else:
                score -= 10
                
        elif (row1 == 0 and prev_col >= 0 and
              cells[row1][col]['value'] == cells[row1][prev_col]['value']):
            score -= 5
            
        else:
            score += 1
        
        if (c1['format_str'] != c2['format_str'] or
            c1['type'] != c2['type']):
            score += 1
        
        if (check_for_numeric_cell(c1) <=1 and check_for_numeric_cell(c2) > 1):
            score += 2
        
        score -= check_for_numeric_cell(c1)
        score += c1['style']['bold']
        score += c1['style']['italic']
        score += (c1['style']['font'] != c2['style']['font'])
        score += (c1['style']['color_idx'] != c2['style']['color_idx'])

        
    for col in range(len(cells[row1])):
        header_text[col] = ''.encode('utf8')
    
    for col in range(len(cells[row1])):
        if use_previous >= 2:
            header_text[col] += ' %s' % (cells[row1-2][col]['value_str'])
        if use_previous >= 1:
            header_text[col] += ' %s' % (cells[row1-1][col]['value_str'])

        header_text[col] += ' %s' % (cells[row1][col]['value_str'])
    
    score = score / col_count /  3.0; # Normalize it to 1
    
    return (score, header_text)

def score_rows(schema_block, cells):
    empty = {}
    by_line_score = {}
    similarity_score = {}
    header_score = {}
    header_text = {}
    
    num_rows = len(cells)
    
    for row in range(num_rows):
        empty[row] = check_for_empty_row(cells, row)
        by_line_score[row] = check_for_by_line(cells, row)
        header_score[row] = 0
        similarity_score[row] = 0

    for row in range(num_rows - 1):
    
        next_row = row + 1;
        
        if empty[row]:
            continue
        
        if not empty[next_row]:
            sim1 = compute_row_similarity(schema_block, cells, row, next_row)
        else:
            sim1 = -10
            

        if empty[next_row] or by_line_score[next_row] > 0.6:
            next_row += 1
        
        if ((next_row + 1 < num_rows) and
            (empty[next_row] or
             by_line_score[next_row] > 0.6)):
             next_row += 1
        
        if next_row < num_rows:
            (header_score[row], header_text[row]) = check_for_header_row(schema_block, cells, row, next_row)
            sim2 = compute_row_similarity(schema_block, cells, row, next_row)
        else:
            sim2 = -10
        
        similarity_score[row] = max(sim1, sim2)
    
    return (empty, by_line_score, header_score, header_text, similarity_score)

def increment_dict(dict, key):
    if dict.has_key(key):
        dict[key] += 1
    else:
        dict[key] = 1

def calculate_schema_score(row_cells, block):
    score = 0.0
    for i, col in enumerate(block['cols']):
        c = row_cells[col]
        if block['data_types'][i].has_key(c['type']):
            score += block['data_types'][i][c['type']] / len(block['data_rows'])
        
    score = score / len(block['cols'])
    return score
    

def classify_rows(block, scores, cells):
    """Having computed metrics for each row, determine a label for each."""
    (empty, by_line_score, header_score, header_text, similarity_score) = scores

    label = {}
    row_score = {}
    in_block = False
    num_rows = len(empty)
    schema_blocks = []
    
    for row in range(num_rows):
        # classify rows based on header score, similarity score        
        row_score[row] = 0.0
        label[row] = 'UNKNOWN'
        
        if empty[row]:
            label[row] = 'EMPTY'
        
        if not in_block:
            if empty[row]:
                continue
            if (by_line_score[row] < 0.2 and header_score[row] < 0.2):
                label[row] = 'METADATA'
            elif (header_score[row] > 0.5 and 
                  header_score.get(row + 1, 0.0) < header_score[row]):
                label[row] = 'HEADER'
                row_score[row] = header_score[row]
                in_block = True
            else:
                label[row] = 'METADATA'
                row_score[row] = by_line_score[row]
        else:
            schema_score = calculate_schema_score(cells[row], current_block)
            if empty[row] and row + 1 < num_rows and empty[row + 1]:
                in_block = False
            elif header_score[row] > 1.0 and similarity_score[row] < .8:
                row_score[row] = header_score[row]
                label[row] = 'HEADER'
                in_block = False
            elif len(current_block['data_rows']) > 2 and schema_score < .25:
                label[row] = 'METADATA'
                row_score[row] = by_line_score[row]
            elif similarity_score[row] > .4:
                label[row] = 'DATA'
                row_score[row] = similarity_score[row]
            elif (((row + 1 >= num_rows) or empty[row + 1]) and
                  row - 1 > 0 and similarity_score[row-1] > .4):
                label[row] = 'DATA'
            elif len(current_block['data_rows']) > 2 and schema_score > .8:
                label[row] = 'DATA'
        
        # create new schema if we found a new header
        if label[row] == 'HEADER':
            current_block = {}
            current_block['cols'] = block['cols']
            current_block['header_row'] = row
            current_block['data_rows'] = []
            current_block['header_text'] = header_text[row]
            current_block['data_types'] = [{} for c in current_block['cols']]
            schema_blocks.append(current_block)
        elif label[row] == 'DATA':
            row_score[row] = schema_score
            current_block['data_rows'].append(row)
            d = current_block['data_types']
            for i, col in enumerate(current_block['cols']):
                increment_dict(d[i],cells[row][col]['type'])

    return (label, row_score, header_score, schema_blocks)
