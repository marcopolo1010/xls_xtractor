#
#
# 2010-01-25 mda     Copied logic from original project files.
# 2010-01-25 mda     Created file.

import xlrd_util, score_util

class XlsXtractor(object):
    _workbook = None
    _sheetinfo = {}
    
    def __init__(self, filename):
        self._filename = filename
        self._workbook = xlrd_util.load_file(filename)
        if self._workbook:
            self.valid = True
            self.nsheets = self._workbook.nsheets
        else:
            self.valid = False
    
    def extract_schemas(self, sheet_id):
        sheetinfo = {}
        self._sheetinfo[sheet_id] = sheetinfo
        
        sheetinfo['cellblock'] = self._parse_sheet(sheet_id)
        self._compute_row_scores(sheet_id)
        self._dumb_classify(sheet_id)
        self._smart_classify(sheet_id)
        
    def output(self, sheet_id, format='html'):
        if format == 'text':
            output_text(sheet_id)
        else:
            output_html(sheet_id)

    def output_text(self, sheet_id):
        sheetinfo = self._sheetinfo[sheet_id]
        scores = sheetinfo['scores']
        labels = sheetinfo['labels']
        print 'ROW  LABEL       META   SIM  HEADER'
        for row in sheetinfo['cellblock'].row_list:
            if scores['empty'][row] == 1.0:
                continue
            print '%3d: ' % row,
            print '%8s (' % labels[row],
            print '%.3f, ' % scores['metadata'][row],
            print '%.3f, ' % scores['similarity'][row],
            print '%.3f)' % scores['header'][row]
                    
    def output_html(self):
        pass

    def validate(self, sheet_id, annotated_labels):
        results = {}
        labels = self._sheetinfo[sheet_id]['labels']
        for row in annotated_labels.keys():
            result = '%s_%s' % (annotated_labels[row], labels[row])
            results[result] = results.get(result, 0) + 1
        return results

    def _parse_sheet(self, sheet_id):
        """Parses spreadsheet cells and stores results in self.cells."""
        return xlrd_util.CellBlock(self._workbook, sheet_id)
        
    def _compute_row_scores(self, sheet_id):
        """Computes probability that this row is byline, header, or data."""
        
        sheetinfo = self._sheetinfo[sheet_id]
        
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
            empty[row] = xlrd_util.check_for_empty_row(cellblock, row)
            try:
                metadata_scores[row] = score_util.calc_metadata_score(cellblock, row)
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
            s = score_util.calc_sim_score(cellblock, row, next_row)
            h = score_util.calc_header_score(cellblock, row, next_row)
            
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

    def _dumb_classify(self, sheet_id):
        """Having computed metrics for each row, determine a label for each."""

        sheetinfo = self._sheetinfo[sheet_id]
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
        sheetinfo['labels'] = labels
        sheetinfo['scores']['row_scores'] = row_scores
    
    def _smart_classify(self, sheet_id):
        pass
