#
#
# 2010-01-25 mda     Copied logic from original project files.
# 2010-01-25 mda     Created file.

import xlrd_util, score_util, show_schema

class XlsXtractor(object):
    workbook = None
    sheetinfo = {}
    
    def __init__(self, filename):
        self.filename = filename
        self.workbook = xlrd_util.load_file(filename)
        if self.workbook:
            self.valid = True
            self.nsheets = self.workbook.nsheets
        else:
            self.valid = False
    
    def extract_schemas(self, sheet_id):
        sheetinfo = {}
        self.sheetinfo[sheet_id] = sheetinfo
        
        sheetinfo['cellblock'] = self._parse_sheet(sheet_id)
        score_util.compute_row_scores(sheetinfo)
        score_util.dumb_classify(sheetinfo)
        score_util.smart_classify(sheetinfo)
        
    def output(self, sheet_id, format='html'):
        if format == 'text':
            output_text(sheet_id)
        else:
            output_html(sheet_id)

    def output_text(self, sheet_id):
        sheetinfo = self._sheetinfo[sheet_id]
        scores = sheetinfo['scores']
        labels = sheetinfo['labels']
        print 'ROW |  LABEL   |  META  |  SIM   | HEADER'
        print '----+----------+--------+--------+-------'
        for row in sheetinfo['cellblock'].row_list:
            if scores['empty'][row] == 1.0:
                continue
            print '%3d: ' % row,
            print '%8s (' % labels[row],
            print '% .3f, ' % scores['metadata'][row],
            print '% .3f, ' % scores['similarity'][row],
            print '% .3f)' % scores['header'][row]

    def output_html(self):
        pass

    def validate(self, sheet_id, annotated_labels):
        results = {}
        labels = self.sheetinfo[sheet_id]['labels']
        for row in annotated_labels.keys():
            result = '%s_%s' % (annotated_labels.get(row,''), labels[row])
            results[result] = results.get(result, 0) + 1
        return results

    def _parse_sheet(self, sheet_id):
        """Parses spreadsheet cells and stores results in self.cells."""
        return xlrd_util.CellBlock(self.workbook, sheet_id)
    
