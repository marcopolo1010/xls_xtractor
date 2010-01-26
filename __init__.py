#
#
# 2010-01-25 mda     Copied logic from original project files.
# 2010-01-25 mda     Created file.

import xlrd_util, score_util

class XlsXtractor(object):
    self._workbook = None
    self._cells = None
    
    def __init__(self, filename):
        self._filename = filename
        self._workbook = xlrd_util.load_file(filename)
    
    def extract_schemas(self):
        for i in range(self._workbook.nsheets):
            self._cells = self._parse_sheets()
            self._compute_row_scores()
            self._dumb_classify()
            self._smart_classify()
        
    def output(format='html'):
        pass

    def output_html():
        pass

    def validate(annotated_labels):
        pass

    def _parse_sheet(sheet_id):
        """Parses spreadsheet cells and stores results in self.cells.
        """
        return xlrd_util.CellBlock(self._workbook, sheet_id)
        
    def _compute_row_scores():
        pass
