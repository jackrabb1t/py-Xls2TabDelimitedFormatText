"""Convert an Excel spreadsheet to a tab-delimited text file"""
import sys
try:
    import xlrd
except ImportError:
    print >> sys.stderr, 'Please see http://pypi.python.org/pypi/xlrd' 
    sys.exit(1)

def create_parser():
    from optparse import OptionParser
    parser = OptionParser('usage: python %s filename.xls [options]' % __file__)
    parser.add_option(
            '--worksheet', 
            '-w', 
            dest='worksheet', 
            type='string',
            default=None, 
            help='The name of the worksheet to open')
    return parser

def get_rows(xls_fname, sheet_name=None):
    """
    Get the table rows from an Excel spreadsheet.

        xls_fname   The filename of the spreadsheet.
        sheet_name  The worsheet to open.  If None, opens the first worksheet.

    Returns a list of table rows.  Each table row is a list of columns.  
    Each column is a string.
    """
    book = xlrd.open_workbook(xls_fname)
    #
    # TODO: proper handling of exceptions and error conditions.  Use asserts
    # for now to see exactly what can break.
    #
    assert book
    assert book.nsheets, 'No worksheets in file'
    if sheet_name:
        sheet = book.sheet_by_name(sheet_name)
        assert book.sheet_loaded(sheet_name), 'Failed to load sheet'
    else:
        sheet = book.sheet_by_index(0)
        assert book.sheet_loaded(0), 'Failed to load sheet'

    rows = []
    for i in range(sheet.nrows):
        col = []
        for j in range(sheet.ncols):
            val = str(sheet.cell_value(i, j))
            #
            # TODO: can this actually happen?
            #
            assert val.find('\t') == -1, 'Found Tab in cell value'
            col.append(val)
        rows.append(col)
    return rows

def main():
    parser = create_parser()
    (options, args) = parser.parse_args()
    if len(args) != 1:
        parser.error('incorrect number of arguments')

    rows = get_rows(args[0], options.worksheet)
    print '\n'.join(map(lambda x: '\t'.join(x), rows))

if __name__ == '__main__':
    main()