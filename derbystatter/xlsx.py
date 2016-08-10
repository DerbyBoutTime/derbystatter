"""A simple xlsx file reader, supporting a very limited subset of the format, designed to
be as similar to xld module API as possible

"""
USE_EL_TREE = 1
if USE_EL_TREE:
    from xml.etree.cElementTree import parse
    nsS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    nsR = '{http://schemas.openxmlformats.org/package/2006/relationships}'
    nsP = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
else:
    from xml.dom.minidom import parse

import zipfile

XL_CELL_EMPTY = 0
XL_CELL_TEXT = 1
XL_CELL_NUMBER = 2
XL_CELL_DATE = 3
XL_CELL_BOOLEAN = 4
XL_CELL_ERROR = 5
XL_CELL_BLANK = 6


class xlsx:
    def __init__(self, fpath):
        """Open the zip file, and set up lazy caches for the xml files and specific sheets"""
        self.zf = zipfile.ZipFile(fpath, 'r')
        self.xmlFiles = {}
        self.sheets = {}

    def xml(self, name):
        """Extra a parsed DOM of the given XML file in the archive"""
        if self.xmlFiles.has_key(name):
            return self.xmlFiles[name]
        try:
            f = self.zf.open(name)
        except:
            return None
        dom = parse(f)
        if USE_EL_TREE:
            dom = dom.getroot()
        self.xmlFiles[name] = dom
        return dom

    def sharedStringElement(self):
        """Get the xlsx shared string file dom"""
        return self.xml('xl/sharedStrings.xml')

    def sharedStringWithIndex(self, index):
        """Find a given shared string by index"""
        if USE_EL_TREE:
            return self.sharedStringElement().findall(nsS + 'si')[index]
        return self.sharedStringElement().getElementsByTagName('si')[index]

    def sheet_by_name(self, name):
        """Find a specific sheet in the xlsx file by name"""
        if not self.sheets:
            workbook = self.xml("xl/workbook.xml")
            workbookRels = self.xml("xl/_rels/workbook.xml.rels")
            relMap = {}
            if USE_EL_TREE:
                for rel in workbookRels.findall(nsR + 'Relationship'):
                    relMap[rel.get('Id')] = rel.get('Target')
                for sheets in workbook.findall(nsS + 'sheets'):
                    for sheet in sheets.findall(nsS + 'sheet'):
                        #print sheet.toxml()
                        #                    print sheet.attributes._attrs
                        childName = relMap[sheet.get(nsP + 'id')]
                        sheetName = sheet.get('name')
                        self.sheets[sheetName] = {'path': childName}
            else:
                for rel in workbookRels.getElementsByTagName('Relationship'):
                    relMap[rel.attributes['Id'].value] = rel.attributes[
                        'Target'].value
                for sheets in workbook.getElementsByTagName('sheets'):
                    for sheet in sheets.getElementsByTagName('sheet'):
                        #print sheet.toxml()
                        #                    print sheet.attributes._attrs
                        childName = relMap[sheet.attributes['r:id'].value]
                        sheetName = sheet.attributes['name'].value
                        self.sheets[sheetName] = {'path': childName}
        value = self.sheets[name]
        if not value.has_key('sheet'):
            sheetXML = self.xml('xl/' + value['path'])
            # print "Loading sheet", name, ':', value['path']
            sheet = xlsxSheet(sheetXML, self, value['path'])
            value['sheet'] = sheet
            # try to load rels
            relsPathParts = value['path'].split('/')
            relsPathParts.insert(-1, '_rels')
            relsPathParts[-1] = relsPathParts[-1] + ".rels"
            relXML = self.xml('xl/' + '/'.join(relsPathParts))
            value['rels'] = relXML
        return value['sheet']

    def relsForSheet(self, sheet):
        for value in self.sheets.values():
            if value.get('sheet') == sheet:
                return value.get('rels')
        return None


class xlsxSheet:
    """Represents a specific sheet in the workbook"""

    def __init__(self, xml, document, path):
        self.xml = xml
        self.book = document
        self.rows = []
        self.path = path

    def allComments(self):
        """Get all the comments for this sheet"""
        rels = self.book.relsForSheet(self)
        if not rels:
            return None
        if USE_EL_TREE:
            import os
            retval = {}
            for rel in rels.findall(nsR + 'Relationship'):
                if rel.get(
                        'Type'
                ) == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
                    baseDir = os.path.split(self.path)[0]
                    commentPath = os.path.normpath(
                        os.path.join(baseDir, rel.get('Target'))
                    )
                    commentsXML = self.book.xml('xl/' + commentPath)
                    commentList = commentsXML.find(nsS + 'commentList')
                    for comment in commentList.findall(nsS + 'comment'):
                        commentText = ""
                        for t in comment.findall(nsS + 'text'):
                            for i in t.itertext():
                                commentText += i
                        retval[comment.get('ref')] = commentText
            return retval

    def row(self, rowIndex):
        """Get a specific row in the sheet"""
        if not self.rows:
            self.sharedFormulas = []
            if USE_EL_TREE:
                for el in self.xml.find(nsS + 'sheetData').findall(nsS + 'row'):
                    rnum = int(el.get('r'))
                    rlen = len(el.get('r'))
                    #print "found row", el, "num", rnum, "len", rlen
                    while len(self.rows) <= rnum:
                        self.rows.append({})
                    cols = self.rows[rnum]
                    for elc in el.findall(nsS + 'c'):
                        rname = elc.get('r')
                        cols[rname[:-rlen]] = elc
            else:
                for el in self.xml.getElementsByTagName('sheetData')[
                        0].getElementsByTagName('row'):
                    rnum = int(el.attributes['r'].value)
                    rlen = len(el.attributes['r'].value)
                    while len(self.rows) <= rnum:
                        self.rows.append({})
                    cols = self.rows[rnum]
                    for elc in el.getElementsByTagName('c'):
                        rname = elc.attributes['r'].value
                        cols[rname[:-rlen]] = elc
        # check for shared formulas
        return self.rows[rowIndex]

    def colName(self, colIndex):
        """Convert a colIndex (0..25, etc...) as A..Z, AA..ZZ"""
        if colIndex < 26:
            return chr(ord('A') + colIndex)
        return chr(ord('A') + colIndex / 26 - 1) + chr(ord('A') + colIndex % 26)

    def cell(self, rowIndex, colIndex):
        """Get the data of a specific cell by row and column (zero based)"""
        row = self.row(rowIndex + 1)
        colName = self.colName(colIndex)
        if not row.has_key(colName):
            return empty_cell
        cell = row[colName]
        if USE_EL_TREE:
            if cell.get('si'):
                print "Cell", colName, rowIndex + 1, "has si:", cell.text
            v = cell.findall(nsS + 'v')
        else:
            if cell.attributes.has_key('si'):
                print "Cell", colName, rowIndex + 1, "has si:", cell.toxml()
            v = cell.getElementsByTagName('v')
        if not v:
            return empty_cell
        if USE_EL_TREE:
            v = v[0].text
            if cell.get('t'):
                cellType = cell.get('t')
                if cellType == 's':
                    v = int(v)
                    return xlsxValue(
                        XL_CELL_TEXT, self.book.sharedStringWithIndex(v)[0].text
                    )
                if cellType == 'str':
                    return xlsxValue(XL_CELL_TEXT, v)
        else:
            v = v[0].firstChild.toxml()
            if cell.attributes.has_key('t'):
                cellType = cell.attributes['t'].value
                if cellType == 's':
                    v = int(v)
                    return xlsxValue(
                        XL_CELL_TEXT, self.book.sharedStringWithIndex(
                            v
                        ).firstChild.firstChild.toxml()
                    )
                if cellType == 'str':
                    return xlsxValue(XL_CELL_TEXT, v)
        return xlsxValue(XL_CELL_NUMBER, float(v))


class xlsxValue:
    """Represents the contents of a given cell in a sheet, as a type and value"""

    def __init__(self, type, value):
        self.type = type
        self.value = value

    def __repr__(self):
        if self.type == XL_CELL_EMPTY:
            return "<EMPTY>"
        if self.type == XL_CELL_TEXT:
            return '"' + self.value + '"'
        if self.type == XL_CELL_NUMBER:
            return "%g" % (self.value, )
        if self.type == XL_CELL_ERROR:
            return "<ERROR " + self.value + ">"
        return "<???>"


empty_cell = xlsxValue(XL_CELL_EMPTY, u'')


def open_workbook(f):
    return xlsx(f)


if __name__ == '__main__':
    x = xlsx(
        '/Users/gandreas/Documents/STATS-2014-04-27-NorthStarRollerGirls_vs_OldCapitolCityRollerGirls.xlsx'
    )
    IBRF = x.sheet_by_name("IGRF")

    print IBRF.cell(2, 1)
    print IBRF.cell(4, 1)
    print IBRF.cell(10, 1)
    print IBRF.allComments()
    Lineups = x.sheet_by_name('Lineups')
    print Lineups.cell(0, 0)
