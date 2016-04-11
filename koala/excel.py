
from koala.xml.constants import (
    SHEET_MAIN_NS,
    REL_NS,
    PKG_REL_NS,
    CONTYPES_NS,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
    WORKSHEET_TYPE,
)
from koala.xml.functions import fromstring, safe_iterator

def read_named_ranges(archive):

    root = fromstring(archive.read(ARC_WORKBOOK))

    return {
        name_node.get('name') : name_node.text
        for name_node in safe_iterator(root, '{%s}definedName' % SHEET_MAIN_NS)
    }
        

def read_cells(archive):
    cells = []
    for sheet in detect_worksheets(archive):
        sheet_name = sheet['title']

        if sheet_name == 'IHS': continue
        
        root = fromstring(archive.read(sheet['path']))
        for c in root.findall('.//{%s}c/*/..' % SHEET_MAIN_NS):
            cell = {'a': '%s!%s' % (sheet_name,c.attrib['r']), 'f': None, 'v': None}
            for child in c:
                if child.text is None: 
                    continue
                elif child.tag == '{%s}f' % SHEET_MAIN_NS :
                    cell['f'] = child.text
                elif child.tag == '{%s}v' % SHEET_MAIN_NS :
                    cell['v'] = child.text
            if cell['f'] or cell['v']:
                cells.append(cell)
    return cells

def read_rels(archive):
    """Read relationships for a workbook"""
    xml_source = archive.read(ARC_WORKBOOK_RELS)
    tree = fromstring(xml_source)
    for element in safe_iterator(tree, '{%s}Relationship' % PKG_REL_NS):
        rId = element.get('Id')
        pth = element.get("Target")
        typ = element.get('Type')
        # normalise path
        if pth.startswith("/xl"):
            pth = pth.replace("/xl", "xl")
        elif not pth.startswith("xl") and not pth.startswith(".."):
            pth = "xl/" + pth
        yield rId, {'path':pth, 'type':typ}

def read_content_types(archive):
    """Read content types."""
    xml_source = archive.read(ARC_CONTENT_TYPES)
    root = fromstring(xml_source)
    contents_root = root.findall('{%s}Override' % CONTYPES_NS)
    for type in contents_root:
        yield type.get('ContentType'), type.get('PartName')



def read_sheets(archive):
    """Read worksheet titles and ids for a workbook"""
    xml_source = archive.read(ARC_WORKBOOK)
    tree = fromstring(xml_source)
    for element in safe_iterator(tree, '{%s}sheet' % SHEET_MAIN_NS):
        attrib = element.attrib
        attrib['id'] = attrib["{%s}id" % REL_NS]
        del attrib["{%s}id" % REL_NS]
        if attrib['id']:
            yield attrib

def detect_worksheets(archive):
    """Return a list of worksheets"""
    # content types has a list of paths but no titles
    # workbook has a list of titles and relIds but no paths
    # workbook_rels has a list of relIds and paths but no titles
    # rels = {'id':{'title':'', 'path':''} }
    content_types = read_content_types(archive)
    valid_sheets = dict((path, ct) for ct, path in content_types if ct == WORKSHEET_TYPE)
    rels = dict(read_rels(archive))
    for sheet in read_sheets(archive):
        rel = rels[sheet['id']]
        rel['title'] = sheet['name']
        rel['sheet_id'] = sheet['sheetId']
        rel['state'] = sheet.get('state', 'visible')
        if ("/" + rel['path'] in valid_sheets
            or "worksheets" in rel['path']): # fallback in case content type is missing
            yield rel


def _get_xml_iter(xml_source):
    """
    Possible inputs: strings, bytes, members of zipfile, temporary file
    Always return a file like object
    """
    if not hasattr(xml_source, 'read'):
        try:
            xml_source = xml_source.encode("utf-8")
        except (AttributeError, UnicodeDecodeError):
            pass
        return BytesIO(xml_source)
    else:
        try:
            xml_source.seek(0)
        except:
            # could be a zipfile
            pass
        return xml_source