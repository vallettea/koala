from __future__ import print_function

from io import BytesIO
import re
import os
import json

from openpyxl.formula.translate import Translator
from openpyxl.cell.text import Text
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.xml.functions import iterparse, fromstring, safe_iterator
try:
    from xml.etree.cElementTree import ElementTree as ET
except ImportError:
    from xml.etree.ElementTree import ElementTree as ET
from openpyxl.xml.constants import (
    SHEET_MAIN_NS,
    REL_NS,
    PKG_REL_NS,
    CONTYPES_NS,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
    WORKSHEET_TYPE,
    SHARED_STRINGS
)

curfile = os.path.abspath(os.path.dirname(__file__))

with open('%s/functions.json' % curfile, 'r') as file:
    existing = json.load(file)

from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile

from koala.Cell import Cell
from koala.utils import CELL_REF_RE, col2num

FLOAT_REGEX = re.compile(r"\.|[E-e]")
CENTRAL_DIRECTORY_SIGNATURE = b'\x50\x4b\x05\x06'

def repair_central_directory(zipFile, is_file_instance): # source: https://bitbucket.org/openpyxl/openpyxl/src/93604327bce7aac5e8270674579af76d390e09c0/openpyxl/reader/excel.py?at=default&fileviewer=file-view-default
    ''' trims trailing data from the central directory
    code taken from http://stackoverflow.com/a/7457686/570216, courtesy of Uri Cohen
    '''
    f = zipFile if is_file_instance else open(zipFile, 'rb+')
    data = f.read()
    pos = data.find(CENTRAL_DIRECTORY_SIGNATURE)  # End of central directory signature
    if (pos > 0):
        sio = BytesIO(data)
        sio.seek(pos + 22)  # size of 'ZIP end of central directory record'
        sio.truncate()
        sio.seek(0)
        return sio

    f.seek(0)
    return f

def read_archive(file_name):
    is_file_like = hasattr(file_name, 'read')
    if is_file_like:
        # fileobject must have been opened with 'rb' flag
        # it is required by zipfile
        if getattr(file_name, 'encoding', None) is not None:
            raise IOError("File-object must be opened in binary mode")

    try:
        archive = ZipFile(file_name, 'r', ZIP_DEFLATED)
    except BadZipfile as e:
        f = repair_central_directory(file_name, is_file_like)
        archive = ZipFile(f, 'r', ZIP_DEFLATED)

    return archive

def _cast_number(value): # source: https://bitbucket.org/openpyxl/openpyxl/src/93604327bce7aac5e8270674579af76d390e09c0/openpyxl/cell/read_only.py?at=default&fileviewer=file-view-default
    "Convert numbers as string to an int or float"
    m = FLOAT_REGEX.search(value)
    if m is not None:
        return float(value)
    return int(value)

debug = False

def read_named_ranges(archive):

    root = fromstring(archive.read(ARC_WORKBOOK))

    dict = {}

    for name_node in safe_iterator(root, '{%s}definedName' % SHEET_MAIN_NS):
        name = name_node.get('name')
        # if name in dict:
        #     raise Exception('Named_range %s is defined in multiple sheets' % name)

        if not name_node.get('hidden'):
            if name_node.get('name') == 'tR':
                dict[name_node.get('name')] = 'Depreciation!A1:A1000'
            elif '!#REF' in name_node.text:
                dict[name_node.get('name')] = '#REF!'
            else:
                dict[name_node.get('name')] = name_node.text.replace('$','').replace(" ","")

    return dict

def read_cells(archive, ignore_sheets = [], ignore_hidden = False):
    global debug

    # print('___### Reading Cells from XLSX ###___')

    cells = {}

    functions = set()

    cts = dict(read_content_types(archive))

    strings_path = cts.get(SHARED_STRINGS) # source: https://bitbucket.org/openpyxl/openpyxl/src/93604327bce7aac5e8270674579af76d390e09c0/openpyxl/reader/excel.py?at=default&fileviewer=file-view-default
    if strings_path is not None:
        if strings_path.startswith("/"):
            strings_path = strings_path[1:]
        shared_strings = read_string_table(archive.read(strings_path))
    else:
        shared_strings = []

    for sheet in detect_worksheets(archive):
        sheet_name = sheet['title']

        function_map = {}

        if sheet_name in ignore_sheets: continue

        root = fromstring(archive.read(sheet['path'])) # it is necessary to use cElementTree from xml module, otherwise root.findall doesn't work as it should

        hidden_cols = False
        nb_hidden = 0

        if ignore_hidden:
            hidden_col_min = None
            hidden_col_max = None

            for col in root.findall('.//{%s}cols/*' % SHEET_MAIN_NS):
                if 'hidden' in col.attrib and col.attrib['hidden'] == '1':
                    hidden_cols = True
                    hidden_col_min = int(col.attrib['min'])
                    hidden_col_max = int(col.attrib['max'])

        for c in root.findall('.//{%s}c/*/..' % SHEET_MAIN_NS):
            cell_data_type = c.get('t', 'n') # if no type assigned, assign 'number'
            cell_address = c.attrib['r']

            skip = False

            if hidden_cols:
                found = re.search(CELL_REF_RE, cell_address)
                col = col2num(found.group(1))

                if col >= hidden_col_min and col <= hidden_col_max:
                    nb_hidden += 1
                    skip = True

            if not skip:
                cell = {'a': '%s!%s' % (sheet_name, cell_address), 'f': None, 'v': None}
                if debug:
                    print('Cell', cell['a'])
                for child in c:
                    child_data_type = child.get('t', 'n') # if no type assigned, assign 'number'

                    if child.tag == '{%s}f' % SHEET_MAIN_NS :
                        if 'ref' in child.attrib: # the first cell of a shared formula has a 'ref' attribute
                            if debug:
                                print('*** Found definition of shared formula ***', child.text, child.attrib['ref'])
                            if "si" in child.attrib:
                                function_map[child.attrib['si']] = (child.attrib['ref'], Translator(str('=' + child.text), cell_address)) # translator of openpyxl needs a unicode argument that starts with '='
                            # else:
                            #     print "Encountered cell with ref but not si: ", sheet_name, child.attrib['ref']
                        if child_data_type == 'shared':
                            if debug:
                                print('*** Found child %s of shared formula %s ***' % (cell_address, child.attrib['si']))

                            ref = function_map[child.attrib['si']][0]
                            formula = function_map[child.attrib['si']][1]

                            translated = formula.translate_formula(cell_address)
                            cell['f'] = translated[1:] # we need to get rid of the '='

                        else:
                            cell['f'] = child.text

                    elif child.tag == '{%s}v' % SHEET_MAIN_NS :
                        if cell_data_type == 's' or cell_data_type == 'str': # value is a string
                            try: # if it fails, it means that cell content is a string calculated from a formula
                                cell['v'] = shared_strings[int(child.text)]
                            except:
                                cell['v'] = child.text
                        elif cell_data_type == 'b':
                            cell['v'] = bool(int(child.text))
                        elif cell_data_type == 'n':
                            cell['v'] = _cast_number(child.text)

                    elif child.text is None:
                        continue

                if cell['f'] is not None:

                    pattern = re.compile(r"([A-Z][A-Z0-9]*)\(")
                    found = re.findall(pattern, cell['f'])

                    map(lambda x: functions.add(x), found)

                if cell['f'] is not None or cell['v'] is not None:
                    should_eval = 'always' if cell['f'] is not None and 'OFFSET' in cell['f'] else 'normal'

                    # cleaned_formula = cell['f']
                    cleaned_formula = cell['f'].replace(", ", ",") if cell['f'] is not None else None
                    if "!" in cell_address:
                        cells[cell_address] = Cell(cell_address, sheet_name, value = cell['v'], formula = cleaned_formula, should_eval=should_eval)
                    else:
                        cells[sheet_name + "!" + cell_address] = Cell(cell_address, sheet_name, value = cell['v'], formula = cleaned_formula, should_eval=should_eval)

        # if nb_hidden > 0:
            # print('Ignored %i hidden cells in sheet %s' % (nb_hidden, sheet_name))

    # print('Nb of different functions %i' % len(functions))
    # print(functions)

    # for f in functions:
    #     if f not in existing:
    #         print('== Missing function: %s' % f)

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

def read_string_table(xml_source):
    """Read in all shared strings in the table"""
    strings = []
    src = _get_xml_iter(xml_source)

    for _, node in iterparse(src):
        if node.tag == '{%s}si' % SHEET_MAIN_NS:

            text = Text.from_tree(node).content
            text = text.replace('x005F_', '')
            strings.append(text)

            node.clear()

    return IndexedList(strings)


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
