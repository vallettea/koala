
import glob
import pandas
import re

import os
import sys
import os.path
import warnings
from reader.xml.constants import (
    DCORE_NS,
    COREPROPS_NS,
    DCTERMS_NS,
    SHEET_MAIN_NS,
    CONTYPES_NS,
    PKG_REL_NS,
    REL_NS,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
    WORKSHEET_TYPE,
    EXTERNAL_LINK,
)
from reader.xml.constants import (
    SHEET_MAIN_NS,
    REL_NS,
    EXT_TYPES,
    PKG_REL_NS
)
from reader.xml.functions import (
	fromstring, 
	tostring, 
	safe_iterator, 
	iterparse
)
import xml.etree.cElementTree as ET
import xml.parsers.expat as Expat

from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile
from sys import exc_info
from io import BytesIO

CENTRAL_DIRECTORY_SIGNATURE = b'\x50\x4b\x05\x06'
SUPPORTED_FORMATS = ('.xlsx', '.xlsm', '.xltx', '.xltm')

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

def read_named_ranges(xml_source, workbook):
    """Read named ranges, excluding poorly defined ranges."""
    sheetnames = set(sheet.title for sheet in workbook.worksheets)
    root = fromstring(xml_source)
    for name_node in safe_iterator(root, '{%s}definedName' %SHEET_MAIN_NS):

        range_name = name_node.get('name')
        if DISCARDED_RANGES.match(range_name):
            warnings.warn("Discarded range with reserved name")
            continue

        node_text = name_node.text

        if external_range(node_text):
            # treat names referring to external workbooks as values
            named_range = NamedValue(range_name, node_text)

        elif refers_to_range(node_text):
            destinations = split_named_range(node_text)
            # it can happen that a valid named range references
            # a missing worksheet, when Excel didn't properly maintain
            # the named range list
            destinations = [(workbook[sheet], cells) for sheet, cells in destinations
                            if sheet in sheetnames]
            if not destinations:
                continue
            named_range = NamedRange(range_name, destinations)
        else:
            named_range = NamedValue(range_name, node_text)

        named_range.scope = name_node.get("localSheetId")

        yield named_range

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

def repair_central_directory(zipFile, is_file_instance):
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

if __name__ == '__main__':

	COL_TAG = '{%s}col' % SHEET_MAIN_NS
	ROW_TAG = '{%s}row' % SHEET_MAIN_NS
	CELL_TAG = '{%s}c' % SHEET_MAIN_NS
	VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
	FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
	MERGE_TAG = '{%s}mergeCell' % SHEET_MAIN_NS
	INLINE_STRING = "{%s}is/{%s}t" % (SHEET_MAIN_NS, SHEET_MAIN_NS)
	INLINE_RICHTEXT = "{%s}is/{%s}r/{%s}t" % (SHEET_MAIN_NS, SHEET_MAIN_NS, SHEET_MAIN_NS)

	files = glob.iglob("./data/[0-9]*.xlsx")
	for file in files:	

		#try:
			file_name = os.path.abspath(file)
			print file_name
			
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

			valid_files = archive.namelist()
			

			named_range = []
			root = ET.fromstring(archive.read(ARC_WORKBOOK))
			for name_node in safe_iterator(root, '{%s}definedName' %SHEET_MAIN_NS):

				range_name = name_node.get('name')
		        print range_name
		        # if DISCARDED_RANGES.match(range_name):
		        #     warnings.warn("Discarded range with reserved name")
		        #     continue

		        # node_text = name_node.text

		        # if external_range(node_text):
		        #     # treat names referring to external workbooks as values
		        #     named_range = NamedValue(range_name, node_text)

		        # elif refers_to_range(node_text):
		        #     destinations = split_named_range(node_text)
		        #     # it can happen that a valid named range references
		        #     # a missing worksheet, when Excel didn't properly maintain
		        #     # the named range list
		        #     destinations = [(workbook[sheet], cells) for sheet, cells in destinations
		        #                     if sheet in sheetnames]
		        #     if not destinations:
		        #         continue
		        #     named_range = NamedRange(range_name, destinations)
		        # else:
		        #     named_range = NamedValue(range_name, node_text)


			cells = []
			for sheet in detect_worksheets(archive):
				sheet_name = sheet['title']

				if sheet_name == 'IHS': continue
				#print '-', sheet_name
				
				root = ET.fromstring(archive.read(sheet['path']))
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
						cells.append(cell);

			for cell in cells:
			#	print cell['a'], cell['f'], cell['v']
				if cell['f'] is not None:
					print cell['f']
    
		# except Exception as e:
		# 	print "Error with ", file, e