import glob
import os.path
import warnings
from io import BytesIO


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
from koala.unzip import read_archive
from koala.excel import read_named_ranges, read_cells


if __name__ == '__main__':


    files = glob.iglob("./data/*.xlsx")
    for file in files:  

        file_name = os.path.abspath(file)
        print file_name
        
        archive = read_archive(file_name)        

        named_range = read_named_ranges(archive)
        print named_range

        cells = read_cells(archive)
        

        for cell in cells:
          if cell['f'] is not None:
                print cell['f']
