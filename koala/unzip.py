#
# source: https://bitbucket.org/openpyxl/openpyxl/src/93604327bce7aac5e8270674579af76d390e09c0/openpyxl/reader/excel.py?at=default&fileviewer=file-view-default 
#______________________________________________________________________________________________________________________________________________________________

from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile

CENTRAL_DIRECTORY_SIGNATURE = b'\x50\x4b\x05\x06'

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