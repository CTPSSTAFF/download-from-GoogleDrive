# download_from_Google_Drive.py: Script to download all the contents of one's Google Drive.
#
# This tool requires the following libraries to have been installed:
#	1. pydrive
#	2. google-auth
# If these have not been installed, this can be accomplished as follows:
#	<python_intallation_directory>/python.exe -m pip install pydrive
#   <python_intallation_directory>/python.exe -m pip install google-auth
# 
# Thsi tool is written in Python 3.x, but the only requirement on Python 3.x
# is the syntax used for calls to "print."
#
# Reference documentation on working with files and folders in Google Drive:
#	1. Google's pydrive library for working with Drive in Python code:
#      https://pythonhosted.org/PyDrive/index.html
#	2. File management with pydrive: https://pythonhosted.org/PyDrive/filelist.html
#	3. File listing with pydrive: https://developers.google.com/drive/v2/reference/files/list#request
#	4. Google Drive REST API for "Files:list": https://developers.google.com/drive/v2/web/search-parameters
#	5. Google Drive REST API for working with folders: https://developers.google.com/drive/v2/web/folder
#      N.B.: In Google Drive, a "folder" is a file with the MIME type "application/vnd.google-apps.folder"
#            and with no extension.
#   6. List of MIME types supported by Google Drive: https://developers.google.com/drive/v3/web/mime-types
#
# License: GNU General Public License
#
# -- Ben Krepp 8/25/2017, 01/{24-26,29-31}/2018, 02/{01,07}/2018

import os
from os.path import splitext
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Module-level 'traceFlag' - global for starters.
# To do: make this into parameter.
# traceFlag = False

# Function: getMimeTypeAndSuffix
#
# Parameter:
#       driveFile       - Google Drive <file> object
#		traceFlag		- Boolean to indicate whether a verbose trace of the
#						  script's execution is to be printed. False by default.
#
# Return the appropriate mimetype and file extension (i.e., suffix) if needed
# to be used to download a given Google Drive file.
# Note: A preference here is made to download files MS Office formats, if available.
#
def getMimeTypeAndSuffix(driveFile, traceFlag):
	retval = { 'mimeType' : '', 'suffix' : '' }
	metadata = driveFile.metadata
	if not('exportLinks' in metadata):
		retval['mimeType'] = metadata['mimeType']
	# 
	# First, the 'document' formats.
	elif 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' in metadata['exportLinks']:
		# A Google Doc or an uploaded MS Word doc.
		retval['mimeType'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
		retval['suffix'] = '.docx'
	elif 'application/vnd.oasis.opendocument.text' in metadata['exportLinks']:
		# An OpenOffice document.
		retval['mimeType'] = 'application/vnd.oasis.opendocument.text'
		retval['suffix'] = '.odt'
	elif 'application/rtf' in metadata['exportLinks']:
		# A rich text format file.
		retval['mimeType'] = 'application/rtf'
		retval['suffix'] = '.rtf'	
	# 
	# Second, the 'spreadsheet' formats.
	elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in metadata['exportLinks']:
		# A Google Sheet or an uploaded MS Excel file.
		retval['mimeType'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
		retval['suffix'] = '.xlsx'
	elif 'application/x-vnd.oasis.opendocument.spreadsheet' in metadata['exportLinks']:
		# An OpenOffice sheet.
		retval['mimeType'] = 'application/x-vnd.oasis.opendocument.spreadsheet'
		retval['suffix'] = '.ods'
	elif 'text/csv' in metadata['exportLinks']:
		# A CSV file.
		retval['mimeType'] = 'text/csv'
		retval['suffix'] = '.csv'
	elif 'text/tsv' in metadata['exportLinks']:
		# A TSV file.
		retval['mimeType'] = 'text/tsv'
		retval['suffix'] = '.tsv'
	#
	# Third, the 'presentation' formats.
	elif 'application/vnd.openxmlformats-officedocument.presentationml.presentation' in metadata['exportLinks']:
		# A Google presentation or an uploaded PowerPoint file.
		retval['mimeType'] = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
		retval['suffix'] = '.pptx'
	elif 'application/vnd.oasis.opendocument.presentation' in metadata['exportLinks']:
		# An OpenOffice presentation.
		retval['mimeType'] = 'application/vnd.oasis.opendocument.presentation'
		retval['suffix'] = '.odp'
	# 
	# Fouth, drawing and image formats.
	elif 'image/jpeg' in metadata['exportLinks']:
		# A JPEG format image.
		retval['mimeType'] = 'image/jpeg'
		retval['suffix'] = '.jpg'
	elif 'image/png' in metadata['exportLinks']:
		# A PNG format image.
		retval['mimeType'] = 'image/png'
		retval['suffix'] = '.png'
	elif 'image/svg' in metadata['exportLinks']:
		# An SVG format image.
		retval['mimeType'] = 'image/svg'
		retval['suffix'] = '.svg'
	#
	# Fifth, throw up your hands and use PDF as a (hopefully) acceptable fallback format.
	else:
		retval['mimeType'] = 'application/pdf'
		retval['suffix'] = '.pdf'
	# end_if
	return retval
# end_def getMimeTypeAndSuffix()

# Function: makePcFileName
#
# Parameter
#       inFilename      - Input file name to be rendered politically correct
#		traceFlag		- Boolean to indicate whether a verbose trace of the
#						  script's execution is to be printed. False by default.
#
# Make a politically correct file name from a possibly un-politicaly correct file name.
# Political correctness is defined as not containing any characters that are illegal in
# a file or folder name on Windows. PC is accomplished by converting any such offending
# characters to underscores.
# 
# The following characters aren't legal in Windows file and folder names:
# 	~, #, %, &, *, {, }, /, \, :, <, >, ?, |, "
#
def makePcFileName(inFilename, traceFlag):
	retval = inFilename.replace('~', '_')
	retval = retval.replace('#', '_')
	retval = retval.replace('%', '_')
	retval = retval.replace('&', '_')
	retval = retval.replace('*', '_')
	retval = retval.replace('{', '_')
	retval = retval.replace('}', '_')
	retval = retval.replace('/', '_')
	retval = retval.replace('\\', '_')
	retval = retval.replace(':', '_')
	retval = retval.replace('<', '_')
	retval = retval.replace('>', '_')
	retval = retval.replace('?', '_')
	retval = retval.replace('|', '_')
	retval = retval.replace('"', '_')
	return retval
# end_def makePcFilename()

# Function: getDownloadFileName
#
# Parameter:
#       hostRootDir     - root directory in the host file system
#       driveName       - Name, i.e., 'title' of the file in Google Drive.
#       suffix          - Suffix (if any) of the file in Drive
#       i               - Iteration counter passed in; used to disambiguate
#                         multiple Drive files with the same 'title'.
#		traceFlag		- Boolean to indicate whether a verbose trace of the
#						  script's execution is to be printed. False by default.
#
# Synthesize and return the name of the file in the host file system
# to which a given file in Google Drive is to be downloaded.
#
def getDownloadFileName(hostRootDir, driveName, suffix, i, traceFlag):
	fileName,extension = splitext(driveName)
	if extension == '':
		# No extension in Google Drive file name, used passed-in suffix.
		extension = suffix
	retval = hostRootDir + '/' + makePcFileName(fileName, traceFlag)
	if i > 0:
		retval = retval + '_' + str(i)
	# end_if
	retval = retval + extension
	return retval
# end_def getDownloadFileName()

# Function: downloadAllFilesInFolder
#
# Parameters:
#       hostRootDir     - Root of the current directory in the host file system
#                         into which files are being downloaded.
#       driveRootDir    - 'id' or 'title' of current root of the Google Drive
#                         <file system> to be downloaded.
#                         For all directories other than the root of the user's
#                         Drive, this will be the 'id' of the relevant object
#                         in Drive. For the root of the user's Drive it is 'root'.
#       drive           - Google Drive 'handle' for authenticated user.
#		traceFlag		- Boolean to indicate whether a verbose trace of the
#						  script's execution is to be printed. False by default.
#
# Downloads all files in the specified Google Drive root directory specified
# by the "driveRootDir" parameter to the root directory in the host file system
# specified by the "hostRootDir" parameter.
#
# Notes:
#       1. When it encounters a folder, this function will call itself recursively to donwload
#          the contents of such folder.
#       2. When using drive.ListFiles() to query the contents of a folder, with the exception
#          of  when qurerying a user's "root" folder in drive the folder is specified by its 'id'.
#          The root folder is specified by the pseudo-id 'root'.
#
def downloadAllFilesInFolder(hostRootDir, driveRootDir, drive, traceFlag):
	# Get list of all files in the specified driveRootDir
	file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % driveRootDir}).GetList()
	
	for f in file_list:
		if traceFlag:
			outStr = 'Processing ' + f['title']
			print(outStr)
			outStr = '\tDrive id = ' + f['id']
			print(outStr)

		driveName = f['title']
		# The following ensures that Drive files whose names contain apostrophes 
		# will be passed appropriately to the drive.ListFile query.
		driveNameForQuery = driveName.replace("'", "\\'")
		queryStr2 = "title='" + driveNameForQuery + "'"		
		driveFiles = drive.ListFile({'q' : queryStr2}).GetList()
		
		i = 0
		for driveFile in driveFiles:
			# If driveFile is a folder, create one in the host file system,
			# and call downloadAllFilesInFolder recursively, passing in
			# the name of the newly created folder in the host file system
			# and the 'id' (NOT the 'title') of the folder in Google Drive.
			if driveFile.metadata['mimeType'] == 'application/vnd.google-apps.folder':
				# "Clean up" driveFile name to make it acceptable in the host file system, if neceesary.
				pcFolderName = makePcFileName(driveFile['title'], traceFlag)
				newFolderName = hostRootDir + '/' + pcFolderName
				os.makedirs(newFolderName)
				driveFolderId = driveFile['id']
				if traceFlag:
					outStr = 'Recursive call to download files to ' + newFolderName
					print(outStr)
					outStr = '\tDrive id = ' + driveFile['id']
					print(outStr)
				# end_if
				downloadAllFilesInFolder(newFolderName, driveFile['id'], drive, traceFlag)
			elif driveFile.metadata['mimeType'] == 'application/vnd.google-apps.form':
				outStr = "*** Google Drive does not support export/download of Google Forms: " + driveFile['title']
				print(outStr)
				outStr = '\tDrive id = ' + driveFile['id']
				print(outStr)
			else:
				# It's a vanilla file.
				# 'mtas' ---> 'mime type and suffix' :-)
				mtas = getMimeTypeAndSuffix(driveFile, traceFlag)
				mimeType = mtas['mimeType']
				suffix = mtas['suffix']
				downloadFileName = getDownloadFileName(hostRootDir, driveName, suffix, i, traceFlag)
				try:
					driveFile.GetContentFile(downloadFileName, mimeType)
					if traceFlag:
						outStr = 'Downloaded ' + driveFile['title'] + ' to ' + downloadFileName
						print(outStr)
						outStr = '\tDrive id = ' + driveFile['id']
						print(outStr)
					# end_if
				except:
					outStr = "*** FAILED to download " + driveFile['title'] + ' to ' + downloadFileName
					print(outStr)
					outStr = '\tDrive id = ' + driveFile['id']
					print(outStr)
			# end_if
			i = i + 1
	# end_for
# end_def downloadAllFilesInFolder()

# Function: downloadEverythingFromDrive
#
# Parameters:
#       hostRootDir     - Directory in the host file system that will be the root
#                         directory of the files downloaded.
#                         NOTE: This directory is ***created*** by this function.
#		traceFlag		- Boolean to indicate whether a verbose trace of the
#						  script's execution is to be printed. False by default.
#
# This performs initialziation, and then calls downloadAllFilesInFolder to download
# all files in the users Google Drive, starting at the 'root' folder.
#
def downloadEverythingFromDrive(hostRootDir, traceFlag=False):
	# Initialization
	gauth = GoogleAuth()
	gauth.LocalWebserverAuth() # Creates local webserver and auto handles authentication.
	os.makedirs(hostRootDir)   # Create root directory for download in host file system.
	drive = GoogleDrive(gauth)
	downloadAllFilesInFolder(hostRootDir, 'root', drive, traceFlag)
# end_def downloadEverythingFromDrive()