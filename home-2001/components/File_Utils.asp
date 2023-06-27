<%

'--------------------------------------------------------------------------------
' Oodles and oodles of file manipulation routines. Very handy for misc. tasks.
' Code ©2000 ASP Emporium, http://www.aspEmporium.com
'
' Dir, File, FileCopy, FileDateTime, FileLen, FileRead, FileWrite, Kill,
' MkDir, MkFile, RmDir, Title, UnMappath, WriteLog added 9/11/2000
'
'--------------------------------------------------------------------------------


' The Dir function checks for the existence of a folder or directory. Pathname 
' must be a path to a directory. Returns True if the folder or directory exists 
' and False if it does not exist. 
' 
' Syntax: boolean = Dir(pathname) 

Private Function Dir(byVal pathname)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Dir = objFSO.FolderExists(pathname)
	Set objFSO = Nothing
End Function


'--------------------------------------------------------------------------------


' The File function checks for the existence of a file. Pathname must be a path to 
' a file, including extension. Returns True if the file exists and False if it does 
' not exist. 
' 
' Syntax: boolean = File(pathname) 

Private Function File(byVal pathname)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	File = objFSO.FileExists(pathname)
	Set objFSO = Nothing
End Function


'--------------------------------------------------------------------------------


' The FileCopy statement copies a file. There are two required arguments, source and 
' destination. The source argument is the full path to the file you are copying and 
' the destination argument is the full path of where the copied file is to be placed. 
' As an added bonus, my version of FileCopy also copies directories and folders. 
' 
' Syntax: FileCopy source, destination 

Private Sub FileCopy(byVal source, byVal destination)
	Dim objFSO, objToCopy, boolErr, strErrDesc
	On Error Resume Next
	Set objFSO = Server.CreateObject("scripting.filesystemobject")
	if instr( right( source, 4 ), "." ) then
		 ' probably a file
		Set objToCopy = objFSO.GetFile(source)
	else
		 ' probably a directory or folder
		Set objToCopy = objFSO.GetFolder(source)
	end if
	objToCopy.Copy destination
	if Err Then 
		boolErr = True
		strErrDesc = Err.Description
	end if
	Set objToCopy = Nothing
	Set objFSO = Nothing
	On Error GoTo 0
	if boolErr then Err.Raise 5104, "FileCopy Statement", strErrDesc
End Sub


'--------------------------------------------------------------------------------


' The FileDateTime function returns the time and date a file was last modified. If the 
' file is not found or an error occurs during the processing of the function, Null is 
' returned. 
' 
' Syntax: datetime = FileDateTime(pathname) 

Private Function FileDateTime(byVal pathname)
	Dim objFSO, objFile
	On Error Resume Next
	Set objFSO	= Server.CreateObject("Scripting.FileSystemObject")
	Set objFile	= objFSO.GetFile(pathname)
	If Err Then
		FileDateTime	= Null
	Else
		FileDateTime	= CDate( objFile.DateLastModified )
	End If
	Set objFile	= Nothing
	Set objFSO	= Nothing
	On Error GoTo 0
End Function


'--------------------------------------------------------------------------------


' The FileLen function returns a long representing the size of a file in bytes. If 
' file is not found, FileLen returns Null. 
' 
' Syntax: long = FileLen(pathname) 

Private Function FileLen(byVal pathname)
	Dim objFSO, objFile
	On Error Resume Next
	Set objFSO	= Server.CreateObject("Scripting.FileSystemObject")
	Set objFile	= objFSO.GetFile(pathname)
	If Err Then
		FileLen = Null
	Else
		FileLen = CLng( objFile.Size )
	End If
	Set objFile	= Nothing
	Set objFSO	= Nothing
	On Error GoTo 0
End Function


'--------------------------------------------------------------------------------


' The FileRead function allows the contents of a file to be read. The FileRead 
' function has one required argument, pathname, which represents the complete path 
' to an already existing file on the server. The ReadFile function returns a string 
' representing the contents of the file. ReadFile returns Null in the event of an 
' error. 
' 
' Syntax: string = FileRead(pathname) 

Private Function FileRead(byVal pathname)
	dim objFSO, objFile, tmp
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(pathname, 1, False)
	tmp = objFile.ReadAll
	If Err Then
		FileRead = Null
	Else
		FileRead = tmp
	End If
	objFile.Close
	Set objFile = Nothing
	Set objFSO = Nothing
	On Error GoTo 0
End Function


'--------------------------------------------------------------------------------


' The FileWrite statement allows writing to a file. The FileWrite statement has three 
' required arguments: pathname, texttowrite, and overwrite. The first argument is 
' pathname. Pathname must be the complete path to an already existing file. The second 
' argument texttowrite is a string containing the text to add to the file. To add more 
' than one line of text, use vbCrLf. The third argument is a boolean representing 
' overwriting. If the overwrite argument is set to True, the complete contents of the 
' file is replaced with texttowrite. If it is set to False, the texttowrite string is 
' appended to the file's contents. 
' 
' Syntax: FileWrite pathname, texttowrite, overwrite 

Private Sub FileWrite(byVal pathname, byVal strToWrite, byVal boolOverWrite)
	dim objFSO, objFile, boolErr, strErrDesc, lngWriteMethod
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if boolOverWrite then
		lngWriteMethod = 2
	else
		lngWriteMethod = 8
	end if
	Set objFile = objFSO.OpenTextFile(pathname, lngWriteMethod, False)
	objFile.Write strToWrite
	If Err Then
		boolErr = True
		strErrDesc = Err.Description
	End If
	objFile.Close
	Set objFile = Nothing
	Set objFSO = Nothing
	On Error GoTo 0
	if boolErr then Err.Raise 5107, "FileWrite Statement", strErrDesc
End Sub


'--------------------------------------------------------------------------------


' The Kill statement deletes files. The required pathname argument must be absolute 
' and include the drive letter. This supports the use of the wildcard character * 
' to delete mulitple files. 
' 
' Syntax: Kill pathname 

Private Sub Kill(byVal pathname)
	Dim objFSO, boolErr, strErrDesc
	On Error Resume Next
	Set objFSO = Server.CreateObject("scripting.filesystemobject")
	objFSO.DeleteFile pathname
	if Err Then 
		boolErr = True
		strErrDesc = Err.Description
	end if
	Set objFSO = Nothing
	On Error GoTo 0
	if boolErr then Err.Raise 5102, "Kill Statement", strErrDesc
End Sub


'--------------------------------------------------------------------------------


' The MkDir statement creates a new directory or folder. The required path 
' argument must be absolute and include the drive letter. If you are unsure 
' of the path to your web server's root, use server.mappath in the path argument. 
' 
' Syntax: MkDir path 

Private Sub MkDir(byVal path)
	Dim objFSO, boolErr, strErrDesc
	boolErr = False
	On Error Resume Next
	Set objFSO = Server.CreateObject("scripting.filesystemobject")
	objFSO.CreateFolder path
	if Err Then 
		boolErr = True
		strErrDesc = Err.Description
	end if
	Set objFSO = Nothing
	On Error GoTo 0
	if boolErr then Err.Raise 5101, "MkDir Statement", strErrDesc
End Sub


'--------------------------------------------------------------------------------


' The MkFile statement creates a file on the server. The MkFile statement has 
' one required argument, pathname. Pathname must be the complete path and file 
' name with extension to be created. 
' 
' Syntax:  MkFile pathname 

Private Sub MkFile(byVal pathname)
	Dim objFSO, boolErr, strErrDesc
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	if objFSO.FileExists(pathname) then
		err.raise 5106, "MkFile Statement", "File [" & pathname & "] " & _
			"Already Exists. Use the Kill statement to delete files."
	else
		On Error Resume Next
		objFSO.CreateTextFile pathname, 2, True
		if Err Then 
			boolErr = True
			strErrDesc = Err.Description
		end if
		On Error GoTo 0
		if boolErr then Err.Raise 5106, "MkFile Statement", strErrDesc
	end if
	Set objFSO = Nothing
End Sub


'--------------------------------------------------------------------------------


' The RmDir statement deletes a directory or folder. The required path argument 
' must be absolute and include the drive letter. If you are unsure of the path 
' to your web server's root, use server.mappath in the path argument. 
' 
' Syntax: RmDir path 

Private Sub RmDir(byVal path)
	Dim objFSO, boolErr, strErrDesc
	boolErr = False
	On Error Resume Next
	Set objFSO = Server.CreateObject("scripting.filesystemobject")
	objFSO.DeleteFolder path
	if Err Then 
		boolErr = True
		strErrDesc = Err.Description
	end if
	Set objFSO = Nothing
	On Error GoTo 0
	if boolErr then Err.Raise 5100, "RmDir Statement", strErrDesc
End Sub


'--------------------------------------------------------------------------------


' The Title function returns a string representing the title of a web page. It 
' specifically scans a web page looking for the title tags and returns whatever 
' is in between. The required argument pathname is the full path to a web page 
' [*.htm, *.html, *.asp, *.cfm, *.cgi, etc...]. 
' 
' Syntax: string = Title(pathname) 

Private Function Title(byVal pathname)
	dim objFSO, objFile, a, tmp, firstCt, secondCt
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(pathname, 1, False)
	a = objFile.ReadAll()
	objFile.Close
	Set objFile = Nothing
	Set objFSO = Nothing
	firstCt  = InStr(UCASE(a), "<TITLE>") + 7
	secondCt = InStr(UCASE(a), "</TITLE>")
	tmp = Mid( a, firstCt, secondCt - firstCt )
	Title = CStr( Trim( tmp ) )
End Function


'--------------------------------------------------------------------------------


' The UnMappath function returns the virtual path of the absolute path specfied 
' in the required argument pathname. 
' 
' Syntax: string = UnMappath(pathname) 

Private Function UnMappath(byVal pathname)
	dim tmp, strRoot
	strRoot = Server.Mappath("/")
	tmp = replace( lcase( pathname ), lcase( strRoot ), "" )
	tmp = replace( tmp, "\", "/" )
	UnMappath = tmp
End Function


%>
