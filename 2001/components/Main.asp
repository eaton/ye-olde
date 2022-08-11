<%
CONST Security_Log_Path = "/logs/security_log.txt"
CONST Visitor_Log_Path = "/logs/visitor_log.txt"
CONST Error_Log_Path = "/logs/error_log.txt"
	
CONST LOGIN_PAGE = "/community/login.asp"
CONST SITE_URL = "http://www.predicate.net"
CONST MAIN_SITE_DB = "DBQ=D:\FTP\drumrush\Database\predicate.mdb;Driver={Microsoft Access Driver (*.mdb)};"


'--------------------------------------------------------------------------------


' The WriteLog statement writes a line of text into a log file. The WriteLog 
' statement automatically inserts the date and time of the log addition along 
' with any text you specify in the logevent argument. WriteLog automatically 
' appends new information to any existing data in a log file if one is found. 
' If the specified log file is not found, it is created for that entry. 
' 
' Syntax: WriteLog pathname, logevent 

Private Sub WriteLog(byVal pathname, byVal logevent)
	Dim objFSO, objFile
	Set objFSO  = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(pathname, 8, True)
	objFile.WriteLine Now() & vbTab & logevent
	objFile.Close
	Set objFile = Nothing
	Set objFSO  = Nothing
End Sub

%>

<!--#include virtual="/components/adovbs.inc"-->