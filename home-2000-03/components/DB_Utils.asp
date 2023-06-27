<%

'--------------------------------------------------------------------------------
'  DB_Utils 1.0 -- Jeff's  bare bones DB connectivity stuff, a couple
'  string manipulation utils, and other miscellaney. there's also a spot
'  for constant declaration -- it gets modified depending on the
'  specific project.
'                             --jeff, 8/15/00
'--------------------------------------------------------------------------------

CONST ferret_db = "DBQ=D:\FTP\tamarac\Database\ferret.mdb;Driver={Microsoft Access Driver (*.mdb)};"
CONST predicate_db = "DBQ=D:\FTP\tamarac\Database\predicate.mdb;Driver={Microsoft Access Driver (*.mdb)};"
CONST verb_db = "DBQ=D:\FTP\tamarac\Database\verb.mdb;Driver={Microsoft Access Driver (*.mdb)};"
CONST newchurch_db = "DBQ=D:\FTP\tamarac\Database\newchurch.mdb;Driver={Microsoft Access Driver (*.mdb)};"


'--------------------------------------------------------------------------------
'  Database Routines Here -- last updated 8/15/00
'--------------------------------------------------------------------------------

Sub MakeConn(Conn,connectionstring)
  Set Conn = Server.CreateObject("ADODB.Connection")
  Conn.Open connectionstring
End Sub

Sub MakeRs(Rs,Conn,sql)
  Set Rs = Server.CreateObject("ADODB.Recordset")
  Rs.Open sql, Conn, adOpenStatic, adLockReadOnly, adCmdText
End Sub

Sub Modify(Conn,sql)
  Conn.Execute sql
End Sub

Sub Destroy(Name)
  Name.Close
  Set Name = Nothing
End Sub



'--------------------------------------------------------------------------------
'  SQL String Manip Stuff here -- last updated 8/15/00
'--------------------------------------------------------------------------------

Function makeSqlStr (string)
  if string = "" then
    string = "NULL"
  else
    string = replace(string, "'", "''")
    string = "'" & string & "'"
  end if
  makeSqlStr = string
End function

Function makeHtmlStr (string, noTags)
  if noTags then
    string = server.HTMLEncode
  end if
  string = replace(string, vbCRLF, "<br>")
  makeHtmlStr = string
end function
%>