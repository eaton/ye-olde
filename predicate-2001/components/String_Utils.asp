<%

'--------------------------------------------------------------------------------
' Oodles and oodles of string manipulation routines. Very handy for misc. tasks.
' Code ©2000 ASP Emporium, http://www.aspEmporium.com
'
' AddChr, Censor, Decrypt, Encrypt, HTMLDecode, Strip, StripHTML, StrExpand
' added 9/11/2000
'
'--------------------------------------------------------------------------------


' The AddChr function is used to add characters to a string. There are three required 
' arguments: Input, Addition, Location. The Input argument is the string to manipulate. 
' The addition argument is a string of text to add to the Input string. The location 
' argument is a long value representing the place to add the text specified in Addition 
' to the Input string. 
' 
' Syntax: string = AddChr(Input, Addition, Location) 

Private Function AddChr(byVal string, byVal addition, byVal location)
	Dim leftString, rightString
	On Error Resume Next
	leftString = Left(string, location)
	If Err Then 
		leftString = ""
		location = 0
	End If
	rightString = Mid(string, location + 1)
	AddChr = leftString & addition & rightString
	On Error GoTo 0
End Function


'--------------------------------------------------------------------------------


' The Censor function removes disallowed words from a string. Censored words are kept 
' in a text file. The format of the text file is simple: one restricted word per line. 
' The system will read the restricted words and replace all of those words with xxx's 
' the same length as the removed word in the string. There is no limit to the amount 
' of words that can be restricted. 
' 
' Syntax: string = Censor(uncensoredstring) 

Private Function Censor(byVal string)
	Const WordList = "C:\dirtywords.txt"
	Dim objFSO, objFile, tmp, item, word, a, b, x, y, i, c, j
	Set objFSO  = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile( WordList )
	tmp = split( objFile.ReadAll(), vbCrLf )
	objFile.Close
	Set objFile = Nothing
	Set objFSO  = Nothing
	x = split( string, " " )
	for i = 0 to ubound(x)
		a = x(i)
		For each item in tmp
			b = item
			if cstr(trim(lcase(a))) = _
			   cstr(Trim(lcase(b))) then
				c = Len(a)
				a = ""
				for j = 1 to c
					a = a & "x"
				next
				a = a & ""
				exit for
			end if
		next
		x(i) = a
	next
	Censor = Join( x, " " )
End Function


'--------------------------------------------------------------------------------


' The Decrypt function changes an encrypted string into readable text. Decrypt only 
' can decrypt strings encoded with the Encrypt Function. 
' 
' Syntax: string = Decrypt(encryptedstring) 

Private Function Decrypt(ByVal encryptedstring)
	Dim x, i, tmp
	encryptedstring = StrReverse( encryptedstring )
	For i = 1 To Len( encryptedstring )
		x = Mid( encryptedstring, i, 1 )
		tmp = tmp & Chr( Asc( x ) - 1 )
	Next
	Decrypt = tmp
End Function


'--------------------------------------------------------------------------------


' The Encrypt function encrypts a string. To decrypt an encrypted string, use the 
' Decrypt Function. Encrypt replaces each letter in a string with a different character, 
' including spaces, and then reverses the scrambled string. Encrypt provides good enough 
' string encryption. 
' 
' Syntax: string = Encrypt(string) 

Private Function Encrypt(ByVal string)
	Dim x, i, tmp
	For i = 1 To Len( string )
		x = Mid( string, i, 1 )
		tmp = tmp & Chr( Asc( x ) + 1 )
	Next
	tmp = StrReverse( tmp )
	Encrypt = tmp
End Function


'--------------------------------------------------------------------------------


' The HTMLDecode function decodes an HTML encoded string back into the original html code. 
' 
' Syntax: string = HTMLDecode(encodedstring) 

Private Function HTMLDecode(byVal encodedstring)
	Dim tmp, i
	tmp = encodedstring
	tmp = Replace( tmp, "&quot;", chr(34) )
	tmp = Replace( tmp, "&lt;"  , chr(60) )
	tmp = Replace( tmp, "&gt;"  , chr(62) )
	tmp = Replace( tmp, "&amp;" , chr(38) )
	tmp = Replace( tmp, "&nbsp;", chr(32) )
	For i = 1 to 255
		tmp = Replace( tmp, "" & i & ";", chr( i ) )
	Next
	HTMLDecode = tmp
End Function


'--------------------------------------------------------------------------------


' The Strip function removes all white space from a string. 
' 
' Syntax: string = Strip(string) 

Private Function Strip(byVal string)
	Strip = Trim( Replace( Replace( Replace( _
		 Replace( Replace( string, vbCrLf, _
		 "" ), vbTab , "" ), " ", "" ), _
		 chr(10), "" ), chr(13), "" ) )
End Function


'--------------------------------------------------------------------------------


' The StripHTML function removes all HTML code from a string. 
' 
' Syntax: string = StripHTML(string) 

Private Function StripHTML(byVal string)
	Dim lngStart, lngEnd, strHTML
	string = Replace( string, vbTab, "" )
	string = Replace( string, vbCrLf, "" )
	string = Trim( string )
	do
		lngStart = Instr(string, "<")
		lngEnd   = InStr(string, ">")
		strHTML  = Mid( string, lngStart, _
			   lngEnd - lngStart + 1)
		string   = Trim(  Replace( string, strHTML, "" )  )
	loop until Not Instr(string, "<") _
		AND Not Instr(string, ">")
	If Instr( string, "<" ) Then _
		string = StripHTML( Trim( string ) )
	StripHTML = Trim( string )
End Function


'--------------------------------------------------------------------------------


' The StrExpand function expands a string by inserting spaces between each 
' character and three spaces if it encounters a space in the string. There are 
' two required arguments: string and usenbsp. String is the string to expand and 
' usenbsp is a boolean value (true, false) indicating whether the spaces should 
' be HTML nonbreaking spaces (&nbsp;) or plain spaces " ". 
' 
' Syntax: string = StrExpand(string, usenbsp) 

Private Function StrExpand(byVal string, byVal usenbsp)
	Dim Tmp, i
	For i = 1 to Len( string )
		Select Case CBool( usenbsp )
			Case False
				If Mid( string, i, 1 ) = " " Then
					Tmp = Tmp & "  "
				Else
					Tmp = Tmp & Mid( string, i, 1 ) & " "
				End If
			Case True
				If Mid( string, i, 1 ) = " " Then
					Tmp = Tmp & "  "
				Else
					Tmp = Tmp & Mid( string, i, 1 ) & " "
				End If
		End Select
	Next
	StrExpand = Tmp
End Function
%>
