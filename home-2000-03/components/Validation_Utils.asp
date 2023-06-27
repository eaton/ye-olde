<%

'--------------------------------------------------------------------------------
' Oodles and oodles of data validation routines. Very handy for misc. tasks.
' Code ©2000 ASP Emporium, http://www.aspEmporium.com
'
' IsAlpha, IsAlphaNumeric, IsEmail, IsState, IsWorthy, IsZip added 9/11/2000
'--------------------------------------------------------------------------------


' The IsAlpha function checks a string and allows only white space, letters of the 
' alphabet, or underscore characters to pass as valid input. 
' 
' Syntax: boolean = IsAlpha(string) 

Private Function IsAlpha(byVal string)
    dim regExp, match, i, spec
    For i = 1 to Len( string )
        spec = Mid(string, i, 1)
        Set regExp = New RegExp
        regExp.Global = True
        regExp.IgnoreCase = True
        regExp.Pattern = "[A-Z]|[a-z]|\s|[_]"
        set match = regExp.Execute(spec)
        If match.count = 0 then
            IsAlpha = False
            Exit Function
        End If
        Set regExp = Nothing
    Next
    IsAlpha = True
End Function


'--------------------------------------------------------------------------------


' The IsAlphaNumeric function checks a string and allows only white space, letters 
' of the alphabet, numbers, or underscore characters to pass as valid input. 
' 
' Syntax: boolean = IsAlphaNumeric(string) 

Private Function IsAlphaNumeric(byVal string)
    dim regExp, match, i, spec
    For i = 1 to Len( string )
        spec = Mid(string, i, 1)
        Set regExp = New RegExp
        regExp.Global = True
        regExp.IgnoreCase = True
        regExp.Pattern = "[A-Z]|[a-z]|\s|[_]|[0-9]|[.]"
        set match = regExp.Execute(spec)
        If match.count = 0 then
            IsAlphaNumeric = False
            Exit Function
        End If
        Set regExp = Nothing
    Next
    IsAlphaNumeric = True
End Function


'--------------------------------------------------------------------------------


' The IsEmail function performs a thorough check of an entered email address. The 
' following criteria is checked by IsEmail:
'    * only one @ sign permitted
'    * domain extension separator (.) must come after the @ symbol
'    * must be at least 2 characters (letters) after the domain extension separator
'    * rejects all illegal characters including spaces.
'    * Allows only numbers, letters, the underscore (_) and the dash (-) character 
'       as valid input (excluding the mandatory "@" and "." symbols).
'    * Minimum of 6 characters
' IsEmail returns True if an email address is validated successfully, otherwise False 
' is returned. If an error is encountered when processing an email address, IsEmail 
' returns False. 
' 
' Syntax: boolean = IsEmail(mailaddress) 

Private Function IsEmail(byVal mailaddress)
    Dim tmp, x, y, bErr, tmp2, objReg
    Dim objMatch, z, i

    bErr = False
    tmp = Trim( mailaddress )
    tmp = CStr( mailaddress )

     ' minimum 6 characters...
    if len(tmp) < 6 then
        IsEmail = False
        Exit Function
    end if

     ' need an @ but only 1 is allowed
    If instr(tmp, "@") then
        x = instr(tmp, "@")
        y = instr(x + 1, tmp, "@")
        On Error Resume Next
        y = CLng(y)
        If Err Then bErr = True Else bErr = False
        On Error GoTo 0
        If bErr Then
            IsEmail = False
            Exit Function
        End If
        if y <> 0 then
            IsEmail = False
            Exit Function
        end if
    Else
        IsEmail = False
        Exit Function
    End If

     ' the "." must come after the "@"
    If InStr( Left( tmp, CLng(x) ), "." ) Then
        IsEmail = False
        Exit Function
    Else
        tmp2 = Right( tmp, Len(tmp) - CLng(x) )
        If InStr( tmp2, "." ) Then
             ' must have at least one character between @ and .
            Set objReg = New RegExp
            With objReg
                .Global = True
                .IgnoreCase = True
                .Pattern = "[A-Z]|[0-9]"
                Set objMatch = .Execute(tmp2)
            End With
            If objMatch.Count = 0 then
                IsEmail = False
                Exit Function
            End If
            Set objMatch = Nothing
            Set objReg = Nothing
        Else
            IsEmail = False
            Exit Function
        End If
    End If

     ' needs to have at least 2 characters (letters) after the .
    z = InStr( tmp, "." )
    tmp2 = Right( tmp, Len(tmp) - z )
    Set objReg = New RegExp
    With objReg
        .Global = True
        .IgnoreCase = True
        .Pattern = "[A-Z][A-Z]"
        Set objMatch = .Execute(tmp2)
    End With
    If objMatch.Count = 0 then
        IsEmail = False
        Exit Function
    End If
    Set objMatch = Nothing
    Set objReg = Nothing

     ' check for illegal characters
    For i = 1 to Len(tmp)
        tmp2 = Mid( tmp, i, 1 )
        Select Case tmp2
            Case "(", ")", ";", ":", ",", "/", "'", chr(34), _
                 "~", "`", "!", "#", "$", "%", "^", "&", "*", _
                 "+", "=", "[", "]", "{", "}", "|", "\", "?", _
                 " ", "<", ">"
                IsEmail = False
                Exit Function
            Case Else
        End Select
    Next

     ' if an address makes it through, it's an email address
    IsEmail = True
End Function


'--------------------------------------------------------------------------------


' The IsState function checks a two-letter string against all valid postal 
' abbreviations for US States. There are currently 51 (including DC) state 
' abbreviations recognized by IsState. IsState returns true if a state abbreviation 
' is valid or False if the abbreviation is unknown. 
' 
' Syntax: boolean = IsState(stateabbr) 

Private Function IsState(byVal stateabbr)
    If Len( stateabbr ) <> 2 Then
        IsState = False
        Exit Function
    ElseIf IsNull( stateabbr ) Then
        IsState = False
        Exit Function
    ElseIf IsEmpty( stateabbr ) Then
        IsState = False
        Exit Function
    ElseIf IsNumeric( stateabbr ) Then
        IsState = False
        Exit Function
    Else
        stateabbr = CStr( stateabbr )
        Select Case UCase(stateabbr)
            Case "AL", "AK", "AZ", "AR", "CA", _
                 "CO", "CT", "DE", "DC", "FL", _
                 "GA", "HI", "ID", "IL", "IN", _
                 "IA", "KS", "KY", "LA", "ME", _
                 "MD", "MA", "MI", "MN", "MS", _
                 "MO", "MT", "NE", "NV", "NH", _
                 "NJ", "NM", "NY", "NC", "ND", _
                 "OH", "OK", "OR", "PA", "RI", _
                 "SC", "SD", "TN", "TX", "UT", _
                 "VT", "VA", "WA", "WV", "WI", _
                 "WY"
                IsState = True
                Exit Function
            Case Else
                IsState = False
                Exit Function
        End Select
    End If
End Function


'--------------------------------------------------------------------------------


' The IsZip function tests a string or a long for validity as a zip code. IsZip 
' will validate either five (xxxxx) or nine digit (xxxxx-xxxx) zip codes. While 
' the IsZip will not verify a zip code, it will ensure that the input is in valid 
' zip code format. Returns True if the input is recognized as a zip code or False 
' if the input is not a zip code. Null or Empty input returns False. 
' 
' Syntax: boolean = IsZip(zipcode) 

Private Function IsZip(byVal zipcode)
	Dim reg
	if IsNull(zipcode) then
		IsZip = False
		exit function
	elseif IsEmpty(zipcode) then
		IsZip = False
		exit function
	end if
	zipcode = Trim( CStr( zipcode ) )
	Select Case Clng( Len(zipcode) )
		Case 10
			Set reg = New RegExp
			With reg
				.IgnoreCase = True
				.Global = True
				.Pattern = "[0-9][0-9][0-9][0-9][0-9]" & _
					   "-[0-9][0-9][0-9][0-9]"
				If .Test(zipcode) Then
					IsZip = True
				Else
					IsZip = False
				End If
			End With
			Set reg = Nothing
			Exit Function
		Case 5
			Set reg = New RegExp
			With reg
				.IgnoreCase = True
				.Global = True
				.Pattern = "[0-9][0-9][0-9][0-9][0-9]"
				If .Test(zipcode) Then
					IsZip = True
				Else
					IsZip = False
				End If
			End With
			Set reg = Nothing
			Exit Function
		Case Else
			IsZip = False
			Exit Function
	End Select
End Function
%>


