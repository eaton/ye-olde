<%
'--------------------------------------------------------------------------------
' 
' Security_Utils.asp -- includes code for checking and changing authentication
' state, enforcing security for protected pages, etc etc.
' Include at the top of all secured pages.
' 
'--------------------------------------------------------------------------------

' enforceSecurity is used at the head of a document, and redirects to the site's
' login page if the user hasn't yet signed in, or if they have insufficient
' permissions to access the particular page. Requires that Application("Login_Page")
' be set in the site's global.asa. Need provisions for setting security levels.
'
' Usage: enforceSecurity 5

Private Sub EnforceSecurity(byVal requiredPermissionLevel)
	If Session("Authenticated") = 0 then
		Session("Source_URL") = Request.ServerVariables("SCRIPT_NAME")
		response.redirect LOGIN_PAGE & "?reason=requiredlogin"
	ElseIf Session("Authenticated") < requiredPermissionLevel then
		Session("Source_URL") = Request.ServerVariables("SCRIPT_NAME")
		response.redirect LOGIN_PAGE & "?reason=permissions"
	End If
end Sub


'--------------------------------------------------------------------------------

' getUserDetails() returns full information on a user when given a userID.
' for permissions information, use getUserPermissions(). For logins, use
' checkUserLogin().
'
' Usage: getUserDetails("1234")

Private Function getUserDetails(byVal userID) 
	Dim cnUser, rsUser, objUserDetails, i
	Set objUserDetails = CreateObject("Scripting.Dictionary")
	MakeConn cnUser, MAIN_SITE_DB
	MakeRS rsUser, cnUser, "SELECT * FROM profile WHERE user_id = " & userID & ";"
	if not rsUser.EOF then
		For each i in rsUser
			objUserDetails.Add rsUser.Key(i), rsUser.Item(i)
		Next
	end if
	getUserDetails = objUserDetails
	Destroy rsUser
	Destroy cnUser
End Function


'--------------------------------------------------------------------------------


Private Function getUserPermissions(byVal userID) 
	getUserPermissions = 0
End Function


'--------------------------------------------------------------------------------


Private Function checkUserLogin(byVal nickname, byVal password)
	Dim cnUser, rsUser, rsuser_id, rsnickname, rspermission_id
	rsuser_id = 0
	rsnickname = "ERROR"
	rspermission_id = 0
	MakeConn cnUser, MAIN_SITE_DB
	MakeRS rsUser, cnUser, "SELECT user_id, nickname, permission_id FROM profile WHERE nickname = '" & nickname & "' AND password = '" & password & "';"
	if not rsUser.EOF then
		rsuser_id = rsUser("user_id")
		rsnickname = rsUser("nickname")
		rspermission_id =  rsUser("permission_id")
	end if
	checkUserLogin = array(rsuser_id, rsnickname, rspermission_id)
	Destroy rsUser
	Destroy cnUser
End Function


'--------------------------------------------------------------------------------

Private Function hashPassword(byVal Password)
	hashPassword = Password
End Function

'--------------------------------------------------------------------------------

Private Function unHashPassword(byVal Password)
	unHashPassword = Password
End Function
%>