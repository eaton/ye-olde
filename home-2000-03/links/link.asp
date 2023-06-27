<%option explicit%>
<!-- #include virtual="/components/main.asp"-->
<!-- #include virtual="/components/DB_Utils.asp"-->
<%
dim sql, verbConn, linkSet               'ADO vars
dim link_id, title, description, url     'column names
dim randMax, redirectStr                 'misc vars

if request.querystring("id") = "" then
	'generate a random link, reload page with link information
	sql = "SELECT COUNT(link_id) AS link_count, MAX(link_id) AS link_max FROM link"
	MakeConn verbConn, MAIN_SITE_DB
	MakeRS linkSet, verbConn, sql

	randomize
	link_id = Int( (linkSet("link_count") - 1)*Rnd() + 1 )
	
	if not linkSet("link_max") = linkSet("link_count") then
		'there might be an error here. when you have time, do some better checking.
	end if
	
	redirectStr = "link.asp?id=" & link_id
	if request.querystring("action") = "jump" then
		redirectStr = redirectStr & "&action=jump"
	end if
	response.redirect redirectStr
else
	link_id = request.querystring("id")
	sql = "SELECT title, description, url FROM link WHERE link_id = " & link_id
	MakeConn verbConn, MAIN_SITE_DB
	MakeRS linkSet, verbConn, sql

	if request.querystring("action") = "jump" then
		'bring up the link information and fill variables
		redirectStr = linkSet("url")
		response.redirect redirectStr
	end if
	
	description = linkSet("description")
	url = linkSet("url")
	title = linkSet("title")
end if

%>
<html>

	<head>
		<meta http-equiv=content-type content="text/html;charset=iso-8859-1">
		<title>[ p r e d i c a t e - d o t - n e t ]</title>
		<link href=/components/predicate.css rel=styleSheet type=text/css>
	</head>

	<body bgcolor=gray>
		<center>
			<table border=0 cellpadding=0 cellspacing=0 width=100% height=95%>
				<tr>
					<td align=center valign=middle>
						<table border=0 cellpadding=0 cellspacing=0 width=550>
							<tr height=20>
								<td height=20 colspan=3 bgcolor=black align=left valign=middle>
									<p class=inverse>
									<table border=0 cellpadding=0 cellspacing=0 width=100%>
										<tr>
											<td>
												<div class=inverse>
													&nbsp;&nbsp;<a href=http://www.predicate.net>predicate.net</a>&nbsp;<b>::</b>&nbsp;<a href=/links/>links</a>&nbsp;<b>::</b>&nbsp;<a href="../<%=url%>"><%=linkSet("title")%></a></div>
											</td>
											<td>
												<div align=right>
													<img height=18 width=72 src=../day_images/trial_and_error.gif border=0></div>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td bgcolor=white align=center valign=middle colspan=3>
									<div class=normal>
										<b>
										<table border=0 cellpadding=0 cellspacing=5 width=100%>
											<tr>
												<td>
													<h3><a href="../<%=url%>"><%=linkSet("title")%></a></h3>
													<p><%if not isNull(description) then
												response.write makeHtmlStr(description, false)
												end if%></td>
											</tr>
										</table>
										</b></div>
								</td>
							</tr>
							<tr height=20>
								<td height=20 colspan=3 bgcolor=white align=left valign=middle>&nbsp;&nbsp;looking for solid web usability advice? visit jakob nielsen's <a href=http://www.useit.com>useit.com</a> site.</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</center>
	</body>

</html>
<%
Destroy linkSet
Destroy verbConn
%>