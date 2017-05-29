<%
'setup for Australian Date/Time
session.lcid = 2057
session.timeout = 420

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsOffice365.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Cache-control" content="no-store">
<title>View  your Office 365 login</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
</head>
<body>
<%
sub main
	session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
	
	call getUser(session("logged_username"))
	
	if not isNull(session("usrUserID")) then	
		call updateUserStatus(session("usrUserID"))
	end if
end sub

call main

dim strMessageText
%>
<p><small>You are logged in as: <%= session("logged_username") %></small></p>
<table width="300" border="0" cellspacing="0" cellpadding="5" style="font-size:large">
  <tr>
    <td>User ID:</td>
    <td><%= session("usrUserID") %></td>
  </tr>
  <tr>
    <td>Default Password:</td>
    <td><%= session("usrPassword") %></td>
  </tr>
</table>
<p><em>If your password is displayed as ? - this may mean your password has already been changed</em>.</p>
<p><%= strMessageText %></p>
</body>
</html>