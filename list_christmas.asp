<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Christmas RSVP List</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script language="JavaScript" type="text/javascript">
function searchRSVP(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strPartner  	= document.forms[0].cboPartner.value;
	var strDiet  		= document.forms[0].cboDiet.value;
	var strTransport 	= document.forms[0].cboTransport.value;

    document.location.href = 'list_christmas.asp?type=search&txtSearch=' + strSearch + '&partner=' + strPartner + '&diet=' + strDiet + '&transport=' + strTransport;
}

function resetSearch(){
	document.location.href = 'list_christmas.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("christmas_search") 	= ""
			session("christmas_partner") 	= ""
			session("christmas_diet") 		= ""
			session("christmas_transport") 	= ""
			session("christmas_initial_page") = 1
		case "search"
			session("christmas_search") 	= trim(Request("txtSearch"))
			session("christmas_partner") 	= Trim(request("partner"))
			session("christmas_diet") 		= Trim(request("diet"))
			session("christmas_transport") 	= Trim(Request("transport"))
			session("christmas_initial_page") = 1
	end select
end sub

sub displayRSVP
	dim iRecordCount
	iRecordCount = 0
    dim strSQL
	dim intRecordCount
	dim strTodayDate
	dim strDays
	
	dim intAttendees
	intAttendees = 0

	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_christmas "
	strSQL = strSQL & "	WHERE partner LIKE '%" & session("christmas_partner") & "%' "
	strSQL = strSQL & "	AND diet LIKE '%" & session("christmas_diet") & "%' "
	strSQL = strSQL & "	AND (name LIKE '%" & session("christmas_search") & "%' "
	strSQL = strSQL & "		OR partner_name LIKE '%" & session("christmas_search") & "%' "
	strSQL = strSQL & "		OR diet_req LIKE '%" & session("christmas_search") & "%')"	
	strSQL = strSQL & "	AND transport LIKE '%" & session("christmas_transport") & "%' "
	strSQL = strSQL & "	ORDER BY name"
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("christmas_initial_page")

		For intRecord = 1 To rs.PageSize			
			intAttendees = intAttendees + rs("no_attendees")
						
			strDays = DateDiff("d",rs("date_created"), strTodayDate)			
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if

			strDisplayList = strDisplayList & "<td>" & rs("name") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("partner") = 1 then
				strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"
			else
				strDisplayList = strDisplayList & "<img src=""images/cross.gif"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("partner_name") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("diet") = 1 then
				strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"
			else
				strDisplayList = strDisplayList & "<img src=""images/cross.gif"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("diet_req") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("transport") = 1 then
				strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"
			else
				strDisplayList = strDisplayList & "<img src=""images/cross.gif"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("date_created"),1) & "</td>"						
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""7"" align=""center"">No RSVP.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""7"" class=""recordspaging"">"	
	strDisplayList = strDisplayList & "<h2>Total: " & intAttendees & " attendees</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
	
	call setSearch

    if trim(session("christmas_initial_page")) = "" then
    	session("christmas_initial_page") = 1
	end if

    call displayRSVP
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<div style="padding:20px 20px 20px 20px;">  
  <h2>Christmas RSVP List</h2>
  <div class="alert alert-search">
    <form name="frmSearch" id="frmSearch" action="list_christmas.asp?type=search" method="post" onsubmit="searchRSVP()">     
      Name / Partner's name / Dietary Requirements:
      <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
      <select name="cboPartner" onchange="searchRSVP()">
        <option value="">All Partner</option>
        <option <% if session("christmas_partner") = "1" then Response.Write " selected" end if%> value="1">Partner is attending</option>
        <option <% if session("christmas_partner") = "0" then Response.Write " selected" end if%> value="0">Partner is not attending</option>
      </select>
      <select name="cboDiet" onchange="searchRSVP()">
        <option value="">All Diet</option>
        <option <% if session("christmas_diet") = "1" then Response.Write " selected" end if%> value="1">Have dietary requirement</option>
        <option <% if session("christmas_diet") = "0" then Response.Write " selected" end if%> value="0">Do not have dietary requirement</option>
      </select>
      <select name="cboTransport" onchange="searchRSVP()">
        <option value="">All Transport</option>
        <option <% if session("christmas_transport") = "1" then Response.Write " selected" end if%> value="1">Need transport</option>
        <option <% if session("christmas_transport") = "0" then Response.Write " selected" end if%> value="0">Do not need transport</option>
      </select>
      <input type="button" name="btnSearch" value="Search" onclick="searchRSVP()" />
      <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
    </form>
  </div>
  <table cellspacing="0" cellpadding="5" class="loan_table" border="0">
  <thead>
    <tr>
      <td>Name</td>
      <td>Partner</td>
      <td>Partner Name</td>
      <td>Diet</td>
      <td>Dietary Requirements</td>      
      <td>Transport</td>
      <td>Submitted</td>
    </tr>
    </thead>
    <tbody>
    <%= strDisplayList %>
    </tbody>
  </table>
</div>
</body>
</html>