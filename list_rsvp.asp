<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="include/stylesheet.css" />
<script>
function searchRSVP(){
    var strSearch 		= document.forms[0].txtSearch.value;
	

    document.location.href = 'list_rsvp.asp?type=search&txtSearch=' + strSearch;
}

function resetSearch(){
	document.location.href = 'list_rsvp.asp?type=reset';
}
</script>
<title>Christmas RSVP List</title>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("christmas_search") 	= ""
						
			session("christmas_initial_page") = 1
		case "search"
			session("christmas_search") 	= trim(Request("txtSearch"))
			
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
	
	strSQL = "SELECT * FROM tbl_rsvp "
	strSQL = strSQL & "	WHERE  "
	strSQL = strSQL & "	name LIKE '%" & session("christmas_search") & "%' "
	strSQL = strSQL & "		OR diet LIKE '%" & session("christmas_search") & "%' "
	strSQL = strSQL & "		OR partner LIKE '%" & session("christmas_search") & "%' "
	strSQL = strSQL & "		OR partnerDiet LIKE '%" & session("christmas_search") & "%'"	
	
	strSQL = strSQL & "	ORDER BY name"
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("christmas_initial_page")

		For intRecord = 1 To rs.PageSize			
			'intAttendees = intAttendees + rs("no_attendees")
						
			strDays = DateDiff("d",rs("dateCreated"), strTodayDate)
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if

			strDisplayList = strDisplayList & "<td>" & rs("name") & ""
			if DateDiff("d",rs("dateCreated"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("diet") & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("partner") & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("partnerDiet") & "</td>"			
			strDisplayList = strDisplayList & "<td>"
			if rs("transportActivity") = 1 then
				strDisplayList = strDisplayList & "<img src=""images/tick.gif""> Bus from Yamaha HO to Staff Activity "
			else
				strDisplayList = strDisplayList & "<img src=""images/cross.gif""> Drive myself"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			select case rs("transportParty") 
				case 1
					strDisplayList = strDisplayList & "<img src=""images/tick.gif""> Bus from Staff Activity to Smart Artz"
				case 2
					strDisplayList = strDisplayList & "<img src=""images/tick.gif""> Bus from Staff Activity to Yamaha HO"
				case 0
					strDisplayList = strDisplayList & "<img src=""images/cross.gif""> Drive myself"
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("dateCreated"),1) & "</td>"						
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete this RSVP by " & rs("name") & " ?');"" href='delete_rsvp.asp?id=" & rs("id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""8"" align=""center"">No RSVP</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""8"" class=""recordspaging"">"	
	'strDisplayList = strDisplayList & "<h2>Total: " & intRecordCount & " records</h2>"    
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

dim rs, intPageCount, intpage, intRecord, strDisplayList
%>
<div style="padding:20px 20px 20px 20px;">
  <h1 align="center">RSVP List</h1>
  <div class="alert alert-search">
    <form name="frmSearch" id="frmSearch" action="list_rsvp.asp?type=search" method="post" onsubmit="searchRSVP()">
      Name / Partner's name / Diet:
      <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
      
      <input type="button" name="btnSearch" value="Search" onclick="searchRSVP()" />
      <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
    </form>
  </div>
  <table cellspacing="0" cellpadding="5" class="loan_table" border="0">
    <thead>
      <tr>
        <td>Name</td>
        <td>Diet</td>       
        <td>Partner</td>
        <td>Partner's Diet</td>        
        <td>Transport to Staff Activity</td>
        <td>Transport to Smart Artz</td>
        <td>Submitted</td>
        <td></td>
      </tr>
    </thead>
    <tbody>
      <%= strDisplayList %>
    </tbody>
  </table>
</div>
</body>
</html>