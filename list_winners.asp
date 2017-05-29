<!--#include file="../include/connection_auction.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>#WinnersOnly</title>
<link REL="stylesheet" HREF="include/stylesheet.css" TYPE="text/css">
<script language="JavaScript" type="text/javascript">
function searchWinner(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strStatus 		= document.forms[0].cboStatus.value;
    document.location.href = 'list_winners.asp?type=search&txtSearch=' + strSearch + '&cboStatus=' + strStatus;	
}
    
function resetSearch(){
	document.location.href = 'list_winners.asp?type=reset';    
}  
</script>
</head>
<body>
<%
sub setSearch	
	select case trim(request("type"))
		case "reset"
			session("adm_winner_search") 		= ""
		case "search"
			session("adm_winner_search") 		= server.htmlencode(trim(Request("txtSearch")))
	end select
end sub

sub listWinner	
	dim iRecordCount
	iRecordCount = 0
    dim strSortBy
	dim strSortItem
    dim strSQL
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount	
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 500
	
	strSQL = "SELECT * FROM tblQARegistration "
	strSQL = strSQL & " WHERE (regUsername LIKE '%" & session("adm_winner_search") & "%' "
	strSQL = strSQL & "			OR regName LIKE '%" & session("adm_winner_search") & "%') "
	strSQL = strSQL & "		AND regID IN (SELECT aucCurrentBidder FROM tblQAAuctions WHERE aucEnded = 'Y')"
	strSQL = strSQL & "	ORDER BY regName"
			
	'response.write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	
		For intRecord = 1 To rs.PageSize 			
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if
			strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""update_winner.asp?id=" & rs("regID") & """>Edit</a></td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("regUsername") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("regName") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & rs("regEmail") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("regSalesOrderNo") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("regInvoiceNo") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next
	else
        strDisplayList = "<tr><td colspan=""5"" align=""center"">No winners found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""5"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "Total Winners: " & intRecordCount & "</small>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
	call setSearch
    call listWinner
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table border="0" cellpadding="0" cellspacing="0" width="600">
  <tr>
    <td valign="top" class="maincontent">
    <h2>yBay Winners | <a href="list_winners2.asp">Firesale Winners</a></h2>
    <div class="alert alert-search">
      <form name="frmSearch" id="frmSearch" action="list_winners.asp?type=search" method="post" onsubmit="searchWinner()">
      <strong>Search:</strong> Username / Name
        <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
        <input type="button" name="btnSearch" value="Search" onclick="searchWinner()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
      </form>
      </div>
      <table cellspacing="0" cellpadding="4" class="database_records" width="100%">
        <tr class="innerdoctitle" align="center">
          <td>&nbsp;</td>
          <td>Username</td>
          <td>Name</td>
          <td>Sales Order No</td>
          <td>Invoice No</td>
        </tr>
        <%= strDisplayList %>
      </table>
    </td>
  </tr>
</table>
</body>
</html>