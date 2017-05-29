<!--#include file="../include/connection_it.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Excel Issue and Return Logs</title>
<link REL="stylesheet" HREF="include/stylesheet.css" TYPE="text/css">
<script language="JavaScript" type="text/javascript">
function searchIssue(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strStatus 		= document.forms[0].cboStatus.value;
    document.location.href = 'list_issues.asp?type=search&txtSearch=' + strSearch;	
}
    
function resetSearch(){
	document.location.href = 'list_issues.asp?type=reset';    
}  
</script>
</head>
<body>
<%
sub setSearch	
	select case trim(request("type"))
		case "reset"
			session("adm_issue_search") = ""
		case "search"
			session("adm_issue_search") = server.htmlencode(trim(Request("txtSearch")))
	end select
end sub

sub listSurvey	
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
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_excel_issue "
	strSQL = strSQL & " WHERE (issProduct LIKE '%" & session("adm_issue_search") & "%') "
	strSQL = strSQL & "	ORDER BY issID"
			
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

			strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""update_issue.asp?id=" & rs("issID") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>" & rs("issCreatedBy") & ", " & FormatDateTime(rs("issDateCreated"),1) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issASC") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issContactName") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issProduct") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issReportedFault") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issDiagnosedFault") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issReason") & "</td>"
			if rs("issReturnDate") = "01/01/1900" or rs("issReturnDate") = "1/1/1900" then
				strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"
			else
				strDisplayList = strDisplayList & "<td align=""center"">" & FormatDateTime(rs("issReturnDate"),1) & "</td>"
			end if
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issReturnConnote") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issComments") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issSpareParts") & "</td>"
			if rs("issDispatchDate") = "01/01/1900" or rs("issDispatchDate") = "1/1/1900" then
				strDisplayList = strDisplayList & "<td class=""orange_text"">TBA</td>"
			else
				strDisplayList = strDisplayList & "<td align=""center"">" & FormatDateTime(rs("issDispatchDate"),1) & "</td>"
			end if
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("issDispatchConnote") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next
	else
        strDisplayList = "<tr><td colspan=""14"" align=""center"" bgcolor=""white"">No records found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""14"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<h3>Total Records: " & intRecordCount & "</h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
	call setSearch
    call listSurvey
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <tr>
    <td valign="top" class="maincontent"><h2>Excel Issues &amp; Return Logs</h2>
      <table width="100%" cellpadding="4" cellspacing="0">
        <tr>
          <td width="180" valign="top"><div class="alert alert-success">
              <h2><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_issue.asp">New Record</a></h2>
            </div>
            <div align="left"><img src="images/legend-blue.gif" border="1" /> = updated today</div></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_issues.asp?type=search" method="post" onsubmit="searchIssue()">
                <strong>Search by Name:</strong>
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <input type="button" name="btnSearch" value="Search" onclick="searchIssue()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="4" class="database_records" width="100%">
        <tr class="innerdoctitle" align="center">
          <td></td>
          <td>Created</td>
          <td>ASC</td>
          <td>ASC Contact Name</td>
          <td>Product</td>
          <td>ASC Reported Fault</td>
          <td>Excel Diagnosed Fault</td>
          <td>Reason for return</td>
          <td>Return Date</td>
          <td>Return Connote</td>
          <td>Comments</td>
          <td>Spare Parts used</td>
          <td>Dispatch Date</td>
          <td>Dispatch Connote</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>