<!--#include file="../include/connection_it.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Loan Stock Workflow Basket</title>
<link REL="stylesheet" HREF="include/stylesheet.css" TYPE="text/css">
<script language="JavaScript" type="text/javascript">
function searchBasket(){    
    var strSearch 		= document.forms[0].txtSearch.value;
    document.location.href = 'view_basket.asp?type=search&txtSearch=' + strSearch;	
}
    
function resetSearch(){
	document.location.href = 'view_basket.asp?type=reset';    
}  
</script>
</head>
<body style="padding:20px 20px 20px 20px;">
<%
sub setSearch	
	select case trim(request("type"))
		case "reset"
			session("loan_basket_search") = ""
		case "search"
			session("loan_basket_search") = server.htmlencode(trim(Request("txtSearch")))
	end select
end sub

sub listBasket	
	dim iRecordCount
	iRecordCount = 0
    dim strSortBy
	dim strSortItem
    dim strSQL
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount	
	
    call OpenWorkflowDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 800
	
	strSQL = "SELECT * FROM workflow_loan_return_item_list "
	strSQL = strSQL & " WHERE (product_code LIKE '%" & session("loan_basket_search") & "%') "
	strSQL = strSQL & "				AND acc_code LIKE '%" & UCASE(trim(session("emp_initial"))) & "%' "
	'strSQL = strSQL & "				AND reference_id IS NULL "
	strSQL = strSQL & "	ORDER BY date_created"
	
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

			strDisplayList = strDisplayList & "<td align=""left"" nowrap>" & rs("date_created") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("product_code") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("serial_number") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & FormatNumber(rs("product_lic")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if len(rs("reference_id")) > 1 then
				strDisplayList = strDisplayList & "Processed"
			else
				strDisplayList = strDisplayList & "Not Processed"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center""><a onclick=""return confirm('Are you sure you want to delete " & rs("product_code") & " ?');"" href='delete_basket.asp?id=" & rs("id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"	
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next
	else
        strDisplayList = "<tr><td colspan=""6"" bgcolor=""white"">No items found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""6"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<p>Total: " & intRecordCount & "</p>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
	call setSearch
    call listBasket
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<h2>Loan Stock Workflow Summary</h2>
<table cellspacing="0" cellpadding="5" width="600">
  <tr align="center">
    <td colspan="7"><div class="alert alert-search">
        <form name="frmSearch" id="frmSearch" action="view_basket.asp?type=search" method="post" onsubmit="searchBasket()">
          <strong>Search by Item Code:</strong>
          <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
          <input type="button" name="btnSearch" value="Search" onclick="searchBasket()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </form>
      </div></td>
  </tr>
  <tr class="innerdoctitle">
    <td align="left">Date Added</td>
    <td>Item Code</td>
    <td>Serial</td>
    <td>LIC</td>
    <td>Status</td>
    <td></td>
  </tr>
  <%= strDisplayList %>
</table>
<h2><a href="http://intranet:89/workflow/Loan_Return/Default.aspx?dc=<%= Request("account") %>" target="_blank">Proceed to Workflow</a></h2>
</body>
</html>