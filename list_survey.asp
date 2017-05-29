<!--#include file="../include/connection_it.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Staff Launch Survey</title>
<link REL="stylesheet" HREF="include/stylesheet.css" TYPE="text/css">
<script language="JavaScript" type="text/javascript">
function searchSurvey(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strStatus 		= document.forms[0].cboStatus.value;
    document.location.href = 'list_survey.asp?type=search&txtSearch=' + strSearch;	
}
    
function resetSearch(){
	document.location.href = 'list_survey.asp?type=reset';    
}  
</script>
</head>
<body>
<%
sub setSearch	
	select case trim(request("type"))
		case "reset"
			session("adm_survey_search") = ""
		case "search"
			session("adm_survey_search") = server.htmlencode(trim(Request("txtSearch")))
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
	
	strSQL = "SELECT * FROM tbl_survey "
	strSQL = strSQL & " WHERE (created_by LIKE '%" & session("adm_survey_search") & "%') "
	strSQL = strSQL & "	ORDER BY survey_id"
			
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

			strDisplayList = strDisplayList & "<td align=""center"">" & rs("created_by") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & rs("overall") & "</td>"
			Select Case rs("overall")
				case 1
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""red""><font color=""white""><span title=""Overall"">" & rs("overall") & "</span></font></td>"
				case 2
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""orange""><font color=""white""><span title=""Overall"">" & rs("overall") & "</span></font></td>"				
				case 3
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""yellow""><span title=""Overall"">" & rs("overall") & "</span></td>"
				case 4
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""blue""><font color=""white""><span title=""Overall"">" & rs("overall") & "</span></font></td>"	
				case 5
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""green""><font color=""white""><span title=""Overall"">" & rs("overall") & "</span></font></td>"				
				case else
					strDisplayList = strDisplayList & "<td align=""center""><span title=""Overall"">" & rs("overall") & "</span></td>"	
			end select
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("overall_comments") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & rs("demo") & "</td>"
			Select Case rs("demo")
				case 1
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""red""><font color=""white""><span title=""Overall"">" & rs("demo") & "</span></font></td>"
				case 2
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""orange""><font color=""white""><span title=""Overall"">" & rs("demo") & "</span></font></td>"				
				case 3
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""yellow""><span title=""Overall"">" & rs("demo") & "</span></td>"
				case 4
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""blue""><font color=""white""><span title=""Overall"">" & rs("demo") & "</span></font></td>"	
				case 5
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""green""><font color=""white""><span title=""Overall"">" & rs("demo") & "</span></font></td>"				
				case else
					strDisplayList = strDisplayList & "<td align=""center""><span title=""Overall"">" & rs("demo") & "</span></td>"	
			end select
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("demo_comments") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & rs("focus") & "</td>"
			Select Case rs("focus")
				case 1
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""red""><font color=""white""><span title=""Overall"">" & rs("focus") & "</span></font></td>"
				case 2
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""orange""><font color=""white""><span title=""Overall"">" & rs("focus") & "</span></font></td>"				
				case 3
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""yellow""><span title=""Overall"">" & rs("focus") & "</span></td>"
				case 4
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""blue""><font color=""white""><span title=""Overall"">" & rs("focus") & "</span></font></td>"	
				case 5
					strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""green""><font color=""white""><span title=""Overall"">" & rs("focus") & "</span></font></td>"				
				case else
					strDisplayList = strDisplayList & "<td align=""center""><span title=""Overall"">" & rs("question_1") & "</span></td>"	
			end select
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("focus_comments") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("future_comments") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & rs("date_created") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next
	else
        strDisplayList = "<tr><td colspan=""8"" align=""center"">No records found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""8"" class=""recordspaging"">"
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
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td valign="top" class="maincontent"><h2>Staff Launch Survey</h2>
      <div class="alert alert-search">
        <form name="frmSearch" id="frmSearch" action="list_survey.asp?type=search" method="post" onsubmit="searchSurvey()">
          <strong>Search by username:</strong>
          <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
          <input type="button" name="btnSearch" value="Search" onclick="searchSurvey()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </form>
      </div></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="4" class="database_records" width="100%">
        <tr class="innerdoctitle" align="center">
          <td>Name</td>
          <td>Overall</td>
          <td>Overall Comments</td>
          <td>Demo</td>
          <td>Demo Comments</td>
          <td>Group</td>
          <td>Group Comments</td>
          <td>Future Comments</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>