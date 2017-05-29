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
<!--#include file="class/clsEmployee.asp " -->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if lt IE 9]>
	<script src="//oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
	<script src="//oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
<![endif]-->
<title>Web Requests</title>
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/style.css">
<script>
function searchBrief(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strDepartment  	= document.forms[0].cboDepartment.value;
	var strUser  		= document.forms[0].cboUser.value;
	var strStatus 		= document.forms[0].cboStatus.value;
	var strSort  		= document.forms[0].cboSort.value;
	
    document.location.href = 'task.asp?type=search&txtSearch=' + strSearch + '&cboDepartment=' + strDepartment + '&cboUser=' + strUser + '&cboStatus=' + strStatus + '&cboSort=' + strSort;
}

function resetSearch(){
	document.location.href = 'task.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("brief_search") 		= ""
			session("brief_department") 	= ""
			session("brief_user") 	= ""
			session("brief_status") 		= ""
			session("brief_sort") 			= ""
			session("brief_initial_page") 	= 1
		case "search"
			session("brief_search") 		= Trim(Request("txtSearch"))
			session("brief_department") 	= Trim(Request("cboDepartment"))
			session("brief_user") 	= Trim(Request("cboUser"))
			session("brief_status") 		= Trim(Request("cboStatus"))
			session("brief_sort") 			= Trim(Request("cboSort"))
			session("brief_initial_page") 	= 1
	end select
end sub

sub displayBrief
	dim strSQL
	
	dim intRecordCount
	
	dim iRecordCount
	iRecordCount = 0    
    		
	dim strTodayDate	
	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100
	
	if session("brief_sort") = "" then
		session("brief_sort") = "project_deadline"
	end if
		
	'if session("brief_department") = "" then
	'	session("brief_department") = "MPD"
	'end if
	
	strSQL = "SELECT * FROM yma_project "
	strSQL = strSQL & "	WHERE (project_created_by LIKE '%" & session("brief_search") & "%' "
	strSQL = strSQL & "			OR project_output_details LIKE '%" & session("brief_search") & "%' "
	strSQL = strSQL & "			OR project_job_no LIKE '%" & session("brief_search") & "%' "
	strSQL = strSQL & "			OR project_title LIKE '%" & session("brief_search") & "%') "
	strSQL = strSQL & "		AND project_department LIKE '%" & session("brief_department") & "%' "
	strSQL = strSQL & "		AND project_created_by LIKE '%" & session("brief_user") & "%' "
	
	if session("brief_status") = "" then
		strSQL = strSQL & "	AND project_status <> '0' "
	else
		strSQL = strSQL & "	AND project_status LIKE '%" & session("brief_status") & "%' "
	end if
	
	strSQL = strSQL & "	ORDER BY " & session("brief_sort")	
	
	rs.Open strSQL, conn

	'Response.Write strSQL

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("brief_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("brief_initial_page") = intpage

			if session("brief_initial_page") < 1 then session("brief_initial_page") = 1
	    case ">"l
		    intpage = Request("intpage") + 1
			session("brief_initial_page") = intpage

			if session("brief_initial_page") > intPageCount then session("brief_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("brief_initial_page") = intpage
    end select
	
    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("brief_initial_page")

		For intRecord = 1 To rs.PageSize			
			strDisplayList = strDisplayList & "<tr>"						
			strDisplayList = strDisplayList & "<td><a href=""update-task.asp?id=" & rs("project_id") & """>" & rs("project_id") & "</a></td>"			
			strDisplayList = strDisplayList & "<td>" & rs("project_created_by") & ", " & FormatDateTime(rs("project_date_created"),1) & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("project_title") & ""
			if DateDiff("d",rs("project_date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=0>"
			end if			
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			Select Case	rs("project_priority")
				case 1
					strDisplayList = strDisplayList & "<font class=""low_font"">Low</font>"
				case 2
					strDisplayList = strDisplayList & "<font class=""medium_font"">Medium</font>"
				case 3
					strDisplayList = strDisplayList & "<font class=""high_font"">High</font>"
			end select
			strDisplayList = strDisplayList & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("project_quote") & "</td>"			
			strDisplayList = strDisplayList & "<td>"	
			select case rs("marketing_manager_approval")
				case 0
					strDisplayList = strDisplayList & "..."
				case 1
					strDisplayList = strDisplayList & "<font color=""green"">Approved</font>"
				case 2
					strDisplayList = strDisplayList & "<font color=""red"">Rejected</font>"	
			end select			
			strDisplayList = strDisplayList & "</td>"			
			strDisplayList = strDisplayList & "<td>"
			Select Case	rs("project_status")
				case 1
					strDisplayList = strDisplayList & "<font color=""red"">Plan</font>"
				case 2
					strDisplayList = strDisplayList & "<font color=""red"">Submitted</font>"
				case 3
					strDisplayList = strDisplayList & "<font color=""red"">Viewed</font>"
				case 4
					strDisplayList = strDisplayList & "<font color=""orange"">Concept</font>"
				case 5
					strDisplayList = strDisplayList & "<font color=""orange"">Draft</font>"
				case 6
					strDisplayList = strDisplayList & "<font color=""orange"">Changes</font>"
				case 7
					strDisplayList = strDisplayList & "<font color=""green"">Pending approval</font>"
				case 8
					strDisplayList = strDisplayList & "<font color=""green"">On-hold</font>"
				case 0
					strDisplayList = strDisplayList & "<font color=""green"">Completed</font>"
				case else
					strDisplayList = strDisplayList & rs("project_status")
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete this web request: - " & rs("project_title") & " ?');"" href='delete-task.asp?id=" & rs("project_id") & "'><img src=""images/icon_trash.png"" border=""0""></a></td>"			
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr><td colspan=""8"">No records found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""8"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""task.asp"" method=""post"">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("brief_initial_page") & ">"

	if session("brief_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("brief_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
	
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strSort & ">"
    strDisplayList = strDisplayList & "<h3>Page: " & session("brief_initial_page") & " to " & intpagecount & "</h3>"
	strDisplayList = strDisplayList & "<h2>Total: <u>" & intRecordCount & "</u> records</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call getEmployeeDetails(session("logged_username"))
	call setSearch
	
	if trim(session("brief_initial_page"))  = "" then
    	session("brief_initial_page") = 1
	end if
	
    call displayBrief
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<div class="container">
  <h2><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add-task.asp">New Task</a></h2>
  <p><img src="images/legend-blue.gif" border="1" /> = updated today</p>
  <form name="frmSearch" id="frmSearch" action="task.asp?type=search" method="post" onsubmit="searchBrief()">
    <input type="text" class="form-control" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" placeholder="Created by / Title / Job no / Details" />
    <select class="form-control" name="cboDepartment" onchange="searchBrief()">
      <option value="">All Departments</option>
      <option <% if session("brief_department") = "MPD" then Response.Write " selected" end if%> value="MPD">All MPD</option>
      <option <% if session("brief_department") = "MPD - PRO" then Response.Write " selected" end if%> value="MPD - PRO">PRO</option>
      <option <% if session("brief_department") = "MPD - TRAD" then Response.Write " selected" end if%> value="MPD - TRAD">TRAD</option>
      <option <% if session("brief_department") = "CA" then Response.Write " selected" end if%> value="CA">CA</option>
      <option <% if session("brief_department") = "YMEC" then Response.Write " selected" end if%> value="YMEC">YMEC</option>
    </select>
    <select name="cboUser" onchange="searchBrief()">
      <option value="">All Users</option>
      <option <% if session("brief_user") = "alex" then Response.Write " selected" end if%> value="alex">Alex</option>
      <option <% if session("brief_user") = "atsuko" then Response.Write " selected" end if%> value="atsuko">Atsuko</option>
      <option <% if session("brief_user") = "carolyn" then Response.Write " selected" end if%> value="carolyn">Carolyn</option>
      <option <% if session("brief_user") = "cameron" then Response.Write " selected" end if%> value="cameron">Cameron</option>
      <option <% if session("brief_user") = "dion" then Response.Write " selected" end if%> value="dion">Dion</option>
      <option <% if session("brief_user") = "eric" then Response.Write " selected" end if%> value="eric">Eric</option>
      <option <% if session("brief_user") = "euan" then Response.Write " selected" end if%> value="euan">Euan</option>
      <option <% if session("brief_user") = "jaclyn" then Response.Write " selected" end if%> value="jaclyn">Jaclyn</option>
      <option <% if session("brief_user") = "jamie" then Response.Write " selected" end if%> value="jamie">Jamie</option>
      <option <% if session("brief_user") = "julian" then Response.Write " selected" end if%> value="julian">Julian</option>
      <option <% if session("brief_user") = "leon" then Response.Write " selected" end if%> value="leon">Leon</option>
      <option <% if session("brief_user") = "mick" then Response.Write " selected" end if%> value="mick">Mick</option>
      <option <% if session("brief_user") = "mattd" then Response.Write " selected" end if%> value="mattd">Matt Dawkins</option>
      <option <% if session("brief_user") = "mattl" then Response.Write " selected" end if%> value="mattl">Matt Livingstone</option>
      <option <% if session("brief_user") = "nathan" then Response.Write " selected" end if%> value="nathan">Nathan</option>
      <option <% if session("brief_user") = "peta" then Response.Write " selected" end if%> value="peta">Peta</option>
    </select>
    <select name="cboStatus" onchange="searchBrief()">
      <option <% if session("brief_status") = "" then Response.Write " selected" end if%> value="">All Status (Exclude Completed)</option>
      <option <% if session("brief_status") = "1" then Response.Write " selected" end if%> value="1" style="color:red">Status: Plan</option>
      <option <% if session("brief_status") = "2" then Response.Write " selected" end if%> value="2" style="color:red">Status: Submitted</option>
      <option <% if session("brief_status") = "3" then Response.Write " selected" end if%> value="3" style="color:red">Status: Viewed</option>
      <option <% if session("brief_status") = "4" then Response.Write " selected" end if%> value="4" style="color:orange">Status: Concept</option>
      <option <% if session("brief_status") = "5" then Response.Write " selected" end if%> value="5" style="color:orange">Status: Draft</option>
      <option <% if session("brief_status") = "6" then Response.Write " selected" end if%> value="6" style="color:orange">Status: Changes</option>
      <option <% if session("brief_status") = "7" then Response.Write " selected" end if%> value="7" style="color:green">Status: Pending approval</option>
      <option <% if session("brief_status") = "8" then Response.Write " selected" end if%> value="8" style="color:green">Status: On-hold</option>
      <option <% if session("brief_status") = "0" then Response.Write " selected" end if%> value="0" style="color:green">Status: Completed</option>
    </select>
    <select name="cboSort" onchange="searchBrief()">
      <option <% if session("brief_sort") = "project_deadline" then Response.Write " selected" end if%> value="project_deadline">Sort by: Deadline</option>
      <option <% if session("brief_sort") = "project_date_created" then Response.Write " selected" end if%> value="project_date_created">Sort by: Date created</option>
      <option <% if session("brief_sort") = "project_created_by" then Response.Write " selected" end if%> value="project_created_by">Sort by: Requested by</option>
    </select>
    <input type="button" name="btnSearch" value="Search" onclick="searchBrief()" />
    <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
  </form>
  <div class="table-responsive">
    <table class="table table-striped">
      <thead>
        <tr>
          <td>Task ID</td>
          <td>Created</td>
          <td>Task Name</td>
          <td>Priority</td>
          <td>Quote</td>
          <td>Marketing Approval</td>
          <td>Status</td>
          <td></td>
        </tr>
      </thead>
      <tbody>
        <%= strDisplayList %>
      </tbody>
    </table>
  </div>
</div>
<script src="//code.jquery.com/jquery.js"></script> 
<script src="bootstrap/js/bootstrap.min.js"></script>
</body>
</html>