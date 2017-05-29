<%@ Language=VBScript %>
<!--#include file="../include/connection_it.asp " -->
<%
dim rs
dim strSQL
dim strSearch
dim strSort

strSearch 	= trim(request("search"))
strSort 	= trim(request("sort"))

dim strTodayDate
strTodayDate = FormatDateTime(Date())

if strSort = "" then
	strSort = "aucItemCode"
end if

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = "SELECT * FROM yma_project "
	strSQL = strSQL & "	WHERE (project_created_by LIKE '%" & session("brief_search") & "%' "
	strSQL = strSQL & "			OR project_output_details LIKE '%" & session("brief_search") & "%' "
	strSQL = strSQL & "			OR project_job_no LIKE '%" & session("brief_search") & "%' "
	strSQL = strSQL & "			OR project_title LIKE '%" & session("brief_search") & "%') "
	strSQL = strSQL & "		AND project_department LIKE '%" & session("brief_department") & "%' "
	strSQL = strSQL & "		AND project_created_by LIKE '%" & session("brief_user") & "%' AND project_status <> '99' "
	
	if session("brief_status") = "" then
		strSQL = strSQL & "	AND project_status <> '0' "
	else
		strSQL = strSQL & "	AND project_status LIKE '%" & session("brief_status") & "%' "
	end if
	
	strSQL = strSQL & "	ORDER BY " & session("brief_sort")	

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=GD-brief-list.xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td><strong>ID</strong></td>"
	response.write "<td><strong>Created</strong></td>"
	response.write "<td><strong>Dept</strong></td>"
	response.write "<td><strong>Title</strong></td>"	
	response.write "<td><strong>Priority</strong></td>"
	response.write "<td><strong>Print</strong></td>"
	response.write "<td><strong>Online</strong></td>"
	response.write "<td><strong>Quote</strong></td>"
	response.write "<td><strong>Actual Hours</strong></td>"
	response.write "<td><strong>1st Draft</strong></td>"
	response.write "<td><strong>2nd Draft</strong></td>"
	response.write "<td><strong>Print/Publish</strong></td>"
	response.write "<td><strong>Requester</strong></td>"
	response.write "<td><strong>Marketing</strong></td>"
	response.write "<td><strong>Progress</strong></td>"
	response.write "<td><strong>Status</strong></td>"
	response.write "<td><strong>Job no</strong></td>"
	response.write "</tr>"
	
	while not rs.eof
		response.write "<tr>"
		response.write "<td>" & rs.fields("project_id") & "</td>"
		response.write "<td>" & rs.fields("project_created_by") & " - " & rs.fields("project_date_created") & "</td>"
		response.write "<td>" & rs.fields("project_department") & "</td>"
		response.write "<td>" & rs.fields("project_title") & "</td>"
		response.write "<td>"
		select case rs.fields("project_priority")
			case 1
				response.write "Low"	
			case 2
				response.write "Medium"
			case 3
				response.write "High"
		end select
		response.write "</td>"		
		response.write "<td>" & rs.fields("project_output_printed") & "</td>"		
		response.write "<td>" & rs.fields("project_output_web") & "</td>"
		response.write "<td>" & rs.fields("project_quote") & "</td>"
		response.write "<td>" & rs.fields("project_actual_hours") & "</td>"
		response.write "<td>" & rs.fields("project_first_deadline") & "</td>"
		response.write "<td>" & rs.fields("project_second_deadline") & "</td>"
		response.write "<td>" & rs.fields("project_deadline") & "</td>"
		response.write "<td>" & rs.fields("product_manager_approval") & "</td>"
		response.write "<td>" & rs.fields("marketing_manager_approval") & "</td>"
		response.write "<td>" & rs.fields("project_progress") & "</td>"
		response.write "<td>"
		select case rs.fields("project_status")
			case 1
				response.write "Plan"	
			case 2
				response.write "Submitted"
			case 3
				response.write "Viewed"
			case 4
				response.write "Concept"	
			case 5
				response.write "Draft"
			case 6
				response.write "Changes"	
			case 7
				response.write "Pending Approval"	
			case 8
				response.write "On-hold"
			case 0
				response.write "Completed"	
		end select
		response.write "</td>"
		response.write "<td>" & rs.fields("project_job_no") & "</td>"
		response.write "</tr>"
		
		intRecordCount = intRecordCount + 1
		rs.movenext
	wend
	response.write "<tr>"
	response.write "<td colspan=""16"" align=""center"">Total: " & intRecordCount & "</td>"
	response.write "</tr>"
	response.write "</table>"
end if

Call CloseDataBase()
%>