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

	strSQL = "SELECT * FROM tbl_auction "
	strSQL = strSQL & "	WHERE aucItemCode LIKE '%" & strSearch & "%'"
	strSQL = strSQL & "		OR aucSerialNo LIKE '%" & strSearch & "%'"
	strSQL = strSQL & "		OR aucAccountCode LIKE '%" & strSearch & "%'"
	strSQL = strSQL & "		OR aucAccountName LIKE '%" & strSearch & "%'"	
	strSQL = strSQL & "	ORDER BY " & strSort

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=auction-list.xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td align=""left""><strong>Item code</strong></td>"
	response.write "<td align=""left""><strong>Serial</strong></td>"
	response.write "<td align=""center""><strong>LIC</strong></td>"	
	response.write "<td align=""right""><strong>Account Code</strong></td>"
	response.write "<td align=""right""><strong>Name</strong></td>"
	response.write "<td align=""right""><strong>Title</strong></td>"
	response.write "<td align=""right""><strong>Description</strong></td>"
	response.write "<td align=""right""><strong>Reserve</strong></td>"
	response.write "<td align=""right""><strong>Location</strong></td>"
	response.write "</tr>"
	
	while not rs.eof
		response.write "<tr>"
		
		response.write "<td align=""left"">" & rs.fields("aucItemCode") & "</td>"
		response.write "<td align=""left"">" & rs.fields("aucSerialNo") & "</td>"
		response.write "<td align=""center"">" & FormatNumber(rs.fields("aucLIC")) & "</td>"		
		response.write "<td align=""right"">" & rs.fields("aucAccountCode") & "</td>"
		response.write "<td align=""center"">" & rs.fields("aucAccountName") & "</td>"
		response.write "<td align=""right"">" & rs.fields("aucItemTitle") & "</td>"
		response.write "<td align=""right"">" & rs.fields("aucDescription") & "</td>"
		response.write "<td align=""right"">" & rs.fields("aucReservePrice") & "</td>"
		response.write "<td align=""right"">" & rs.fields("aucLocation") & "</td>"
		
		response.write "</tr>"
		
		intRecordCount = intRecordCount + 1
		rs.movenext
	wend
	response.write "<tr>"
	response.write "<td colspan=""9"" align=""center"">Total: " & intRecordCount & " items</td>"
	response.write "</tr>"
	response.write "</table>"
end if

Call CloseDataBase()
%>