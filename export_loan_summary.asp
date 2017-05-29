<%@ Language=VBScript %>
<!--#include file="../include/connection_base.asp " -->
<%
dim rs
dim strSQL
dim strSort

strSort 	= trim(request("sort"))

dim strTodayDate
strTodayDate = FormatDateTime(Date())

if strSort = "" then
	strSort = "account"
end if

Call OpenBaseDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = strSQL & "SELECT TRIM(B9URKC) as account_code, TRIM(Y1KOM1) as account_name, SUM(B9SKSU - B9AHEN) AS total_qty, "
	strSQL = strSQL & "	 SUM(case when B9STJN <> '00' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) else 0 end) AS total_lic "
	strSQL = strSQL & "	FROM BF9EP "
	strSQL = strSQL & "		INNER JOIN EF2SP ON B9SOSC = E2SOSC "
	strSQL = strSQL & "		INNER JOIN YF1MP ON CONCAT(B9URKC,B9JURC) = Y1KOKC "
	strSQL = strSQL & "	WHERE "
	strSQL = strSQL & "		Y1SKKI <> 'D' "
	'strSQL = strSQL & "				AND (B9URKC LIKE '%" & UCASE(strSearch) & "%' "
	'strSQL = strSQL & "				OR Y1KOM1 LIKE '%" & UCASE(strSearch) & "%') "
	strSQL = strSQL & "		AND B9AHEN < B9SKSU "
	strSQL = strSQL & "	 	AND (E2NGTY = "
	strSQL = strSQL & "	(SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only) "
	strSQL = strSQL & "		AND E2NGTM = "
	strSQL = strSQL & "	(SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)) "
	strSQL = strSQL & "		GROUP BY B9URKC, Y1KOM1 "
	strSQL = strSQL & " 		ORDER BY "
	
	select case strSort
		case "account"
			strSQL = strSQL & "		1"
		case "name"
			strSQL = strSQL & "		2"
		case "qty"
			strSQL = strSQL & "		3 DESC"
		case "lic"
			strSQL = strSQL & "		4 DESC"
	end select	

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=loanstock_summary.xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td align=""left""><strong>Account</strong></td>"
	response.write "<td align=""left""><strong>Name</strong></td>"
	response.write "<td align=""center""><strong>Total Qty</strong></td>"
	response.write "<td align=""right""><strong>Total LIC</strong></td>"
	response.write "</tr>"
	
	while not rs.eof
		response.write "<tr>"
		response.write "<td align=""left"">" & rs.fields("account_code") & "</td>"
		response.write "<td align=""left"">" & rs.fields("account_name") & "</td>"
		response.write "<td align=""center"">" & rs.fields("total_qty") & "</td>"
		response.write "<td align=""right"">" & FormatNumber(rs.fields("total_lic")) & "</td>"
		response.write "</tr>"
		rs.movenext
	wend
		
	response.write "</table>"
end if

Call CloseBaseDataBase()
%>