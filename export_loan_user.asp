<%@ Language=VBScript %>
<!--#include file="../include/connection_base.asp " -->
<%
dim rs
dim strSQL
dim strSearch
dim strAccount
dim intYear
dim intMonth
dim strSort
dim strDays

strSearch 	= trim(request("search"))
strAccount 	= trim(request("account"))
intYear 	= trim(request("year"))
intMonth 	= trim(request("month"))
strSort 	= trim(request("sort"))

dim strTodayDate
strTodayDate = FormatDateTime(Date())

if intSort = "" then
	intSort = "oldest"
end if

dim intTotalLIC
intTotalLIC = 0

dim intTotalQty
intTotalQty = 0

dim intRecordCount
intRecordCount = 0

Call OpenBaseDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = strSQL & "SELECT B9JUNO AS order_no, B9JUGY AS order_line, "
	strSQL = strSQL & " B9SKNO AS shipment_no, B9SKGY AS ship_line, B9URKC AS account_code, B9SCSS AS warehouse, "
	strSQL = strSQL & " B9GREG AS item_group, B9SOSC AS item_code, B9SKSU - B9AHEN AS loan_qty, Y1KOM1 as account_name, "
	strSQL = strSQL & "	B9SKJY, "
	strSQL = strSQL & "	B9SKJM, "
	strSQL = strSQL & "	B9SKJD, "
	strSQL = strSQL & " RIGHT('0' || B9SKJD,2)|| '/' || RIGHT ('0' || B9SKJM,2) || '/' || B9SKJY as loan_date, "
	strSQL = strSQL & " case "
	strSQL = strSQL & " 	when B9STJN <> '00' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) "
	strSQL = strSQL & " 		else 0 "
	strSQL = strSQL & " 	end AS lic, B9SIBN AS serial_no, "
	strSQL = strSQL & " B9ASFN AS comment "
	strSQL = strSQL & " FROM BF9EP "
	strSQL = strSQL & " 	INNER JOIN EF2SP ON B9SOSC = E2SOSC "
	strSQL = strSQL & " 	INNER JOIN YF1MP ON CONCAT(B9URKC,B9JURC) = Y1KOKC "
	strSQL = strSQL & " WHERE Y1SKKI <> 'D' AND B9AHEN < B9SKSU "
	strSQL = strSQL & "				AND (B9SOSC LIKE '%" & UCASE(strSearch) & "%' "
	strSQL = strSQL & "					OR B9JUNO LIKE '%" & UCASE(strSearch) & "%' "
	strSQL = strSQL & "					OR B9SIBN LIKE '%" & UCASE(strSearch) & "%' "
	strSQL = strSQL & "					OR B9SKNO LIKE '%" & UCASE(strSearch) & "%') "
	strSQL = strSQL & "				AND B9URKC LIKE '%" & UCASE(strAccount) & "%' "
	strSQL = strSQL & "				AND B9SKJY LIKE '%" & intYear & "%' "
	strSQL = strSQL & "				AND B9SKJM LIKE '%" & intMonth & "%' "
	strSQL = strSQL & " AND (E2NGTY = "
	strSQL = strSQL & "	 (SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)"
	strSQL = strSQL & "	AND E2NGTM = "
	strSQL = strSQL & "	 (SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only))"
	strSQL = strSQL & "		ORDER BY "	
	
	select case intSort
		case "oldest"
			strSQL = strSQL & "		B9SKJY ASC, B9SKJM ASC, B9SKJD ASC"			
		case "latest"
			strSQL = strSQL & "		B9SKJY DESC, B9SKJM DESC, B9SKJD DESC"
		case "product"
			strSQL = strSQL & "		item_code"
		case "expensive"
			strSQL = strSQL & "		lic DESC"
		case "cheapest"
			strSQL = strSQL & "		lic"
		case "serial"
			strSQL = strSQL & "		serial_no"
		case "order"
			strSQL = strSQL & "		order_no"
		case "shipment"
			strSQL = strSQL & "		shipment_no"
	end select	

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=loanstock_" & strAccount & ".xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td align=""left""><strong>Account</strong></td>"
	response.write "<td align=""left""><strong>Name</strong></td>"
	response.write "<td align=""left""><strong>Item code</strong></td>"
	response.write "<td align=""left""><strong>Serial</strong></td>"
	response.write "<td align=""center""><strong>Qty</strong></td>"
	response.write "<td align=""center""><strong>LIC</strong></td>"	
	response.write "<td align=""right""><strong>Loan date</strong></td>"
	response.write "<td align=""right""><strong>Day count</strong></td>"
	response.write "</tr>"
	
	while not rs.eof
		strDays = DateDiff("d",rs("loan_date"), strTodayDate)
		intTotalLIC = intTotalLIC + Cint(rs("lic"))
		intTotalQty = intTotalQty + Cint(rs("loan_qty"))
		response.write "<tr>"
		response.write "<td align=""left"">" & rs.fields("account_code") & "</td>"
		response.write "<td align=""left"">" & rs.fields("account_name") & "</td>"
		response.write "<td align=""left"">" & rs.fields("item_code") & "</td>"
		response.write "<td align=""left"">" & rs.fields("serial_no") & "</td>"
		response.write "<td align=""center"">" & rs.fields("loan_qty") & "</td>"
		response.write "<td align=""center"">" & FormatNumber(rs.fields("lic")) & "</td>"		
		response.write "<td align=""right"">" & rs.fields("loan_date") & "</td>"
		response.write "<td align=""right"">" & strDays & "</td>"
		response.write "</tr>"
		
		intRecordCount = intRecordCount + 1
		rs.movenext
	wend
	response.write "<tr>"
	response.write "<td colspan=""8"" align=""center"">Total LIC: $" & FormatNumber(intTotalLIC) & "</td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td colspan=""8"" align=""center"">Total: " & intTotalQty & " stocks</td>"
	response.write "</tr>"
	response.write "</table>"
end if

Call CloseBaseDataBase()
%>