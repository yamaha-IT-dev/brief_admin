<%@ Language=VBScript %>
<!--#include file="../include/connection_it.asp " -->
<%
dim rs
dim strSQL
dim strAccount
strAccount 	= trim(request("account"))

dim strTodayDate
strTodayDate = FormatDateTime(Date())

dim intTotalLIC
intTotalLIC = 0

dim intTotalQty
intTotalQty = 0

dim intRecordCount
intRecordCount = 0

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = strSQL & "SELECT DISTINCT order_no, order_line, department, account_code, account_code_ext, account_name, item_code, serial_no, "
	strSQL = strSQL & "		B9AHEN, B9SKSU, lic, qty, loan_year, loan_month, loan_day, loan_date, stockLocation, stockRenewalCounter, stockID, renStatus, renActive, renExpiryDate, renDateCreated "
	strSQL = strSQL & ", product_code, aucItemCode "
	strSQL = strSQL & " FROM OPENQUERY "
	strSQL = strSQL & " (AS400, 'SELECT B9JUNO AS order_no, B9JUGY AS order_line, Y1REGN AS department, B9URKC AS account_code, B9JURC AS account_code_ext, "
	strSQL = strSQL & "			Y1KOM1 AS account_name, B9SOSC AS item_code, B9SIBN AS serial_no, B9SKJY AS loan_year, B9SKJM AS loan_month, B9SKJD AS loan_day, "
	strSQL = strSQL & "			B9SKSU - B9AHEN AS qty, B9AHEN, B9SKSU, "
	strSQL = strSQL & " 		RIGHT(''0'' || B9SKJD,2)|| ''/'' || RIGHT (''0'' || B9SKJM,2) || ''/'' || B9SKJY AS loan_date, "
 	strSQL = strSQL & " 		CASE WHEN B9STJN <> ''00'' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) "
	strSQL = strSQL & " 			ELSE 0 "
	strSQL = strSQL & " 		END AS lic "	
	strSQL = strSQL & " 	FROM BF9EP "
	strSQL = strSQL & "			INNER JOIN EF2SP ON B9SOSC = E2SOSC "
	strSQL = strSQL & " 		INNER JOIN YF1MP ON CONCAT(B9URKC,B9JURC) = Y1KOKC "
	strSQL = strSQL & " WHERE Y1SKKI <> ''D'' AND E2NGTY = "
	strSQL = strSQL & "	(SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)"
	strSQL = strSQL & " 	AND E2NGTM = "
	strSQL = strSQL & "	(SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)"
	strSQL = strSQL & "	')"
	strSQL = strSQL & "		LEFT JOIN tbl_loan_location ON stockOrderNo = order_no AND stockOrderLine = order_line "
	strSQL = strSQL & "		LEFT JOIN (SELECT renOrderNo, renOrderLine, renStatus, renActive, renDateCreated, renExpiryDate FROM tbl_loan_renewal GROUP BY renOrderNo, renOrderLine, renStatus, renActive, renDateCreated, renExpiryDate) AS RENEWAL ON order_no = renOrderNo AND order_line = renOrderLine "
	strSQL = strSQL & "		LEFT JOIN yamaha_workflow..workflow_loan_return_item_list ON yamaha_workflow..workflow_loan_return_item_list.order_number = order_no AND yamaha_workflow..workflow_loan_return_item_list.order_lines = order_line "
	strSQL = strSQL & "		LEFT JOIN tbl_auction ON aucOrderNo = order_no AND aucOrderLine = order_line "
	strSQL = strSQL & " WHERE B9AHEN < B9SKSU and (renActive = 0 or renOrderNo is null) "
	strSQL = strSQL & "		AND account_code LIKE '%" & UCASE(strAccount) & "%' "
	strSQL = strSQL & " ORDER BY loan_year ASC, loan_month ASC, loan_day ASC"	

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=loanstock_" & UCASE(strAccount) & ".xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td align=""left""><strong>Account</strong></td>"
	response.write "<td align=""left""><strong>Name</strong></td>"
	response.write "<td align=""left""><strong>Item code</strong></td>"
	response.write "<td align=""left""><strong>Serial</strong></td>"
	response.write "<td align=""center""><strong>Qty</strong></td>"
	response.write "<td align=""right""><strong>Loan date</strong></td>"
	response.write "<td align=""right""><strong>Location</strong></td>"
	response.write "</tr>"
	
	while not rs.eof
		response.write "<tr>"
		response.write "<td align=""left"">" & rs.fields("account_code") & "</td>"
		response.write "<td align=""left"">" & rs.fields("account_name") & "</td>"
		response.write "<td align=""left"">" & rs.fields("item_code") & "</td>"
		response.write "<td align=""left"">" & rs.fields("serial_no") & "</td>"
		response.write "<td align=""center"">" & rs.fields("qty") & "</td>"	
		response.write "<td align=""right"">" & rs.fields("loan_date") & "</td>"
		response.write "<td align=""right"">" & rs.fields("stockLocation") & "</td>"
		response.write "</tr>"

		rs.movenext
	wend

	response.write "</table>"
end if

Call CloseDataBase()
%>