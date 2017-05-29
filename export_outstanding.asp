<%@ Language=VBScript %>
<!--#include file="../include/connection_base.asp " -->
<%
dim rs
dim strSQL
dim strSearch
dim strUser
dim intYear
dim intMonth
dim strSort
dim strDays

strSearch 	= trim(request("search"))
strUser 	= trim(request("user"))
intYear 	= trim(request("year"))
intMonth 	= trim(request("month"))
strSort 	= trim(request("sort"))

dim strTodayDate
strTodayDate = FormatDateTime(Date())

if intSort = "" then
	intSort = "oldest"
end if

dim intTotalDebit
intTotalDebit = 0

dim intTotalCredit
intTotalCredit = 0

dim intRecordCount
intRecordCount = 0

Call OpenBaseDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = strSQL & "SELECT G1SKYC AS DealerCode, Y1KOM1 AS DealerName, G1UKNO AS InvoiceNo, G1UKGN AS Line, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD, G1TSYK AS CreditFlag,"
	strSQL = strSQL & " SUM(G1MKZB) AS Amount"	
	strSQL = strSQL & " FROM GF1EP INNER JOIN YF1MP ON Y1KOKC = G1SKYC WHERE"	
	strSQL = strSQL & " G1SKKI <> 'D' AND Y1SKKI <> 'D' "
	strSQL = strSQL & " AND G1SKYC = '" & Ucase(session("outstanding_search")) & "' "
	strSQL = strSQL & " AND G1KSNO = 0 AND G1SHKY <> 0"
	strSQL = strSQL & " GROUP BY G1SKYC, Y1KOM1, G1UKNO, G1UKGN, G1TSYK, Y1YGKG, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD "
	strSQL = strSQL & " ORDER BY G1SHKY, G1SHKM, G1SHKD"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=outstanding-invoices-" & Ucase(session("outstanding_search")) & ".xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td><strong>Dealer Code</strong></td>"
	response.write "<td><strong>Dealer Name</strong></td>"
	response.write "<td><strong>Invoice</strong></td>"
	response.write "<td><strong>Invoice Date</strong></td>"
	response.write "<td><strong>Due Date</strong></td>"	
	response.write "<td><strong>Credit</strong></td>"
	response.write "<td><strong>Debit</strong></td>"
	response.write "</tr>"
	
	while not rs.eof		
		'intTotalDebit = intTotalDebit + Cint(rs("Amount"))
		'intTotalCredit = intTotalCredit + Cint(rs("Amount"))
		response.write "<tr>"
		response.write "<td>" & rs.fields("DealerCode") & "</td>"
		response.write "<td>" & rs.fields("DealerName") & "</td>"
		response.write "<td>" & rs.fields("InvoiceNo") & ""
		if len(trim(rs.fields("Line"))) > 1 then
			response.write "-" & rs("Line") & ""
		end if
		response.write "</td>"	
		response.write "<td>" & rs.fields("G1SEKD") & "/" & rs.fields("G1SEKM") & "/" & rs.fields("G1SEKY") & "</td>"
		response.write "<td>" & rs.fields("G1SHKD") & "/" & rs.fields("G1SHKM") & "/" & rs.fields("G1SHKY") & "</td>"
		if rs("CreditFlag") = "1" then
			response.write "<td>" & rs.fields("Amount") & "</td>"
			intTotalCredit = intTotalCredit + CCur(rs("Amount"))
		else
			response.write "<td></td>"
		end if
		if rs("CreditFlag") = "0" then
			response.write "<td>" & rs.fields("Amount") & "</td>"
			intTotalDebit = intTotalDebit + CCur(rs("Amount"))
		else
			response.write "<td></td>"
		end if
		response.write "<td>" & strDays & "</td>"
		response.write "</tr>"
		
		intRecordCount = intRecordCount + 1
		rs.movenext
	wend
	response.write "<tr>"
	response.write "<td colspan=""5""></td><td><b>" & FormatNumber(intTotalCredit) & "</b></td>"
	response.write "<td><b>" & FormatNumber(intTotalDebit) & "</b></td>"
	response.write "</tr>"
	response.write "</table>"
end if

Call CloseBaseDataBase()
%>