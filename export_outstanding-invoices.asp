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

if session("outstanding_start_year") = "" then
	session("outstanding_start_year") = "2015"
end if
	
if session("outstanding_end_year") = "" then
	session("outstanding_end_year") = "2015"
end if

    strSQL = strSQL & "SELECT G1SKYC AS DealerCode, Y1KOM1 AS DealerName, B.G1UKNO AS InvoiceNo, G1UKGN AS Line, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD, B.G1TSYK AS CreditFlag,"
    strSQL = strSQL & " CONCAT(CONCAT(CONCAT(CONCAT(G1SHKD,'/'),G1SHKM),'/'),G1SHKY) AS DueDate,"
    strSQL = strSQL & " SUM(G1MKZB) AS Amount, B6AHNO "
    strSQL = strSQL & " FROM GF1EP B INNER JOIN YF1MP ON Y1KOKC = G1SKYC "
	strSQL = strSQL & " INNER JOIN (SELECT DISTINCT G1UKNO, CASE WHEN G1TSYK = 1 THEN '' ELSE B6AHNO END AS B6AHNO, G1TSYK FROM GF1EP LEFT JOIN BF6EP ON G1UKNO = B6INNO where b6ingy <> 999 GROUP BY G1UKNO, B6AHNO, G1TSYK) A on B.G1UKNO = A.G1UKNO AND B.G1TSYK = A.G1TSYK "
    strSQL = strSQL & " WHERE G1SKKI <> 'D' AND Y1SKKI <> 'D' "
    strSQL = strSQL & " AND G1SKYC = '" & Ucase(session("outstanding_search")) & "' "
    strSQL = strSQL & " AND G1KSNO = 0 AND G1SHKY <> 0 "
    'strSQL = strSQL & " AND G1SHKY * 10000 + G1SHKM * 100 + G1SHKD BETWEEN 20150101 AND 20150630 "
    strSQL = strSQL & " AND G1SHKY * 10000 + G1SHKM * 100 + G1SHKD BETWEEN " & session("outstanding_start_year") & session("outstanding_start_month") & session("outstanding_start_date") & " AND " & session("outstanding_end_year") & session("outstanding_end_month") & session("outstanding_end_date") & " "
    strSQL = strSQL & " GROUP BY G1SKYC, Y1KOM1, B.G1UKNO, G1UKGN, B.G1TSYK, Y1YGKG, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD, B6AHNO, B.G1TSYK "
    strSQL = strSQL & " ORDER BY G1SHKY, G1SHKM, G1SHKD"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=outstanding-invoices-" & Ucase(session("outstanding_search")) & ".xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td><strong>Dealer Code</strong></td>"
	response.write "<td><strong>Dealer Name</strong></td>"
	response.write "<td><strong>Order No</strong></td>"
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
		response.write "<td>" & rs.fields("b6ahno") & "</td>"
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
	response.write "<td colspan=""6""></td><td><b>" & FormatNumber(intTotalCredit) & "</b></td>"
	response.write "<td><b>" & FormatNumber(intTotalDebit) & "</b></td>"
	response.write "</tr>"
	response.write "</table>"
end if

Call CloseBaseDataBase()
%>