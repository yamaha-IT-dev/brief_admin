<%@ Language=VBScript %>
<!--#include file="../include/connection_base.asp " -->
<%
dim rs
dim strSQL

Call OpenBaseDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = "SELECT E1OPEC, E1NSKY, E1NSKM, E1NSKD, E1NSKY * 10000 +  E1NSKM * 100 + E1NSKD AS INV_MOV_DATE, E1SOSC, E1SKNO, "
	strSQL = strSQL & "	((E1AKKB*-2)+1)* E1NSKS AS QTY, E1SISC, E1SOCD"
	strSQL = strSQL & "	FROM EF1BP "
	strSQL = strSQL & "	WHERE ("
	strSQL = strSQL & "			E1SOSC LIKE '%" & Ucase(Session("inventory_search")) & "%' "
	strSQL = strSQL & "			OR E1SKNO LIKE '%" & Session("inventory_search") & "%')"
	if Session("inventory_month") <> "" then
		strSQL = strSQL & " AND E1NSKM = '" & Trim(Session("inventory_month")) & "' "
	end if
	strSQL = strSQL & "		AND E1TRTI = 'AH'"
	strSQL = strSQL & "		AND E1SKKI <> 'D'"
	strSQL = strSQL & "		AND E1NSKY >= 2014"	
	strSQL = strSQL & "		AND E1OPEC LIKE '%" & Session("inventory_user") & "%' "
	strSQL = strSQL & "		AND E1SISC LIKE '%" & Session("inventory_vendor") & "%' "
	strSQL = strSQL & "		AND E1SOCD LIKE '%" & Session("inventory_warehouse") & "%' "
	strSQL = strSQL & "	ORDER BY "
			
	select case Session("inventory_sort")
		case "oldest"
			strSQL = strSQL & "INV_MOV_DATE"
		case "operator"
			strSQL = strSQL & "E1OPEC"
		case "product"
			strSQL = strSQL & "E1SOSC"
		case "shipment"
			strSQL = strSQL & "E1SKNO"
		case "qty"
			strSQL = strSQL & "INV_MOV_QTY DESC"
		case "vendor"
			strSQL = strSQL & "E1SISC"
		case "warehouse"
			strSQL = strSQL & "E1SOCD"	
		case else
			strSQL = strSQL & "INV_MOV_DATE DESC"
	end select

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=warehouse-inventory.xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td><strong>Operator</strong></td>"
	response.write "<td><strong>Invoice Move Date</strong></td>"
	response.write "<td><strong>Product</strong></td>"
	response.write "<td><strong>Shipment</strong></td>"
	response.write "<td><strong>Qty</strong></td>"
	response.write "<td><strong>Vendor</strong></td>"
	response.write "<td><strong>Warehouse</strong></td>"
	response.write "</tr>"
	while not rs.eof
		response.write "<tr>"
		response.write "<td>" & rs.fields("E1OPEC") & "</td>"
		response.write "<td>" & rs.fields("E1NSKD") & "/" & rs.fields("E1NSKM") & "/" & rs.fields("E1NSKY") & "</td>"
		response.write "<td>" & rs.fields("E1SOSC") & "</td>"
		response.write "<td>" & rs.fields("E1SKNO") & "</td>"
		response.write "<td>" & rs.fields("QTY") & "</td>"
		response.write "<td>" & rs.fields("E1SISC") & "</td>"		
		response.write "<td>" & rs.fields("E1SOCD") & "</td>"
		response.write "</tr>"		
		rs.movenext
	wend
	response.write "</table>"
end if

Call CloseBaseDataBase()
%>