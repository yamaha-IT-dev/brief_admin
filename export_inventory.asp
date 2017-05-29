<%@ Language=VBScript %>
<!--#include file="../include/connection_base.asp " -->
<%
dim rs
dim strSQL

dim intRecordCount
intRecordCount = 0

Call OpenBaseDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT E1OPEC AS OP, E1NSKY, E1NSKM, E1NSKD, E1NSKY * 10000 +  E1NSKM * 100 + E1NSKD AS INV_MOV_DATE, E1SOSC, E1SKNO AS SHIPMENT, "
strSQL = strSQL & "	((E1AKKB*-2)+1)* E1NSKS AS INV_MOV_QTY, E1SISC AS VENDOR_CODE, E1SOCD AS WAREHOUSE "
strSQL = strSQL & "	FROM EF1BP "
strSQL = strSQL & "	WHERE"
strSQL = strSQL & "		E1TRTI = 'AH'"
strSQL = strSQL & "		AND E1SKKI <> 'D'"
strSQL = strSQL & "		AND E1NSKY >= 2014"	
	
rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=warehouse-inventory.xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"

	response.write "<td><strong>Product</strong></td>"

	response.write "</tr>"
	while not rs.eof
		response.write "<tr>"

		response.write "<td>" & rs.fields("E1SOSC") & "</td>"

		response.write "</tr>"	
		
		intRecordCount = intRecordCount + 1	
		rs.movenext
	wend
	response.write "</table>"
end if

Call CloseBaseDataBase()
%>