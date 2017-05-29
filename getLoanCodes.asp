<!--#include file="../include/connection_base.asp " -->
<%
response.expires=-1

call OpenBaseDataBase()

Dim strSQL
dim rs
dim intRecord

strSQL = "SELECT DISTINCT B9URKC, Y1KOM1 FROM BF9EP  "	
strSQL = strSQL & " INNER JOIN YF1MP ON CONCAT(B9URKC,B9JURC) = Y1KOKC WHERE B9SKKI <> 'D' ORDER BY B9URKC"

set rs = server.CreateObject("ADODB.Recordset")
set rs = conn.execute(strSQL)

strDisplayList = ""

do while not rs.eof
	strDisplayList = strDisplayList & Trim(rs("B9URKC")) & " " & Trim(rs("Y1KOM1")) & ", "
	rs.movenext
loop

strDisplayList = left(strDisplayList,len(strDisplayList)-1)

response.write strDisplayList

call CloseBaseDataBase
%>