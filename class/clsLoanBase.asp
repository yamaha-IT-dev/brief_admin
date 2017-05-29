<%
'-----------------------------------------------
' GET LOAN STOCK FROM BASE
'-----------------------------------------------
Function getLoanStock(intOrderNo, intOrderLine)
	dim strSQL
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

    call OpenBaseDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = strSQL & "SELECT B9JUNO AS order_number, B9JUGY AS order_line, "
	strSQL = strSQL & " B9SKNO AS shipment_no, B9SKGY AS ship_line, B9URKC AS account, B9SCSS AS warehouse, "
	strSQL = strSQL & " B9GREG AS item_group, B9SOSC AS product, B9SKSU - B9AHEN AS qty, "
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
	strSQL = strSQL & " INNER JOIN EF2SP ON B9SOSC = E2SOSC "
	strSQL = strSQL & " WHERE B9AHEN < B9SKSU "
	strSQL = strSQL & "		AND B9JUNO = '" & Trim(intOrderNo) & "' "
	strSQL = strSQL & "		AND B9JUGY = '" & Trim(intOrderLine) & "' "
	strSQL = strSQL & " 	AND (E2NGTY = "
	strSQL = strSQL & "	 (SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)"
	strSQL = strSQL & "		AND E2NGTM = "
	strSQL = strSQL & "	 (SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only))"
	
	'response.Write strSQL
	
	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("loan_account") 		= rs("account")
		session("loan_product") 		= rs("product")	
		session("loan_serial_no") 		= rs("serial_no")
		session("loan_qty") 			= rs("qty")
		session("loan_lic") 			= rs("lic")		
		session("loan_date") 			= rs("loan_date")
		session("loan_first_expiry") 	= DateAdd("m", 3, rs("loan_date"))
		session("loan_final_expiry") 	= DateAdd("m", 6, rs("loan_date"))		
		session("loan_date_diff") 		= DateDiff("d",session("loan_date"), strTodayDate)
		session("loan_shipment_no") 	= rs("shipment_no")
    end if

    call CloseDataBase()
end function
%>