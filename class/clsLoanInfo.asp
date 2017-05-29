<%
'-----------------------------------------------
' ADD LOAN STOCK INFO
'-----------------------------------------------
function addLoanStockInfo(stockOrderNo, stockOrderLine, stockLocation, stockSerialNo, stockComments, stockCreatedBy)
	dim strSQL

	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_loan_stock (stockOrderNo, stockOrderLine, stockLocation, stockSerialNo, stockComments, stockCreatedBy) VALUES ("
	strSQL = strSQL & " '" & Server.HTMLEncode(stockOrderNo) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(stockOrderLine) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(stockLocation) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(stockSerialNo) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(stockComments) & "',"
	strSQL = strSQL & " '" & Trim(stockCreatedBy) & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The Loan Stock Info has been added.</div>"
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOAN STOCK INFO
'-----------------------------------------------
function addLoanStockToBasket(strItemCode, strSerialNo, intLIC, strLocation, strAccountCode)
	dim strSQL

	Call OpenWorkflowDataBase()
		
	strSQL = "INSERT INTO workflow_loan_return_item_list (item_code, serial_number, product_lic, loan_stock_location, account_code) VALUES ("
	strSQL = strSQL & " '" & Server.HTMLEncode(strItemCode) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strSerialNo) & "',"
	strSQL = strSQL & " '" & intLIC & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strLocation) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strAccountCode) & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The item has been added to your Workflow Basket.</div>"
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' LIST LOAN STOCK INFO
'-----------------------------------------------
function listLoanStockInfo(stockOrderNo, stockOrderLine)
    dim strSQL
	dim intRecordCount	
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_loan_stock "
	strSQL = strSQL & "	WHERE stockOrderNo = '" & stockOrderNo & "' "
	strSQL = strSQL & "		AND stockOrderLine = '" & stockOrderLine & "' "
	strSQL = strSQL & "	ORDER BY stockSerialNo"
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	
	session("loanstock_info_record_count") = rs.recordcount
	
    strLoanStockInfoList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
		For intRecord = 1 To rs.PageSize			
			strLoanStockInfoList = strLoanStockInfoList & "<tr>"
			strLoanStockInfoList = strLoanStockInfoList & "	<form method=""post"" name=""form_update_stock_info"" id=""form_update_stock_info"" onsubmit=""return validateUpdateStockInfoForm(this)"">"
			strLoanStockInfoList = strLoanStockInfoList & "	<input type=""hidden"" name=""action"" value=""Update"">"
			strLoanStockInfoList = strLoanStockInfoList & "	<input type=""hidden"" name=""stockID"" value=""" & rs("stockID") & """>"
			strLoanStockInfoList = strLoanStockInfoList & "<td align=""center""><input type=""text"" id=""txtUpdateLocation"" name=""txtUpdateLocation"" maxlength=""15"" size=""18"" value=""" & Trim(rs("stockLocation")) & """ ></td>"
			strLoanStockInfoList = strLoanStockInfoList & "<td align=""center""><input type=""text"" id=""txtUpdateSerialNo"" name=""txtUpdateSerialNo"" maxlength=""12"" size=""15"" value=""" & Trim(rs("stockSerialNo")) & """ ></td>"
			strLoanStockInfoList = strLoanStockInfoList & "<td align=""center""><input type=""text"" id=""txtUpdateComments"" name=""txtUpdateComments"" maxlength=""50"" size=""20"" value=""" & Trim(rs("stockComments")) & """ ></td>"
			strLoanStockInfoList = strLoanStockInfoList & "<td align=""center""><input type=""submit"" value=""Save"" /></td>"
			strLoanStockInfoList = strLoanStockInfoList & "</form>"
			'strLoanStockInfoList = strLoanStockInfoList & "<td align=center>" & trim(rs("stockCreatedBy")) & " - " & FormatDateTime(rs("stockDateCreated"),2) & "</td>"			
			strLoanStockInfoList = strLoanStockInfoList & "<td align=""center"">"
			strLoanStockInfoList = strLoanStockInfoList & "	<form method=""post"" name=""form_basket_stock_info"" id=""form_basket_stock_info"" onsubmit=""return validateBasketStockInfoForm(this)"">"
			strLoanStockInfoList = strLoanStockInfoList & "		<input type=""hidden"" name=""action"" value=""Basket"">"
			strLoanStockInfoList = strLoanStockInfoList & "		<input type=""hidden"" name=""stockLocation"" value=""" & Trim(rs("stockLocation")) & """>"
			strLoanStockInfoList = strLoanStockInfoList & "		<input type=""hidden"" name=""stockSerialNo"" value=""" & Trim(rs("stockSerialNo")) & """>"
			strLoanStockInfoList = strLoanStockInfoList & "		<input type=""submit"" value=""Add to Basket"" />"
			strLoanStockInfoList = strLoanStockInfoList & "	</form>"
			strLoanStockInfoList = strLoanStockInfoList & "</td>"
			strLoanStockInfoList = strLoanStockInfoList & "<td align=""center""><a onclick=""return confirm('Are you sure you want to delete " & rs("stockLocation") & " ?');"" href='delete_loanstock.asp?id=" & rs("stockID") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strLoanStockInfoList = strLoanStockInfoList & "</tr>"
			
			rs.movenext
			
			If rs.EOF Then Exit For
		next
	else
        strLoanStockInfoList = "<tr><td colspan=""6"">&nbsp;</td></tr>"
	end if
	
	strLoanStockInfoList = strLoanStockInfoList & "<tr>"
	
    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE LOAN STOCK INFO
'----------------------------------------------------------------------------------------
Function updateLoanStockInfo(stockID, stockLocation, stockSerialNo, stockComments, stockModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_stock SET "
	strSQL = strSQL & "stockLocation = '" & Server.HTMLEncode(stockLocation) & "',"
	strSQL = strSQL & "stockSerialNo = '" & Server.HTMLEncode(stockSerialNo) & "',"
	strSQL = strSQL & "stockComments = '" & Server.HTMLEncode(stockComments) & "',"
	strSQL = strSQL & "stockDateModified = GetDate(),"
	strSQL = strSQL & "stockModifiedBy = '" & Trim(stockModifiedBy) & "' WHERE stockID = " & stockID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	'On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The Loan Stock Info has been updated.</div>"
	end if

	Call CloseDataBase()
end function


%>