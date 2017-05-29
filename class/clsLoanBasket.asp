<%
'-----------------------------------------------
' ADD LOAN STOCK TO BASKET
'-----------------------------------------------
function addBasket(strItemCode, strSerialNo, intLIC, strAccountCode)
	dim strSQL

	Call OpenWorkflowDataBase()
		
	strSQL = "INSERT INTO workflow_loan_return_item_list (item_code, serial_number, product_lic, account_code) VALUES ("
	strSQL = strSQL & " '" & Server.HTMLEncode(strItemCode) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strSerialNo) & "',"
	strSQL = strSQL & " '" & intLIC & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strAccountCode) & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Item has been added to your Workflow Basket.</div>"
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOAN STOCK TO BASKET
'-----------------------------------------------
function addItemBasket(strItemCode, strSerialNo, strLIC, strAccountCode, strOrderNo, strOrderLine)
	Dim cmdObj, paraObj
	
	Session("anchorTag") = strOrderNo & "" & strOrderLine
	
    call OpenWorkflowDataBase
	
    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spAddBasket"
    cmdObj.CommandType = AdCmdStoredProc
		
	Set paraObj = cmdObj.CreateParameter("@item_code",AdVarChar,AdParamInput,50, strItemCode)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@serial_number",AdVarChar,AdParamInput,50, strSerialNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@product_lic",AdVarChar,AdParamInput,50, strLIC)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@account_code",AdVarChar,AdParamInput,9, strAccountCode)
	cmdObj.Parameters.Append paraObj	
	Set paraObj = cmdObj.CreateParameter("@order_no",AdInteger,AdParamInput,4, strOrderNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@order_line",AdInteger,AdParamInput,4, strOrderLine)
	cmdObj.Parameters.Append paraObj
	
    On Error Resume Next
        Dim rs
        Dim id
        set rs = cmdObj.Execute
        id = rs(0)
        set rs = nothing
    On error Goto 0
	
    if CheckForSQLError(conn,"Add",MessageText) = TRUE then
        addItemBasket = FALSE
        strMessageText = MessageText
		'strMessageText = err.description
    else
		addItemBasket = TRUE
		Response.Redirect(Request.ServerVariables("HTTP_REFERER") & "#" & Session("anchorTag"))
		'Response.Redirect("confirm.asp")
		strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Item has been added to your Workflow Basket.</div>"	
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)
	
    call CloseDataBase
end function
%>