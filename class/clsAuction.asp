<%
'-----------------------------------------------
' ADD LOAN STOCK TO AUCTION
'-----------------------------------------------
function addAuction(strItemCode, strSerialNo, strLIC, strAccountCode, strAccountName, strOrderNo, strOrderLine, strCreatedBy)
	Dim cmdObj, paraObj
	
    call OpenDataBase
	
    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spAddAuction"
    cmdObj.CommandType = AdCmdStoredProc
		
	Set paraObj = cmdObj.CreateParameter("@item_code",AdVarChar,AdParamInput,50, strItemCode)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@serial_number",AdVarChar,AdParamInput,50, strSerialNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@product_lic",AdVarChar,AdParamInput,50, strLIC)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@account_code",AdVarChar,AdParamInput,9, strAccountCode)
	cmdObj.Parameters.Append paraObj	
	Set paraObj = cmdObj.CreateParameter("@account_name",AdVarChar,AdParamInput,50, strAccountName)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@order_no",AdInteger,AdParamInput,4, strOrderNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@order_line",AdInteger,AdParamInput,4, strOrderLine)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@created_by",AdVarChar,AdParamInput,50, strCreatedBy)
	cmdObj.Parameters.Append paraObj
	
    On Error Resume Next
        Dim rs
        Dim id
        set rs = cmdObj.Execute
        id = rs(0)
        set rs = nothing
    On error Goto 0
	
    if CheckForSQLError(conn,"Add",MessageText) = TRUE then
        addAuction = FALSE
        strMessageText = MessageText
		'strMessageText = err.description
    else
		addAuction = TRUE
		Response.Redirect("confirm.asp")		
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)
	
    call CloseDataBase
end function

'----------------------------------------------------------------------------------------
' UPDATE AUCTION
'----------------------------------------------------------------------------------------
Function updateAuction(aucID, aucItemTitle, aucDescription, aucReservePrice, aucLocation, aucModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_auction SET "
	strSQL = strSQL & "aucItemTitle = '" & Server.HTMLEncode(aucItemTitle) & "',"
	strSQL = strSQL & "aucDescription = '" & Server.HTMLEncode(aucDescription) & "',"
	strSQL = strSQL & "aucReservePrice = '" & Server.HTMLEncode(aucReservePrice) & "',"
	strSQL = strSQL & "aucLocation = '" & Server.HTMLEncode(aucLocation) & "',"
	strSQL = strSQL & "aucDateModified = GetDate(),"
	strSQL = strSQL & "aucModifiedBy = '" & Trim(aucModifiedBy) & "' WHERE aucID = " & aucID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	'On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The Auction Item has been updated.</div>"
	end if

	Call CloseDataBase()
end function
%>