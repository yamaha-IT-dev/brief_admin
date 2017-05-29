<%
'-----------------------------------------------
' ADD FREIGHT ITEM
'-----------------------------------------------
function addFreightItem(intFreightID, strName, strDetails, intQty, strPallet, intLength, intWidth, intHeight, intWeight, strCreatedBy)
	Dim cmdObj, paraObj
	
    call OpenDataBase

    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spAddFreightItem"
    cmdObj.CommandType = AdCmdStoredProc
	
	Set paraObj = cmdObj.CreateParameter("@freight_id",AdInteger,AdParamInput,4,intFreightID)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@name",AdVarChar,AdParamInput,50,Server.HTMLEncode(strName))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@details",AdVarChar,AdParamInput,50,Server.HTMLEncode(strDetails))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@qty",AdInteger,AdParamInput,4,intQty)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pallet",AdVarChar,AdParamInput,20,strPallet)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@length",AdVarChar,AdParamInput,8,intLength)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@width",AdVarChar,AdParamInput,8,intWidth)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@height",AdVarChar,AdParamInput,8,intHeight)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@weight",AdVarChar,AdParamInput,8,intWeight)
	cmdObj.Parameters.Append paraObj	
	Set paraObj = cmdObj.CreateParameter("@created_by",AdVarChar,AdParamInput,50,strCreatedBy)
	cmdObj.Parameters.Append paraObj

    On Error Resume Next
        Dim rs
        Dim id
        set rs = cmdObj.Execute
        id = rs(0)
        set rs = nothing		
		'response.Write cmdObj.Execute
    On error Goto 0

    if CheckForSQLError(conn,"Add",MessageText) = TRUE then
        addFreightItem = FALSE
        strMessageText = MessageText
    else
		addFreightItem = TRUE
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		'Response.Redirect("home.asp")
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)

    call CloseDataBase
end function

'-----------------------------------------------
' DELETE FREIGHT ITEM
'-----------------------------------------------
function deleteFreightItem(intItemID)
	dim strSQL
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
			
	strSQL = "DELETE FROM mpd_freight WHERE item_id = " & intItemID
	
	rs.Open strSQL, conn
	
	Set rs = nothing
	
	if err <> 0 then
		strMessageText = err.description
	else
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if
	
    call CloseDataBase()
end function


%>