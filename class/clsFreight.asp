<%
function setFreightSessionVariables
	Session("strPickupName")		= Trim(Request.Form("txtPickupName"))
	Session("strPickupContact")		= Trim(Request.Form("txtPickupContact"))
	Session("strPickupPhone")		= Trim(Request.Form("txtPickupPhone"))
	Session("strPickupAddress")		= Trim(Request.Form("txtPickupAddress"))
	Session("strPickupCity")		= Trim(Request.Form("txtPickupCity"))
	Session("strPickupState")		= Trim(Request.Form("cboPickupState"))
	Session("intPickupPostcode")	= Trim(Request.Form("txtPickupPostcode"))
	Session("strPickupComments")	= Trim(Request.Form("txtPickupComments"))
				
	Session("strReceiverName")		= Trim(Request.Form("txtReceiverName"))
	Session("strReceiverContact")	= Trim(Request.Form("txtReceiverContact"))
	Session("strReceiverPhone")		= Trim(Request.Form("txtReceiverPhone"))
	Session("strReceiverAddress")	= Trim(Request.Form("txtReceiverAddress"))
	Session("strReceiverCity")		= Trim(Request.Form("txtReceiverCity"))
	Session("strReceiverState")		= Trim(Request.Form("cboReceiverState"))
	Session("intReceiverPostcode")	= Trim(Request.Form("txtReceiverPostcode"))
	Session("strReceiverComments")	= Trim(Request.Form("txtReceiverComments"))
end function

function clearFreightSessionVariables
	Session("strPickupName")		= ""
	Session("strPickupContact")		= ""
	Session("strPickupPhone")		= ""
	Session("strPickupAddress")		= ""
	Session("strPickupCity")		= ""
	Session("strPickupState")		= ""
	Session("intPickupPostcode")	= ""
	Session("strPickupComments")	= ""
				
	Session("strReceiverName")		= ""
	Session("strReceiverContact")	= ""
	Session("strReceiverPhone")		= ""
	Session("strReceiverAddress")	= ""
	Session("strReceiverCity")		= ""
	Session("strReceiverState")		= ""
	Session("intReceiverPostcode")	= ""
	Session("strReceiverComments")	= ""
end function

'-----------------------------------------------
' ADD FREIGHT
'-----------------------------------------------
function addFreight(strPickupName, strPickupContact, strPickupPhone, strPickupAddress, strPickupCity, strPickupState, intPickupPostcode, strPickupComments, strReceiverName, strReceiverContact, strReceiverPhone, strReceiverAddress, strReceiverCity, strReceiverState, intReceiverPostcode, strReceiverComments, strCreatedBy)
	Dim cmdObj, paraObj	
	Session("new_freight_id") = ""
	
    call OpenDataBase

    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spAddFreight"
    cmdObj.CommandType = AdCmdStoredProc
	
	Set paraObj = cmdObj.CreateParameter("@pickup_name",AdVarChar,AdParamInput,50,Server.HTMLEncode(strPickupName))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pickup_contact",AdVarChar,AdParamInput,50,Server.HTMLEncode(strPickupContact))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pickup_phone",AdVarChar,AdParamInput,15,Server.HTMLEncode(strPickupPhone))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pickup_address",AdVarChar,AdParamInput,50,Server.HTMLEncode(strPickupAddress))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pickup_city",AdVarChar,AdParamInput,50,Server.HTMLEncode(strPickupCity))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pickup_state",AdVarChar,AdParamInput,8,strPickupState)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pickup_postcode",AdInteger,AdParamInput,4,intPickupPostcode)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@pickup_comments",AdVarChar,AdParamInput,120,Server.HTMLEncode(strPickupComments))
	cmdObj.Parameters.Append paraObj
	
	Set paraObj = cmdObj.CreateParameter("@receiver_name",AdVarChar,AdParamInput,50,Server.HTMLEncode(strReceiverName))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@receiver_contact",AdVarChar,AdParamInput,50,Server.HTMLEncode(strReceiverContact))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@receiver_phone",AdVarChar,AdParamInput,15,Server.HTMLEncode(strReceiverPhone))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@receiver_address",AdVarChar,AdParamInput,50,Server.HTMLEncode(strReceiverAddress))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@receiver_city",AdVarChar,AdParamInput,50,Server.HTMLEncode(strReceiverCity))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@receiver_state",AdVarChar,AdParamInput,8,strReceiverState)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@receiver_postcode",AdInteger,AdParamInput,4,intReceiverPostcode)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@receiver_comments",AdVarChar,AdParamInput,120,Server.HTMLEncode(strReceiverComments))
	cmdObj.Parameters.Append paraObj
	
	Set paraObj = cmdObj.CreateParameter("@created_by",AdVarChar,AdParamInput,50,strCreatedBy)
	cmdObj.Parameters.Append paraObj

    On Error Resume Next
        Dim rs
        Dim id
        set rs = cmdObj.Execute
        id = rs(0)
		
		Session("new_freight_id") = rs("new_freight_id")
		
        set rs = nothing		
		'response.Write cmdObj.Execute		
    On error Goto 0

    if CheckForSQLError(conn,"Add",MessageText) = TRUE then
        addFreight = FALSE
        strMessageText = MessageText
    else
		addFreight = TRUE
		'Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		Response.Redirect("confirm_freight.asp")
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)

    call CloseDataBase
end function

'-----------------------------------------------
' GET FREIGHT
'-----------------------------------------------
Function getFreight(intFreightID)
	Dim strSQL
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * "
	strSQL = strSQL & "	USR.username, USR.firstname, USR.lastname, USR.storename, USR.dealer_code, USR.branch, "
	strSQL = strSQL & "	USR.address, USR.city, USR.state, USR.postcode, USR.phone "
	strSQL = strSQL & "		FROM mpd_freight "
	strSQL = strSQL & "			LEFT JOIN tbl_users USR ON GRA.created_by = USR.user_id "
	strSQL = strSQL & "		WHERE created_by = " & session("UsrUserID") & " AND gra_id = " & intFreightID

	rs.Open strSQL, conn

	'response.write strSQL

    if not DB_RecSetIsEmpty(rs) Then
		session("username") 			= rs("username")
		session("firstname") 			= rs("firstname")
		session("lastname") 			= rs("lastname")
		session("storename") 			= rs("storename")
		session("dealer_code") 			= rs("dealer_code")
		session("branch") 				= rs("branch")
		session("address") 				= rs("address")
		session("city") 				= rs("city")
		session("state") 				= rs("state")
		session("postcode") 			= rs("postcode")
		session("phone") 				= rs("phone")
		session("model_no") 			= rs("model_no")
		session("serial_no") 			= rs("serial_no")
		session("invoice_no") 			= rs("invoice_no")
		session("invoice_date") 		= rs("invoice_date")
		session("date_purchased") 		= rs("date_purchased")
		session("claim_no") 			= rs("claim_no")
		session("order_no") 			= rs("order_no")
		session("reason") 				= rs("reason")
		session("fault") 				= rs("fault")
		session("test_performed") 		= rs("test_performed")
		session("accessories") 			= rs("accessories")
		session("packaging") 			= rs("packaging")
		session("gra_no") 				= rs("gra_no")		
		session("status") 				= rs("status")
		session("date_created") 		= rs("date_created")
		session("created_by") 			= rs("created_by")
		session("date_modified") 		= rs("date_modified")
		session("modified_by") 			= rs("modified_by")
		session("comments") 			= rs("comments")
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' UPDATE FREIGHT
'-----------------------------------------------
function updateFreight(intID, strModelNo, strSerialNo, strInvoiceNo, strDatePurchased, strClaimNo, strOrderNo, intReason, strFault, strTests, intAccessories, intPackaging, intModifiedBy)
    Dim cmdObj, paraObj

    call OpenDataBase

    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spUpdateReturn"
    cmdObj.CommandType = AdCmdStoredProc
   
	Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4, intID)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,20, strModelNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,20, strSerialNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,20, strInvoiceNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,20, strDatePurchased)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,20, strClaimNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,20, strOrderNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4, intReason)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,100,strFault)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,100, strTests)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,2, intAccessories)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,2, intPackaging)
	cmdObj.Parameters.Append paraObj	
	Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4, intModifiedBy)
	cmdObj.Parameters.Append paraObj

    On Error Resume Next
    cmdObj.Execute

	response.Write cmdObj.Execute
    'On error Goto 0

    if CheckForSQLError(conn,"Update",strMessageText) = TRUE then
	    updateCustomer = FALSE
    else
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
        'response.write "success!!"
		updateCustomer = TRUE
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)

    call CloseDataBase
end function

'-----------------------------------------------
' DELETE RETURN
'-----------------------------------------------
function deleteReturn(intGraID)
	dim strSQL
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
			
	strSQL = "DELETE FROM yma_gra WHERE gra_id = " & intGraID
	
	rs.Open strSQL, conn
	
	Set rs = nothing
	
	if err <> 0 then
		strMessageText = err.description
	else
		Response.Redirect("home.asp")
	end if
	
    call CloseDataBase()
end function


%>