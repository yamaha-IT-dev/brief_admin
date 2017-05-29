<%
function newLoanTransfer(strAccountCode, strModelNo, strSerialNo, strOrderNo, strOrderLine)	
	session("newAccountCode") 	= strAccountCode
	session("newModelNo") 		= strModelNo
	session("newSerialNo") 		= strSerialNo
	session("newOrderNo") 		= strOrderNo
	session("newOrderLine")		= strOrderLine

	response.Redirect("add-loan-transfer.asp")
end function

function addTransfer(traAccountCode, traModelNo, traSerialNo, traOrderNo, traOrderLine, traQty, traRecipient, traCreatedBy)
	dim strSQL
	dim strRecipient

	Call OpenDataBase()

	strSQL = "INSERT INTO tbl_loan_transfer (traAccountCode, traModelNo, traSerialNo, traOrderNo, traOrderLine, traQty, traRecipient, traCreatedBy) VALUES ("
	strSQL = strSQL & "'" & Server.HTMLEncode(Ucase(traAccountCode)) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Ucase(traModelNo)) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(traSerialNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(traOrderNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(traOrderLine) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(traQty) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Ucase(traRecipient)) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(traCreatedBy) & "')"

	on error resume next
	conn.Execute strSQL

	response.Write strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields


	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update

		emailFrom 	 = "automailer@gmx.yamaha.com"

		'Select Case traDepartment
		'	case "AV"
		'		emailTo  	= "russell.wykes@music.yamaha.com"
		'	case "PRO"
		'		emailTo  	= "nathan.biggin@music.yamaha.com"
		'	case "TRAD"
		'		emailTo 	= "cameron.tait@music.yamaha.com"
		'	case "YME"
		'		emailTo 	= "carolyn.simonds@music.yamaha.com"
		'end select

		if left(Ucase(traAccountCode), 4) = "7YME" then 
				emailTo 	= "chris.noonan@music.yamaha.com"
		elseif Ucase(traAccountCode) = "7MH000" or Ucase(traAccountCode) = "7AY000" or Ucase(traAccountCode) = "7AD000" or Ucase(traAccountCode) = "7BE001" or Ucase(traAccountCode) = "7CH001" or Ucase(traAccountCode) = "7EO000" _
				or Ucase(traAccountCode) = "7EM001" or Ucase(traAccountCode) = "7FOC00" or Ucase(traAccountCode) = "7GL000" or Ucase(traAccountCode) = "7JW001" or Ucase(traAccountCode) = "7JG000" or Ucase(traAccountCode) = "7AUDW0" _
				or Ucase(traAccountCode) = "7LB000" or Ucase(traAccountCode) = "7MC001" or Ucase(traAccountCode) = "7MMC00" or Ucase(traAccountCode) = "7MT000" or Ucase(traAccountCode) = "7MD000" or Ucase(traAccountCode) = "7ML003" _
				or Ucase(traAccountCode) = "7MIRG0" or Ucase(traAccountCode) = "7NB001" or Ucase(traAccountCode) = "7PW000" or Ucase(traAccountCode) = "7BP000" or Ucase(traAccountCode) = "7SB000" or Ucase(traAccountCode) = "7SL001" _
				or Ucase(traAccountCode) = "7SVR01" or Ucase(traAccountCode) = "7TM000" or Ucase(traAccountCode) = "7CS001" or Ucase(traAccountCode) = "7CHUR0" or Ucase(traAccountCode) = "7JB000" then
				emailTo  	= "nathan.biggin@music.yamaha.com, cameron.tait@music.yamaha.com"		
		elseif Ucase(traAccountCode) = "7AP001" or Ucase(traAccountCode) = "7BG000" or Ucase(traAccountCode) = "7CM000" or Ucase(traAccountCode) = "7CJ000" or Ucase(traAccountCode) = "" or Ucase(traAccountCode) = "7DH000" _
				or Ucase(traAccountCode) = "7DH002" or Ucase(traAccountCode) = "7DL000" or Ucase(traAccountCode) = "7GL001" or Ucase(traAccountCode) = "7GS002" or Ucase(traAccountCode) = "7SMC55" or Ucase(traAccountCode) = "7RW001" _
				or Ucase(traAccountCode) = "7SMC01" or Ucase(traAccountCode) = "7STA00" or Ucase(traAccountCode) = "7PR001" then
				emailTo 	 = "russell.wykes@music.yamaha.com"
		elseif Ucase(traAccountCode) = "9SLA01" then
				emailTo 	 = "drew.morrow@music.yamaha.com"
		else emailTo 	 = "it-aus@music.yamaha.com"
		end if

		'emailBcc 	 = "harsono.setiono@music.yamaha.com"
		'emailCc 	 = Trim(renCreatedByEmail)
		emailSubject = "New Loan Stock Transfer"
		emailBodyText =	"G'day!" & vbCrLf _
						& " " & vbCrLf _
						& "There is a new Loan Stock Transfer." & vbCrLf _
						& " " & vbCrLf _
						& "FROM: " & traAccountCode & vbCrLf _
						& "TO: " & traRecipient & vbCrLf _
						& " " & vbCrLf _
						& "Model no: " & traModelNo & vbCrLf _
						& "Serial no: " & traSerialNo & vbCrLf _
						& "Qty: " & traQty & vbCrLf _
						& " " & vbCrLf _
						& "Submitted by: " & traCreatedBy & vbCrLf _
						& " " & vbCrLf _
						& "Go to the below link to approve:" & vbCrLf _
						& "http://intranet:96/loan-transfer.asp" & vbCrLf _
						& " " & vbCrLf _
						& "Thank you."

		Set oMail.Configuration = iConf

		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send

		Set iConf = Nothing
		Set Flds = Nothing

		response.Redirect("loan-transfer.asp")
	end if

	Call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOAN TRANSFER
'-----------------------------------------------
function addLoanTransfer(traAccountCode, traModelNo, traSerialNo, traQty, traRecipient, traCreatedBy)
	Dim cmdObj, paraObj

    call OpenDataBase

    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spAddLoanTransfer"
    cmdObj.CommandType = AdCmdStoredProc

	Set paraObj = cmdObj.CreateParameter("@traAccountCode",AdVarChar,AdParamInput,50,traAccountCode)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@traModelNo",AdVarChar,AdParamInput,50,traModelNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@traSerialNo",AdVarChar,AdParamInput,50,traSerialNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@traOrderNo",AdInteger,AdParamInput,4,session("newOrderNo"))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@traOrderLine",AdInteger,AdParamInput,4,session("newOrderLine"))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@traQty",AdInteger,AdParamInput,4,traQty)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@traRecipient",AdVarChar,AdParamInput,50,traRecipient)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@traCreatedBy",AdVarChar,AdParamInput,50,traCreatedBy)
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



		Response.Redirect("loan-transfer.asp")
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)

    call CloseDataBase
end function

'-----------------------------------------------
' GET LOAN TRANSFER
'-----------------------------------------------
Function getLoanTransfer(traID)
    dim strSQL

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_loan_transfer WHERE traID = " & traID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("traAccountCode") 	= rs("traAccountCode")
		session("traModelNo") 		= rs("traModelNo")
		session("traSerialNo") 		= rs("traSerialNo")
		session("traQty") 			= rs("traQty")
		session("traConnote") 		= rs("traConnote")
		session("traRecipientID") 	= rs("traRecipientID")

		session("traRecipientConfirmation") 	= rs("traRecipientConfirmation")
		session("traRecipientConfirmationDate") = rs("traRecipientConfirmationDate")
		session("traMarketingApproval")			= rs("traMarketingApproval")
		session("traMarketingApprovalDate")		= rs("traMarketingApprovalDate")
		session("traLogisticsConfirmation")		= rs("traLogisticsConfirmation")
		session("traLogisticsConfirmationDate")	= rs("traLogisticsConfirmationDate")

		session("traDateCreated")  	= rs("traDateCreated")
		session("traCreatedBy") 	= rs("traCreatedBy")
		session("traDateModified") 	= rs("traDateModified")
		session("traModifiedBy") 	= rs("traModifiedBy")

		session("traComments") 		= rs("traComments")
		session("traStatus") 	   	= rs("traStatus")
    end if

    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE LOAN TRANSFER
'----------------------------------------------------------------------------------------
Function updateTransfer(traID, traModelNo, traSerialNo, traQty, traRecipient, traConnote, traComments, traModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_transfer SET "
	strSQL = strSQL & "traModelNo = '" & Server.HTMLEncode(traModelNo) & "',"
	strSQL = strSQL & "traSerialNo = '" & Server.HTMLEncode(traSerialNo) & "',"
	strSQL = strSQL & "traQty = '" & Server.HTMLEncode(traQty) & "',"
	strSQL = strSQL & "traRecipient = '" & Server.HTMLEncode(traRecipient) & "',"
	strSQL = strSQL & "traConnote = '" & Server.HTMLEncode(traConnote) & "',"
	strSQL = strSQL & "traComments = '" & Server.HTMLEncode(traComments) & "',"
	strSQL = strSQL & "traDateModified = GetDate(),"
	strSQL = strSQL & "traModifiedBy = '" & Trim(traModifiedBy) & "' WHERE traID = " & traID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	'On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The transfer has been updated.</div>"
	end if

	Call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE LOAN TRANSFER CONNOTE
'----------------------------------------------------------------------------------------
Function updateTransferConnote(traID, traConnote, traModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_transfer SET "	
	strSQL = strSQL & "traConnote = '" & Server.HTMLEncode(traConnote) & "',"
	strSQL = strSQL & "traDateModified = GetDate(),"
	strSQL = strSQL & "traModifiedBy = '" & Trim(traModifiedBy) & "' WHERE traID = " & traID

	'response.Write strSQL

	on error resume next
	conn.Execute strSQL

	'On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if

	Call CloseDataBase()
end function

'-----------------------------------------------
' APPROVE TRANSFER BY MARKETING MANAGER
'-----------------------------------------------
Function approveTransfer(traID,traModifiedBy,traCreatedByEmail,traRecipientEmail)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_transfer SET "
	strSQL = strSQL & "traMarketingApproval = '1',"
	strSQL = strSQL & "traMarketingApprovalDate = GetDate(), "
	strSQL = strSQL & "traMarketingApprovalBy = '" & traModifiedBy & "',"
	strSQL = strSQL & "traModifiedBy = '" & traModifiedBy & "',"	
	strSQL = strSQL & "traDateModified = GetDate() "
	strSQL = strSQL & "	WHERE traID = " & traID

	response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields

	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update

		emailFrom 	= "automailer@gmx.yamaha.com"
		'emailTo 	 = "harsono.setiono@music.yamaha.com"
		emailTo 	 = Trim(traRecipientEmail)
		emailCc 	 = Trim(traCreatedByEmail)
		emailSubject = "Loan Stock Transfer - Approved by Marketing"
		emailBodyText =	"Hi there!" & vbCrLf _
						& " " & vbCrLf _
						& "Loan Stock Transfer (ID:" & traID & ") has been approved by Marketing Manager and requires your acknowledgement." & vbCrLf _
						& " " & vbCrLf _
						& "http://intranet:96/loan-transfer.asp" & vbCrLf _	
						& " " & vbCrLf _
						& "Thank you."

		Set oMail.Configuration = iConf

		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send

		Set iConf = Nothing
		Set Flds = Nothing

		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if

	Call CloseDataBase()
end function

'-----------------------------------------------
' REJECT TRANSFER BY MARKETING MANAGER
'-----------------------------------------------
Function rejectTransfer(traID,traModifiedBy,traCreatedByEmail,traRecipientEmail)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_transfer SET "
	strSQL = strSQL & "traMarketingApproval = '2',"
	strSQL = strSQL & "traMarketingRejectionDate = GetDate(), "
	strSQL = strSQL & "traMarketingRejectionBy = '" & traModifiedBy & "',"
	strSQL = strSQL & "traStatus = '2',"
	strSQL = strSQL & "traModifiedBy = '" & traModifiedBy & "',"
	strSQL = strSQL & "traDateModified = GetDate() "
	strSQL = strSQL & "	WHERE traID = " & traID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields

	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update

		emailFrom 	= "automailer@gmx.yamaha.com"
		'emailTo 	 = "harsono.setiono@music.yamaha.com"
		emailTo 	 = Trim(traCreatedByEmail)
		emailCc 	 = Trim(traRecipientEmail)
		emailSubject = "Loan Stock Transfer - Rejected by Marketing"
		emailBodyText =	"Hi there!" & vbCrLf _
						& " " & vbCrLf _
						& "Your Loan Stock Transfer (ID:" & traID & ") has been rejected by Marketing Manager." & vbCrLf _
						& " " & vbCrLf _
						& "http://intranet:96/loan-transfer.asp" & vbCrLf _	
						& " " & vbCrLf _
						& "Thank you."

		Set oMail.Configuration = iConf

		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send

		Set iConf = Nothing
		Set Flds = Nothing

		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if

	Call CloseDataBase()
end function

'-----------------------------------------------
' ACKNOWLEDGE BY RECIPIENT
'-----------------------------------------------
Function acknowledgeTransfer(traID,traModifiedBy,traCreatedByEmail)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_transfer SET "
	strSQL = strSQL & "traRecipientConfirmation = '1',"
	strSQL = strSQL & "traRecipientConfirmationDate = GetDate(), "
	strSQL = strSQL & "traRecipientConfirmationBy = '" & traModifiedBy & "',"
	strSQL = strSQL & "traModifiedBy = '" & traModifiedBy & "',"	
	strSQL = strSQL & "traDateModified = GetDate() "
	strSQL = strSQL & "	WHERE traID = " & traID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields

	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update

		emailFrom 	= "automailer@gmx.yamaha.com"
		emailTo 	 = "logistic_supportYMA@gmx.yamaha.com"
		emailCc 	 = Trim(traCreatedByEmail)
		emailSubject = "Loan Stock Transfer - Logistics to confirm"
		emailBodyText =	"G'day Logistics!" & vbCrLf _
						& " " & vbCrLf _
						& "There is a Loan Stock Transfer (ID:" & traID & ") for you to confirm:" & vbCrLf _
						& " " & vbCrLf _
						& "http://intranet:96/loan-transfer.asp" & vbCrLf _	
						& " " & vbCrLf _
						& "Thank you."

		Set oMail.Configuration = iConf

		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send

		Set iConf = Nothing
		Set Flds = Nothing

		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if

	Call CloseDataBase()
end function

'-----------------------------------------------
' CONFIRM BY LOGISTICS
'-----------------------------------------------
Function confirmTransfer(traID,traModifiedBy,traCreatedByEmail,traRecipientEmail)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_transfer SET "
	strSQL = strSQL & "traLogisticsConfirmation = '1',"
	strSQL = strSQL & "traLogisticsConfirmationDate = GetDate(), "
	strSQL = strSQL & "traLogisticsConfirmationBy = '" & traModifiedBy & "',"
	strSQL = strSQL & "traModifiedBy = '" & traModifiedBy & "',"
	strSQL = strSQL & "traDateModified = GetDate(), traStatus = 0 "
	strSQL = strSQL & "	WHERE traID = " & traID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields

	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update

		emailFrom 	= "automailer@gmx.yamaha.com"
		'emailTo 	 = "harsono.setiono@music.yamaha.com"
		emailTo 	 = Trim(traRecipientEmail)
		emailCc 	 = Trim(traCreatedByEmail)
		emailSubject = "Loan Stock Transfer - Complete"
		emailBodyText =	"G'day!" & vbCrLf _
						& " " & vbCrLf _
						& "Your Loan Stock Transfer (ID:" & traID & ") is now complete." & vbCrLf _
						& " " & vbCrLf _
						& "http://intranet:96/loan-transfer.asp" & vbCrLf _	
						& " " & vbCrLf _
						& "Thank you."

		Set oMail.Configuration = iConf

		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send

		Set iConf = Nothing
		Set Flds = Nothing

		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if

	Call CloseDataBase()
end function
%>