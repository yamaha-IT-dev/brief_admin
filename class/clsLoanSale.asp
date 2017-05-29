<%
function newLoanSale(strAccountCode, strModelNo, strSerialNo, strOrderNo, strOrderLine)	
	session("newAccountCode") 	= strAccountCode
	session("newModelNo") 		= strModelNo
	session("newSerialNo") 		= strSerialNo
	session("newOrderNo") 		= strOrderNo
	session("newOrderLine")		= strOrderLine
	
	response.Redirect("add-loan-sale.asp")
end function

function addSale(saleAccountCode, saleModelNo, saleSerialNo, saleOrderNo, saleOrderLine, saleQty, saleDealerCode, salePurchaseOrderNo, saleCreatedBy)
	dim strSQL
	dim strRecipient
	
	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_loan_sale (saleAccountCode, saleModelNo, saleSerialNo, saleOrderNo, saleOrderLine, saleQty, saleDealerCode, salePurchaseOrderNo, saleCreatedBy) VALUES ("
	strSQL = strSQL & "'" & Server.HTMLEncode(Ucase(saleAccountCode)) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Ucase(saleModelNo)) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(saleSerialNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(saleOrderNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(saleOrderLine) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(saleQty) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(saleDealerCode) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(salePurchaseOrderNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(saleCreatedBy) & "')"
	
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
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.18"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
		
		emailFrom 	 = "automailer@gmx.yamaha.com"
		emailTo 	 = "logistic_supportYMA@gmx.yamaha.com"
		emailCc 	 = "harsono.setiono@music.yamaha.com"
		'emailCc 	 = Trim(renCreatedByEmail)
		emailSubject = "New Loan Stock Sale"		
		emailBodyText =	"G'day Logistics Team," & vbCrLf _
						& " " & vbCrLf _
						& "There is a new Loan Stock Sale to a Dealer." & vbCrLf _
						& " " & vbCrLf _
						& "Loan Account: " & saleAccountCode & vbCrLf _
						& "Dealer Code: " & saleDealerCode & vbCrLf _
						& " " & vbCrLf _						
						& "Model no: " & saleModelNo & vbCrLf _
						& "Serial no: " & saleSerialNo & vbCrLf _
						& "Qty: " & saleQty & vbCrLf _
						& "Order no: " & salePurchaseOrderNo & vbCrLf _																			
						& " " & vbCrLf _
						& "Submitted by: " & saleCreatedBy & vbCrLf _
						& " " & vbCrLf _
						& "Go to the below link to approve:" & vbCrLf _
						& "http://intranet:96/loan-sale.asp" & vbCrLf _	
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
		
		response.Redirect("loan-sale.asp")	
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' GET LOAN TRANSFER
'-----------------------------------------------
Function getLoanTransfer(saleID)
    dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_loan_sale WHERE saleID = " & saleID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("saleAccountCode") 	= rs("saleAccountCode")
		session("saleModelNo") 		= rs("saleModelNo")
		session("saleSerialNo") 		= rs("saleSerialNo")
		session("saleQty") 			= rs("saleQty")
		session("saleConnote") 		= rs("saleConnote")
		session("saleDealerCodeID") 	= rs("saleDealerCodeID")
		
		session("saleDealerCodeConfirmation") 	= rs("saleDealerCodeConfirmation")
		session("saleDealerCodeConfirmationDate") = rs("saleDealerCodeConfirmationDate")
		session("saleMarketingApproval")			= rs("saleMarketingApproval")
		session("saleMarketingApprovalDate")		= rs("saleMarketingApprovalDate")
		session("saleLogisticsConfirmation")		= rs("saleLogisticsConfirmation")
		session("saleLogisticsConfirmationDate")	= rs("saleLogisticsConfirmationDate")
		
		session("saleDateCreated")  	= rs("saleDateCreated")
		session("saleCreatedBy") 	= rs("saleCreatedBy")
		session("saleDateModified") 	= rs("saleDateModified")
		session("saleModifiedBy") 	= rs("saleModifiedBy")	
		
		session("saleComments") 		= rs("saleComments")
		session("saleStatus") 	   	= rs("saleStatus")
    end if

    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE LOAN TRANSFER
'----------------------------------------------------------------------------------------
Function updateTransfer(saleID, saleModelNo, saleSerialNo, saleQty, saleDealerCode, saleConnote, saleComments, saleModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_sale SET "	
	strSQL = strSQL & "saleModelNo = '" & Server.HTMLEncode(saleModelNo) & "',"
	strSQL = strSQL & "saleSerialNo = '" & Server.HTMLEncode(saleSerialNo) & "',"
	strSQL = strSQL & "saleQty = '" & Server.HTMLEncode(saleQty) & "',"
	strSQL = strSQL & "saleDealerCode = '" & Server.HTMLEncode(saleDealerCode) & "',"
	strSQL = strSQL & "saleConnote = '" & Server.HTMLEncode(saleConnote) & "',"
	strSQL = strSQL & "saleComments = '" & Server.HTMLEncode(saleComments) & "',"
	strSQL = strSQL & "saleDateModified = GetDate(),"
	strSQL = strSQL & "saleModifiedBy = '" & Trim(saleModifiedBy) & "' WHERE saleID = " & saleID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	'On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The salensfer has been updated.</div>"
	end if

	Call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE LOAN TRANSFER CONNOTE
'----------------------------------------------------------------------------------------
Function updateSaleOrderNo(saleID, saleOrderNo, saleModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_sale SET "	
	strSQL = strSQL & "salePurchaseOrderNo = '" & Server.HTMLEncode(saleOrderNo) & "',"	
	strSQL = strSQL & "saleDateModified = GetDate(),"
	strSQL = strSQL & "saleModifiedBy = '" & Trim(saleModifiedBy) & "' WHERE saleID = " & saleID

	response.Write strSQL
	
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
' CONFIRM BY LOGISTICS
'-----------------------------------------------
Function confirmSale(saleID,saleModifiedBy,saleCreatedByEmail)
	dim strSQL
	
	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_sale SET "
	strSQL = strSQL & "saleLogisticsConfirmation = '1',"
	strSQL = strSQL & "saleLogisticsConfirmationDate = GetDate(), "
	strSQL = strSQL & "saleLogisticsConfirmationBy = '" & saleModifiedBy & "',"	
	strSQL = strSQL & "saleModifiedBy = '" & saleModifiedBy & "',"	
	strSQL = strSQL & "saleDateModified = GetDate(), saleStatus = 0 "
	strSQL = strSQL & "	WHERE saleID = " & saleID
	
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
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.18"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
		
		emailFrom 	= "automailer@gmx.yamaha.com"		
		emailTo 	 = Trim(saleCreatedByEmail)
		'emailCc 	 = "harsono.setiono@music.yamaha.com"
		emailSubject = "Loan Stock Sale - Complete"
		emailBodyText =	"G'day!" & vbCrLf _
						& " " & vbCrLf _
						& "Your Loan Stock Sale (ID:" & saleID & ") is now complete." & vbCrLf _
						& " " & vbCrLf _
						& "http://intranet:96/loan-sale.asp" & vbCrLf _	
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