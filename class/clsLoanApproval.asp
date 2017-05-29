<%
'----------------------------------------------------------------------------------------
' INCREMENT RENEWAL COUNTER
'----------------------------------------------------------------------------------------
Function incrementRenewalCounter(stockOrderNo, stockOrderLine)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_location SET "
	strSQL = strSQL & "stockRenewalCounter = stockRenewalCounter + 1 WHERE stockOrderNo = " & stockOrderNo & " AND stockOrderLine = " & stockOrderLine

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "Counter has been updated"
	end if

	Call CloseDataBase()
end function

'-----------------------------------------------
' RENEWAL APPROVAL
'-----------------------------------------------
Function approveLoanRenewal(intRenewalID, strItemCode, strRequesterEmail, strRenComments, strApprovedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE tbl_loan_renewal SET "
	strSQL = strSQL & "renStatus = '0',"
	strSQL = strSQL & "renApprovalDate = GetDate(),"
	strSQL = strSQL & "renModifiedBy = '" & strApprovedBy & "'" 
	strSQL = strSQL & " WHERE renID = " & intRenewalID
	
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
		
		emailFrom 	= "automailer@music.yamaha.com"
		emailTo 	= Trim(strRequesterEmail)
		'emailCc 	= "Harsono_Setiono@gmx.yamaha.com"
		emailSubject = "Loan Stock Renewal - APPROVED"
		
		emailBodyText =	"G'day," & vbCrLf _
						& " " & vbCrLf _
						& "Your Loan Stock Renewal Request (" & strItemCode & ") has been APPROVED by " & strApprovedBy & "." & vbCrLf _
						& "Comments: " & strRenComments & vbCrLf _
						& " " & vbCrLf _
						& "Thank you. (This is an automated email - please do not reply to this email)"
						
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
		
		strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Renewal has been approved by GM.</div>"
	end if
	
	Call CloseDataBase()
end Function

'-----------------------------------------------
' RENEWAL REJECTION
'-----------------------------------------------
Function rejectLoanRenewal(intRenewalID, strItemCode, strRequesterEmail, strRenComments, strApprovedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE tbl_loan_renewal SET "
	strSQL = strSQL & "renStatus = '2',"
	strSQL = strSQL & "renApprovalDate = GetDate(),"
	strSQL = strSQL & "renModifiedBy = '" & strApprovedBy & "'" 
	strSQL = strSQL & " WHERE renID = " & intRenewalID
	
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
						
		emailFrom 	= "automailer@music.yamaha.com"
		emailTo 	= Trim(strRequesterEmail)
		'emailCc 	= "Harsono_Setiono@gmx.yamaha.com"
		emailSubject = "Loan Stock Renewal - REJECTED"
		
		emailBodyText =	"G'day," & vbCrLf _
						& " " & vbCrLf _
						& "Your Loan Stock Renewal Request (" & strItemCode & ") has been REJECTED by " & strApprovedBy & "." & vbCrLf _
						& "Comments: " & strRenComments & vbCrLf _
						& " " & vbCrLf _						
						& "Thank you. (This is an automated email - please do not reply to this email)"
						
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
		
		strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_cross.jpg""> The Renewal has been rejected by GM.</div>"
	end if
	
	Call CloseDataBase()
end Function


Function updateComments(intRenewalID, strItemCode, strRenComments, renModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_renewal SET "	
	strSQL = strSQL & "renComments = '" & Server.HTMLEncode(strRenComments) & "', "	
	strSQL = strSQL & "renDateModified = GetDate(), "
	strSQL = strSQL & "renModifiedBy = '" & Trim(renModifiedBy) & "' WHERE renID = " & intRenewalID

	response.Write strSQL
	
	on error resume next
	conn.Execute strSQL

	'On error Goto 0

	if err <> 0 then
		strMessageText = err.description
		response.write  err.description
	else
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))		
	end if

	Call CloseDataBase()
end function
%>