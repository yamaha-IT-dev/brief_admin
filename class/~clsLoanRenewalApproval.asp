<%
'-----------------------------------------------
' GENERAL MANAGER APPROVAL
'-----------------------------------------------
Function approveRenewal(intRenewalID, strRequesterEmail, strModifiedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE tbl_loan_renewal SET "
	strSQL = strSQL & "renApproval = '1',"
	strSQL = strSQL & "renApprovalDate = GetDate(),"
	strSQL = strSQL & "renDateModified = GetDate(),"
	strSQL = strSQL & "renModifiedBy = '" & strModifiedBy & "'" 
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
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.18"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
						
		emailFrom 	= "automailer@gmx.yamaha.com"		
		emailTo 	= Trim(strRequesterEmail)
		'emailCc 	= "Carolyn_Simonds@gmx.yamaha.com"
		'emailBcc 	= "Harsono_Setiono@gmx.yamaha.com"
		emailSubject = "Loan Stock Renewal - Approved"
				
		emailBodyText =	"G'day," & vbCrLf _
						& " " & vbCrLf _
						& "Your Loan Stock Renewal Request has been approved by GM." & vbCrLf _
						& " " & vbCrLf _
						& "Account: " & session("loan_account") & vbCrLf _
						& "Product: " & session("loan_product") & vbCrLf _
						& "Serial : " & session("loan_serial_no") & vbCrLf _
						& " " & vbCrLf _
						& "Please click on the below link to view it:" & vbCrLf _
						& "http://intranet:96/view_loan.asp?order=" & Trim(Request("order")) & "&line=" & Trim(Request("line")) & "" & vbCrLf _
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
		
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The request has been approved by GM.</div>"
	end if
	
	Call CloseDataBase()
end Function

'-----------------------------------------------
' GENERAL MANAGER REJECTION
'-----------------------------------------------
Function rejectRenewal(intRenewalID, strRequesterEmail, strModifiedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE tbl_loan_renewal SET "
	strSQL = strSQL & "renApproval = '2',"
	strSQL = strSQL & "renApprovalDate = GetDate(),"
	strSQL = strSQL & "renDateModified = GetDate(),"
	strSQL = strSQL & "renModifiedBy = '" & strModifiedBy & "'" 
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
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.18"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
						
		emailFrom 	= "automailer@gmx.yamaha.com"
		emailTo 	= Trim(strRequesterEmail)
		'emailCc 	= "Carolyn_Simonds@gmx.yamaha.com"	
		'emailBcc 	= "Harsono_Setiono@gmx.yamaha.com"
		emailSubject = "Loan Stock Renewal - Rejected"
						
		emailBodyText =	"G'day," & vbCrLf _
						& " " & vbCrLf _
						& "Your Product Renewal request for this following Dealer has been rejected by GM." & vbCrLf _
						& " " & vbCrLf _
						& "Account: " & session("loan_account") & vbCrLf _
						& "Product: " & session("loan_product") & vbCrLf _
						& "Serial : " & session("loan_serial_no") & vbCrLf _
						& " " & vbCrLf _
						& "Please click on the below link to update it:" & vbCrLf _
						& "http://intranet:96/view_loan.asp?order=" & Trim(Request("order")) & "&line=" & Trim(Request("line")) & "" & vbCrLf _
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
		
		strMessageText = "<div class=""rejection_text""><img src=""images/icon_cross.jpg""> The request has been rejected by GM.</div>"
	end if
	
	Call CloseDataBase()
end Function
%>