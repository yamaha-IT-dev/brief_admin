<%
'-----------------------------------------------
' GET PROJECT
'-----------------------------------------------
Function getBrief(intID)
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_project WHERE project_id = " & intID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("project_id") 				= rs("project_id")
		session("project_department") 		= rs("project_department")
		session("project_title") 			= rs("project_title")
		session("project_priority") 		= rs("project_priority")
		session("project_gl_code") 			= rs("project_gl_code")
		session("project_contact") 			= rs("project_contact")
		session("project_output_printed") 	= rs("project_output_printed")
		session("project_output_web") 		= rs("project_output_web")
		session("project_output_details") 	= rs("project_output_details")
		session("project_first_deadline") 	= rs("project_first_deadline")
		session("project_second_deadline") 	= rs("project_second_deadline")
		session("project_deadline") 		= rs("project_deadline")
		session("project_image_location") 	= rs("project_image_location")
		session("project_copy_location") 	= rs("project_copy_location")
		session("project_aspect_1") 		= rs("project_aspect_1")
		session("project_aspect_2") 		= rs("project_aspect_2")
		session("project_aspect_3") 		= rs("project_aspect_3")
		session("project_look_feel") 		= rs("project_look_feel")
		session("project_description") 		= rs("project_description")
		session("project_quote") 			= rs("project_quote")
		session("project_progress") 		= rs("project_progress")
		session("project_comments") 		= rs("project_comments")
		session("project_actual_hours") 	= rs("project_actual_hours")
		session("project_date_created") 	= rs("project_date_created")
		session("project_created_by") 		= rs("project_created_by")
		session("project_date_modified") 	= rs("project_date_modified")
		session("project_modified_by") 		= rs("project_modified_by")
		session("project_status") 			= rs("project_status")
		session("project_job_no") 			= rs("project_job_no")
		
		session("product_manager_notified") 			= rs("product_manager_notified")
		session("product_manager_approval") 			= rs("product_manager_approval")
		session("product_manager_approval_date") 		= rs("product_manager_approval_date")
		session("product_manager_approval_by") 			= rs("product_manager_approval_by")
		session("product_manager_approval_message") 	= rs("product_manager_approval_message")
		session("marketing_manager_approval") 			= rs("marketing_manager_approval")
		session("marketing_manager_approval_date") 		= rs("marketing_manager_approval_date")
		session("marketing_manager_approval_by") 		= rs("marketing_manager_approval_by")
		session("marketing_manager_approval_message") 	= rs("marketing_manager_approval_message")		
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' PRODUCT MGR APPROVE BRIEF
'-----------------------------------------------
function approveProductManager(intBriefID,strModifiedBy,strDepartment)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE yma_project SET "
	strSQL = strSQL & "product_manager_approval = '1',"
	strSQL = strSQL & "product_manager_approval_date = getdate(),"
	strSQL = strSQL & "product_manager_approval_by = '" & strModifiedBy & "' WHERE project_id = " & intBriefID
	
	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		call addLog(intBriefID,projectModuleID,"Approved")
		
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
						
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.18"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
						
		emailFrom 	= "automailer@gmx.yamaha.com"
		'emailTo 	= trim(strRequesterEmail)
		
		select case strDepartment
			case "MPD"
				emailTo 	= "jaclyn.williams@music.yamaha.com"
				emailSubject = "GD Brief - Your (Jaclyn) approval is needed"
			case "MPD - PRO"				
				emailTo 	= "nathan.biggin@music.yamaha.com"
				emailSubject = "GD Brief - PRO Marketing Manager approval is needed"
			case "MPD - TRAD"
				emailTo 	= "cameron.tait@music.yamaha.com"
				emailSubject = "GD Brief - TRAD Marketing Manager approval is needed"
			case "CA"
				emailTo 	= "nathan.biggin@music.yamaha.com"
				emailSubject = "GD Brief - PRO Marketing Manager approval is needed"	
			case "YMEC"
				emailTo 	= "carolyn.simonds@music.yamaha.com"
				emailSubject = "GD Brief - YMEC Manager approval is needed"	
			case else
				emailTo 	= "nathan.biggin@music.yamaha.com"
				emailSubject = "GD Brief - Marketing Manager approval is needed"	
		end select						
				
		emailBodyText =	"G'day," & vbCrLf _
						& " " & vbCrLf _
						& "Marketing Manager approval is needed for the brief titled: " & session("project_title") & "." & vbCrLf _
						& " " & vbCrLf _
						& "Please click on the below link to approve it:" & vbCrLf _
						& "http://intranet:96/update_brief.asp?id=" & intBriefID & "" & vbCrLf _																	
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
		
		strMessageText = "The brief has been approved by requester."
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' MARKETING MGR APPROVE BRIEF
'-----------------------------------------------
function approveMarketingManager(intBriefID,strRequesterEmail,strModifiedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE yma_project SET "
	'strSQL = strSQL & "project_status = '0',"
	strSQL = strSQL & "marketing_manager_approval = '1',"
	strSQL = strSQL & "marketing_manager_approval_date = getdate(),"
	strSQL = strSQL & "marketing_manager_approval_by = '" & strModifiedBy & "' WHERE project_id = " & intBriefID
	
	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		call addLog(intBriefID,projectModuleID,"Marketing approved")
		
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
						
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.18"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
						
		emailFrom 	= "automailer@gmx.yamaha.com"
		emailTo 	= trim(strRequesterEmail)
		'emailTo 	= "harsono_setiono@gmx.yamaha.com"
		emailSubject = "GD Brief - Your brief has been approved"
				
		emailBodyText =	"Hello," & vbCrLf _
						& " " & vbCrLf _
						& "Marketing Manager has approved your brief titled: " & session("project_title") & "." & vbCrLf _
						& " " & vbCrLf _
						& "Please click on the below link to view it:" & vbCrLf _
						& "http://intranet:96/update_brief.asp?id=" & intBriefID & "" & vbCrLf _																	
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
		
		strMessageText = "The brief has been approved by Marketing Manager."
	end if
	
	Call CloseDataBase()
end function
%>