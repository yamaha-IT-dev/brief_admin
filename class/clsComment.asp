<%
'-----------------------------------------------
' LIST COMMENTS
'-----------------------------------------------
function listComments(intID,intTypeID)
    dim strSQL
	dim intRecordCount
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_comments "
	strSQL = strSQL & "	WHERE associated_id = '" & intID & "' "
	strSQL = strSQL & "		AND comment_type = '" & intTypeID & "' "
	strSQL = strSQL & "	ORDER BY comment_date"
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	

    strCommentsList = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
	
		For intRecord = 1 To rs.PageSize
		
			strCommentsList = strCommentsList & "<tr><td class=""comment_column"">"	& trim(rs("comments")) & "</td></tr>"
			strCommentsList = strCommentsList & "<tr><td class=""comment_content""><strong>" & trim(rs("comment_by")) & "</strong> - " & WeekDayName(WeekDay(rs("comment_date"))) & ", " & FormatDateTime(rs("comment_date"),1) & " at " & FormatDateTime(rs("comment_date"),3) & "</td></tr>"
			
			rs.movenext
			
			If rs.EOF Then Exit For
		next
	else
        strCommentsList = "<tr><td>&nbsp;</td></tr>"
	end if
	
	strCommentsList = strCommentsList & "<tr>"
	
    call CloseDataBase()
end function

'-----------------------------------------------
' ADD COMMENT
'-----------------------------------------------
function addComment(intID,intTypeID)
	dim strSQL
	
	dim strComment
	strComment 		= Replace(Request.Form("txtComment"),"'","''")
	
	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_comments ("
	strSQL = strSQL & " comment_type, "
	strSQL = strSQL & " comments, "
	strSQL = strSQL & " associated_id, "
	strSQL = strSQL & " comment_by"
	strSQL = strSQL & ") VALUES ("
	strSQL = strSQL & " '" & intTypeID & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strComment) & "',"	
	strSQL = strSQL & " '" & intID & "',"
	strSQL = strSQL & " '" & session("logged_username") & "')"

	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		dim oMail
		dim iConf
		dim Flds
		
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
			
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.18"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		iConf.Fields.Update
							
		'emailTo 		= "harsono_setiono@gmx.yamaha.com"
		'if Session("project_department") = "AV" or Session("project_department") = "YMEC" then
		'	emailTo		= "mark_underwood@gmx.yamaha.com"
		'else
			emailTo		= "sonja.loader-jurac@music.yamaha.com"
		'end if
		
		'emailCc 		= "harsono_setiono@gmx.yamaha.com"		
		emailFrom 		= "automailer@gmx.yamaha.com"
		emailSubject 	= "New comment has been posted - " & session("project_title")
		emailBodyText  	= "G'day," & vbCrLf _
						& "" & vbCrLf _
						& "A comment has been posted to this brief: " & session("project_title") & " " & vbCrLf _
						& "------------------------------------------------------------" & vbCrLf _
						& "" & strComment & vbCrLf _
						& "by: " & session("logged_username") & " " & vbCrLf _
						& "------------------------------------------------------------" & vbCrLf _
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
			
		Set iConf 		= Nothing
		Set Flds 		= Nothing
				
		strMessageText = "Comment has been added."
	end if 
	
	Call CloseDataBase()
end function
%>