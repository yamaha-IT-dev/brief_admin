<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsBrief.asp" -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!--#include file="class/clsLog.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Submit Brief</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<%
'-----------------------------------------------
' CONFIRM BRIEF
'-----------------------------------------------
sub confirmBrief
	call OpenDataBase()
	
	strSQL = "UPDATE yma_project SET "	
	strSQL = strSQL & "project_status = '2',"
	strSQL = strSQL & "project_date_modified = getdate(),"
	strSQL = strSQL & "project_modified_by = '" & session("logged_username") & "' WHERE project_id = " & session("project_id")
		
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
		
		objCDOSYSMail.Configuration = objCDOSYSCnfg
		
		'if Session("project_department") = "AV" or Session("project_department") = "YMEC" then
		'	emailTo		= "mark_underwood@gmx.yamaha.com"
		'else
			emailTo		= "sonja.loader-jurac@music.yamaha.com"
		'end if
		emailCc 		= session("emp_email") 		
		'emailBcc 		= "alexander_payne@gmx.yamaha.com"
		'emailBcc 		= "harsono_setiono@gmx.yamaha.com"
		emailFrom 		= "automailer@gmx.yamaha.com"		
		emailSubject 	= "New Graphic Design Brief - " & Session("project_department")
		
		emailBodyText   = "Title:        " & Session("project_title") & vbCrLf _
						& "Department:   " & Session("project_department") & vbCrLf _						
						& "Submitted by: " & session("logged_username") & vbCrLf _
						& "Priority:     " & session("project_priority") & vbCrLf _
						& "----------------------------------------------------------------------------" & vbCrLf _
						& "OUTPUT" & vbCrLf _
						& "----------------------------------------------------------------------------" & vbCrLf _
						& "Printed:      " & Session("project_output_printed") & vbCrLf _
						& "Online:       " & Session("project_output_web") & vbCrLf _
						& "Details:      " & Session("project_output_details") & vbCrLf _
						& "----------------------------------------------------------------------------" & vbCrLf _
						& "DEADLINES" & vbCrLf _
						& "----------------------------------------------------------------------------" & vbCrLf _
						& "First draft:         " & Session("project_first_deadline") & vbCrLf _
						& "Second draft:        " & Session("project_second_deadline") & vbCrLf _			
						& "Printing/Publishing: " & Session("project_deadline") & vbCrLf _
						& "----------------------------------------------------------------------------" & vbCrLf _
						& "Images?              " & Session("project_image_location") & vbCrLf _
						& "Copy?                " & Session("project_copy_location") & vbCrLf _
						& "Important aspect 1:  " & Session("project_aspect_1") & vbCrLf _
						& "Important aspect 2:  " & Session("project_aspect_2") & vbCrLf _
						& "Important aspect 3:  " & Session("project_aspect_3") & vbCrLf _
						& "Look and feel:       " & Session("project_look_feel") & vbCrLf _
						& "Description:         " & Session("project_description") & vbCrLf _
						& "GL Code:             " & Session("project_gl_code") & vbCrLf _
						& "----------------------------------------------------------------------------" & vbCrLf _
						& " " & vbCrLf _
						& "This is an automated email - please do not reply to this email."
		
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
		
		call addLog(session("project_id"),projectModuleID,"Submitted draft")

		Response.Redirect("thank-you.asp")
	end if

	call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		select case Trim(Request.Form("Action"))
			case "Confirm"
				call confirmBrief	
		end select
	end if
end sub

call main

dim strMessageText
dim strCommentsList
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td valign="top" class="maincontent"><img src="images/backward_arrow.gif" border="0" /> <a href="default.asp">Back to Home</a>
      <h2>Submit Brief</h2>
      <form action="" method="post" name="form_confirm_project" id="form_confirm_project">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="item_maintenance_header">Brief Details</td>
          </tr>
          <tr>
            <td colspan="2"><strong>Title:</strong> <%= session("project_title") %></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Priority:</strong> 
			<%	select case session("project_priority") 
					case 1
						response.Write("<font class=""low_font"">Low</font>")
					case 2
						response.Write("<font class=""medium_font"">Medium</font>")
					case 3
						response.Write("<font class=""high_font"">High</font>")
				end select %></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Output:</strong>
              <% if session("project_output_printed") = "1" then Response.Write " Printed" end if%>
              <% if session("project_output_web") = "1" then Response.Write " Online" end if%></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Details:</strong><br />
              <%= session("project_output_details") %></td>
          </tr>
          <tr>
            <td width="50%"><strong>1<sup>st</sup> draft deadline:</strong><br />
              <%= session("project_first_deadline") %></td>
            <td width="50%" valign="top"><strong>2<sup>nd</sup> draft deadline:</strong><br />
              <%= session("project_second_deadline") %></td>
          </tr>
          <tr>
            <td><strong>Printing / publishing deadline:</strong><br />
              <%= session("project_deadline") %></td>
            <td><strong>GL code:</strong><br />
              <%= session("project_gl_code") %></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Image(s) supplied location:</strong><br />
              <%= session("project_image_location") %></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Copy supplied location:</strong><br />
              <%= session("project_copy_location") %></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Most important aspects:</strong>
            <ol>
            	<li><%= Session("project_aspect_1") %></li>
                <li><%= Session("project_aspect_2") %></li>
                <li><%= Session("project_aspect_3") %></li>
            </ol>
            </td>
          </tr>
          <tr>
            <td colspan="2"><strong>Up to 3 words that describe the look and feel:</strong><br />
            <%= Session("project_look_feel") %></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Publishing / Printing requirement:</strong><br />
              <%= session("project_description") %></td>
          </tr>
          <tr>
            <td colspan="2"><input type="hidden" name="Action" value="Confirm" />
          <input type="submit" value="Submit Brief" <% if session("project_status") = "0" then Response.Write "disabled" end if%> /></td>
          </tr>
        </table>
      </form>
      <p align="right"><input type="button" value="Back" onclick="goBack()" /></p>
    </td>
  </tr>
</table>
</body>
</html>