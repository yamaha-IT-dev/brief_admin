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
    dim strTitle
    dim strOutputDetails
    dim strImagesLocation
    dim strCopyLocation
    dim strAspect1
    dim strAspect2
    dim strAspect3
    dim strLookFeel
    dim strDescription

    strTitle            = Replace(Session("new_project_title"),"'","''")
    strOutputDetails    = Replace(Session("new_project_output_details"),"'","''")
    strImagesLocation   = Replace(Session("new_project_image_location"),"'","''")
    strCopyLocation     = Replace(Session("new_project_copy_location"),"'","''")
    strAspect1          = Replace(Session("new_project_aspect_1"),"'","''")
    strAspect2          = Replace(Session("new_project_aspect_2"),"'","''")
    strAspect3          = Replace(Session("new_project_aspect_3"),"'","''")
    strLookFeel         = Replace(Session("new_project_look_feel"),"'","''")
    strDescription      = Replace(Session("new_project_description"),"'","''")

    call OpenDataBase()

    strSQL = "INSERT INTO yma_project ("
    strSQL = strSQL & "project_department, "
    strSQL = strSQL & "project_title, "
    strSQL = strSQL & "project_priority, "
    strSQL = strSQL & "project_output_printed, "
    strSQL = strSQL & "project_output_web, "
    strSQL = strSQL & "project_output_details, "
    strSQL = strSQL & "project_first_deadline, "
    strSQL = strSQL & "project_second_deadline, "
    strSQL = strSQL & "project_deadline, "
    strSQL = strSQL & "project_image_location, "
    strSQL = strSQL & "project_copy_location, "
    strSQL = strSQL & "project_aspect_1, "
    strSQL = strSQL & "project_aspect_2, "
    strSQL = strSQL & "project_aspect_3, "
    strSQL = strSQL & "project_look_feel, "
    strSQL = strSQL & "project_description, "
    strSQL = strSQL & "project_created_by, "
    strSQL = strSQL & "project_status"
    strSQL = strSQL & ") VALUES ("
    strSQL = strSQL & "'" & Trim(Session("new_project_department")) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strTitle) & "',"
    strSQL = strSQL & "'" & Session("new_project_priority") & "',"
    strSQL = strSQL & "'" & Session("new_project_output_printed") & "',"
    strSQL = strSQL & "'" & Session("new_project_output_web") & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strOutputDetails) & "',"
    strSQL = strSQL & " CONVERT(datetime,'" & Session("new_project_first_deadline") & "',103),"
    strSQL = strSQL & " CONVERT(datetime,'" & Session("new_project_second_deadline") & "',103),"
    strSQL = strSQL & " CONVERT(datetime,'" & Session("new_project_deadline") & "',103),"
    strSQL = strSQL & "'" & Server.HTMLEncode(strImagesLocation) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strCopyLocation) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strAspect1) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strAspect2) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strAspect3) & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strLookFeel) & "',"
    strSQL = strSQL & "'" & strDescription & "',"
    strSQL = strSQL & "'" & session("logged_username") & "',2)"

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

        if Session("new_project_department") = "AV" then
            emailTo       = "russell.wykes@music.yamaha.com"
        end if

        emailCc         = session("emp_email")
        'emailBcc        = "victor.samson@music.yamaha.com"
        emailFrom       = "noreply@music.yamaha.com"
        emailSubject    = "New Graphic Design Brief - " & Session("new_project_department")
        emailBodyText   = "Title:        " & Session("new_project_title") & vbCrLf _
                        & "Department:   " & Session("new_project_department") & vbCrLf _
                        & "Submitted by: " & session("logged_username") & vbCrLf _
                        & "Priority:     " & Session("new_project_priority") & vbCrLf _
                        & "----------------------------------------------------------------------------" & vbCrLf _
                        & "OUTPUT" & vbCrLf _
                        & "----------------------------------------------------------------------------" & vbCrLf _
                        & "Printed:      " & Session("new_project_output_printed") & vbCrLf _
                        & "Online:       " & Session("new_project_output_web") & vbCrLf _
                        & "Details:      " & Session("new_project_output_details") & vbCrLf _
                        & "----------------------------------------------------------------------------" & vbCrLf _
                        & "DEADLINES" & vbCrLf _
                        & "----------------------------------------------------------------------------" & vbCrLf _
                        & "First draft:         " & Session("new_project_first_deadline") & vbCrLf _
                        & "Second draft:        " & Session("new_project_second_deadline") & vbCrLf _
                        & "Printing/Publishing: " & Session("new_project_deadline") & vbCrLf _
                        & "----------------------------------------------------------------------------" & vbCrLf _
                        & "Images?              " & Session("new_project_image_location") & vbCrLf _
                        & "Copy?                " & Session("new_project_copy_location") & vbCrLf _
                        & "Important aspect 1:  " & Session("new_project_aspect_1") & vbCrLf _
                        & "Important aspect 2:  " & Session("new_project_aspect_2") & vbCrLf _
                        & "Important aspect 3:  " & Session("new_project_aspect_3") & vbCrLf _
                        & "Look and feel:       " & Session("new_project_look_feel") & vbCrLf _
                        & "Description:         " & Session("new_project_description") & vbCrLf _
                        & "----------------------------------------------------------------------------" & vbCrLf _
                        & " " & vbCrLf _
                        & "This is an automated email - please do not reply to this email."

        Set oMail.Configuration = iConf
        oMail.To        = emailTo
        oMail.Cc        = emailCc
        oMail.Bcc       = emailBcc
        oMail.From      = emailFrom
        oMail.Subject   = emailSubject
        oMail.TextBody  = emailBodyText
        oMail.Send

        Set iConf = Nothing
        Set Flds = Nothing

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
        <td valign="top" class="maincontent">
            <img src="images/backward_arrow.gif" border="0" /> <a href="default.asp">Back to Home</a>
        
		<div>
				<%if strMessageText <> "" then %>
					<div class="alert alert-danger">
						<strong><% Response.Write(strMessageText ) %></strong> 							
					</div>
				<% end if %>			
		</div>

		<h2>Submit Brief</h2>
            <form action="" method="post" name="form_confirm_project" id="form_confirm_project">
                <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                    <tr>
                        <td colspan="2" class="item_maintenance_header">
                            Brief Details
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Department: <%= Session("new_project_department") %></strong>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Title:</strong> <%= Session("new_project_title") %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Priority:</strong>
                            <%
                                select case Session("new_project_priority")
                                    case 1
                                        response.Write("<font class=""low_font"">Low</font>")
                                    case 2
                                        response.Write("<font class=""medium_font"">Medium</font>")
                                    case 3
                                        response.Write("<font class=""high_font"">High</font>")
                                end select
                            %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Output:</strong>
                            <% if Session("new_project_output_printed") = "1" then Response.Write " Printed" end if%>
                            <% if Session("new_project_output_web") = "1" then Response.Write " Online" end if%>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Details:</strong>
                            <br />
                            <%= Session("new_project_output_details") %>
                        </td>
                    </tr>
                    <tr>
                        <td width="50%">
                            <strong>1<sup>st</sup> draft deadline:</strong>
                            <br />
                            <%= Session("new_project_first_deadline") %>
                        </td>
                        <td width="50%" valign="top">
                            <strong>2<sup>nd</sup> draft deadline:</strong>
                            <br />
                            <%= Session("new_project_second_deadline") %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Printing / publishing deadline:</strong>
                            <br />
                            <%= Session("new_project_deadline") %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Image(s) supplied location:</strong>
                            <br />
                            <%= Session("new_project_image_location") %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Copy supplied location:</strong>
                            <br />
                            <%= Session("new_project_copy_location") %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Most important aspects:</strong>
                            <ol>
                                <li><%= Session("new_project_aspect_1") %></li>
                                <li><%= Session("new_project_aspect_2") %></li>
                                <li><%= Session("new_project_aspect_3") %></li>
                            </ol>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Up to 3 words that describe the look and feel:</strong>
                            <br />
                            <%= Session("new_project_look_feel") %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>Publishing / Printing requirement:</strong>
                            <br />
                            <%= Session("new_project_description") %>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <input type="hidden" name="Action" value="Confirm" />
                            <input type="submit" value="Submit Brief" <% if Session("new_project_status") = "0" then Response.Write "disabled" end if%> />
                        </td>
                    </tr>
                </table>
            </form>
			<div>
				<%if strMessageText <> "" then %>
					<div class="alert alert-danger">
						<strong><% Response.Write(strMessageText ) %></strong> 							
					</div>
				<% end if %>			
			</div>
            <p align="right"><input type="button" value="Back" onclick="goBack()" /></p>
        </td>
    </tr>
</table>
</body>
</html>
