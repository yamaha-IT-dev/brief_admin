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
<title>Update Brief</title>

<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link rel="stylesheet" href="include/pikaday.css" type="text/css" />
<script src="../include/generic_form_validations.js"></script>
<script src="include/usableforms.js"></script>
<script>
function validateFormOnSubmit(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtTitle,"Title");
    reason += validateEmptyField(theForm.txtOutputDetails,"Output details");
    reason += validateDate(theForm.txtFirstDeadline);
    reason += validateDeadline(theForm.txtDeadline);
    reason += validateEmptyField(theForm.txtImagesLocation,"Images Location");
    reason += validateEmptyField(theForm.txtCopyLocation,"Copy Location");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);
        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Update';
        return true;
    }
}

function copyBrief(theForm) {
    if (confirm ("Please click OK to copy this brief.")) {
        theForm.Action.value = 'Copy';
        return true;
    } else {
        return false;
    }
}

function validateDetailsFormOnSubmit(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateNumeric(theForm.txtQuote,"Quote");
    reason += validateSpecialCharacters(theForm.txtQuote,"Quote");
    reason += validateSpecialCharacters(theForm.txtActualHours,"Actual hours");

    if (theForm.cboProgress.value == "100" && theForm.cboStatus.value != "0") {
        alert("Please set to status to Completed.");
        blnSubmit = false;
        return false;
    }

    if (theForm.cboProgress.value != "100" && theForm.cboStatus.value == "0") {
        alert("Please set to progress to 100%.");
        blnSubmit = false;
        return false;
    }

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);
        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Update Details';
        return true;
    }
}

function submitComment(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtComment,"Comment");
    reason += validateSpecialCharacters(theForm.txtComment,"Comment");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Comment';

        return true;
    }
}

function submitNotifyRequester(theForm) {
    var reason = "";
    var blnSubmit = true;

    reason += validateEmptyField(theForm.txtNotificationMessage,"Message");
    reason += validateSpecialCharacters(theForm.txtNotificationMessage,"Message");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);

        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Notify';

        return true;
    }
}

function submitProductManagerApproval(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true) {
        theForm.Action.value = 'Approve';

        return true;
    }
}

function submitMarketingApproval(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true) {
        theForm.Action.value = 'Marketing';

        return true;
    }
}
</script>
<%

'-----------------------------------------------
' COPY THIS BRIEF RECORD
'-----------------------------------------------
sub copyBrief
    dim strSQL

    dim strDepartment
    dim strTitle
    dim intPriority
    dim intOutputPrinted
    dim intOutputOnline
    dim strOutputDetails
    dim strFirstDeadline
    dim strSecondDeadline
    dim strDeadline
    dim strImagesLocation
    dim strCopyLocation
    dim strAspect1
    dim strAspect2
    dim strAspect3
    dim strLookFeel
    dim strDescription

    strDepartment       = Trim(session("project_department"))
    strTitle            = Replace(session("project_title"),"'","''")
    intPriority         = Trim(session("project_priority"))
    intOutputPrinted    = Trim(session("project_output_printed"))
    intOutputOnline     = Trim(session("project_output_web"))
    strOutputDetails    = Replace(session("project_output_details"),"'","''")
    strFirstDeadline    = Trim(session("project_first_deadline"))
    strSecondDeadline   = Trim(session("project_second_deadline"))
    strDeadline         = Trim(session("project_deadline"))

    If IsNull(session("project_image_location")) Or Len(session("project_image_location")) = 0 Then
        strImagesLocation = session("project_image_location")
    else
        strImagesLocation = Replace(session("project_image_location"),"'","''")
    end if

    If IsNull(session("project_copy_location")) Or Len(session("project_copy_location")) = 0 Then
        strCopyLocation = session("project_copy_location")
    else
        strCopyLocation = Replace(session("project_copy_location"),"'","''")
    end if

    If IsNull(session("project_aspect_1")) Or Len(session("project_aspect_1")) = 0 Then
        strAspect1 = session("project_aspect_1")
    else
        strAspect1 = Replace(session("project_aspect_1"),"'","''")
    end if

    If IsNull(session("project_aspect_2")) Or Len(session("project_aspect_2")) = 0 Then
        strAspect2 = session("project_aspect_2")
    else
        strAspect2 = Replace(session("project_aspect_2"),"'","''")
    end if

    If IsNull(session("project_aspect_3")) Or Len(session("project_aspect_3")) = 0 Then
        strAspect3 = session("project_aspect_3")
    else
        strAspect3 = Replace(session("project_aspect_3"),"'","''")
    end if

    If IsNull(session("project_look_feel")) Or Len(session("project_look_feel")) = 0 Then
        strLookFeel = session("project_look_feel")
    else
        strLookFeel = Replace(session("project_look_feel"),"'","''")
    end if

    If IsNull(session("project_description")) Or Len(session("project_description")) = 0 Then
        strDescription = session("project_description")
    else
        strDescription = Replace(session("project_description"),"'","''")
    end if

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
    strSQL = strSQL & "'" & strDepartment & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strTitle) & " COPY',"
    strSQL = strSQL & "'" & intPriority & "',"
    strSQL = strSQL & "'" & intOutputPrinted & "',"
    strSQL = strSQL & "'" & intOutputOnline & "',"
    strSQL = strSQL & "'" & Server.HTMLEncode(strOutputDetails) & "',"
    strSQL = strSQL & " CONVERT(datetime,'" & strFirstDeadline & "',103),"
    strSQL = strSQL & " CONVERT(datetime,'" & strSecondDeadline & "',103),"
    strSQL = strSQL & " CONVERT(datetime,'" & strDeadline & "',103),"
    strSQL = strSQL & "'" & strImagesLocation & "',"
    strSQL = strSQL & "'" & strCopyLocation & "',"
    strSQL = strSQL & "'" & strAspect1 & "',"
    strSQL = strSQL & "'" & strAspect2 & "',"
    strSQL = strSQL & "'" & strAspect3 & "',"
    strSQL = strSQL & "'" & strLookFeel & "',"
    strSQL = strSQL & "'" & strDescription & "',"
    strSQL = strSQL & "'" & session("logged_username") & "',1)"

    'response.Write strSQL 
    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        Response.Redirect("thank-you_draft.asp")
    end if

    call CloseDataBase()
end sub

'-----------------------------------------------
' UPDATE BRIEF
'-----------------------------------------------
sub updateBrief
    dim strSQL
    dim intID
    intID = request("id")

    dim strTitle
    dim intPriority
    dim strEmail
    dim intOutputPrinted
    dim intOutputOnline
    dim strOutputDetails
    dim strFirstDeadline
    dim strSecondDeadline
    dim strDeadline
    dim strImagesLocation
    dim strCopyLocation
    dim strAspect1
    dim strAspect2
    dim strAspect3
    dim strLookFeel
    dim strDescription
	
	
    strTitle            = Replace(Request.Form("txtTitle"),"'","''")
    intPriority         = Request.Form("cboPriority")
    strEmail            = Trim(Request.Form("txtEmail"))
    intOutputPrinted    = Request.Form("chkOutputPrinted")
    intOutputOnline     = Request.Form("chkOutputOnline")
    strOutputDetails    = Replace(Request.Form("txtOutputDetails"),"'","''")
    strFirstDeadline    = Trim(Request.Form("txtFirstDeadline"))
    strSecondDeadline   = Trim(Request.Form("txtSecondDeadline"))
    strDeadline         = Trim(Request.Form("txtDeadline"))	
    strImagesLocation   = Replace(Request.Form("txtImagesLocation"),"'","''")
    strCopyLocation     = Replace(Request.Form("txtCopyLocation"),"'","''")
    strAspect1          = Replace(Request.Form("txtAspect1"),"'","''")
    strAspect2          = Replace(Request.Form("txtAspect2"),"'","''")
    strAspect3          = Replace(Request.Form("txtAspect3"),"'","''")
    strLookFeel         = Replace(Request.Form("txtLookFeel"),"'","''")
    strDescription      = Replace(Request.Form("txtDescription"),"'","''")

    Call OpenDataBase()

    strSQL = "UPDATE yma_project SET "
    strSQL = strSQL & "project_title = '" & Server.HTMLEncode(strTitle) & "',"
    strSQL = strSQL & "project_priority = '" & intPriority & "',"
    strSQL = strSQL & "project_output_printed = '" & intOutputPrinted & "',"
    strSQL = strSQL & "project_output_web = '" & intOutputOnline & "',"
    strSQL = strSQL & "project_output_details = '" & Server.HTMLEncode(strOutputDetails) & "',"
    strSQL = strSQL & "project_first_deadline = CONVERT(datetime,'" & strFirstDeadline & "',103),"
    strSQL = strSQL & "project_second_deadline = CONVERT(datetime,'" & strSecondDeadline & "',103),"
    strSQL = strSQL & "project_deadline = CONVERT(datetime,'" & strDeadline & "',103),"
    strSQL = strSQL & "project_image_location = '" & Server.HTMLEncode(strImagesLocation) & "',"
    strSQL = strSQL & "project_copy_location = '" & Server.HTMLEncode(strCopyLocation) & "',"
    strSQL = strSQL & "project_aspect_1 = '" & Server.HTMLEncode(strAspect1) & "',"
    strSQL = strSQL & "project_aspect_2 = '" & Server.HTMLEncode(strAspect2) & "',"
    strSQL = strSQL & "project_aspect_3 = '" & Server.HTMLEncode(strAspect3) & "',"
    strSQL = strSQL & "project_look_feel = '" & Server.HTMLEncode(strLookFeel) & "',"
    strSQL = strSQL & "project_description = '" & Server.HTMLEncode(strDescription) & "',"
    strSQL = strSQL & "project_date_modified = getdate(),"
    strSQL = strSQL & "project_modified_by = '" & session("logged_username") & "' WHERE project_id = " & intID

    'response.Write strSQL

    on error resume next
    conn.Execute strSQL
	
	
    if err <> 0 then
        strMessageText = err.description
		strMessageType="alert alert-danger"
    else
        call addLog(intID,projectModuleID,"Updated brief")
        strMessageText = "The brief has been updated."
		strMessageType="alert alert-success"
    end if

    Call CloseDataBase()
end sub

'-----------------------------------------------
' UPDATE BRIEF DETAILS
'-----------------------------------------------
sub updateBriefDetails
    dim strSQL
    dim intID
    intID = request("id")

    dim strQuote
    dim intProgress
    dim strActualHours
    dim intStatus
    dim strJobNo

    strQuote          = Trim(Request.Form("txtQuote"))
    intProgress       = Trim(Request.Form("cboProgress"))
    strActualHours    = Trim(Request.Form("txtActualHours"))
    intStatus         = Trim(Request.form("cboStatus"))
    strJobNo          = Trim(Request.form("txtJobNo"))

    Call OpenDataBase()

    strSQL = "UPDATE yma_project SET "
    strSQL = strSQL & "project_quote = '" & strQuote & "',"
    strSQL = strSQL & "project_progress = '" & intProgress & "',"
    strSQL = strSQL & "project_actual_hours = '" & strActualHours & "',"
    strSQL = strSQL & "project_status = '" & intStatus & "',"
    strSQL = strSQL & "project_job_no = '" & strJobNo & "',"
    strSQL = strSQL & "project_date_modified = getdate(),"
    strSQL = strSQL & "project_modified_by = '" & session("logged_username") & "' WHERE project_id = " & intID

    'response.Write strSQL

    on error resume next
    conn.Execute strSQL

    if err <> 0 then
        strMessageText = err.description
		strMessageType="alert alert-danger"
    else
        call addLog(intID,projectModuleID,"Updated details")
        strMessageText = "The record has been updated."
		strMessageType="alert alert-success"
    end if

    Call CloseDataBase()
end sub

sub notifyRequester(intID,strRequesterEmail,strEmailMessage)
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

    emailFrom     = "automailer@gmx.yamaha.com"
    emailTo       = trim(strRequesterEmail)
    emailCc       = session("emp_email")
    emailSubject  = "GD Brief - Your approval is needed"

    emailBodyText = "Hi " & session("requester_firstname") & "," & vbCrLf _
                  & " " & vbCrLf _
                  & "Your approval is needed for the brief: " & session("project_title") & "." & vbCrLf _
                  & " " & vbCrLf _
                  & "Comments: " & strEmailMessage & vbCrLf _
                  & "by: " & session("logged_username") & vbCrLf _
                  & " " & vbCrLf _
                  & "Please click on the below link to approve it:" & vbCrLf _
                  & "http://intranet:96/update_brief.asp?id=" & intID & "" & vbCrLf _
                  & " " & vbCrLf _
                  & "Thank you. (This is an automated email - please do not reply to this email)"

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

    strMessageText = "The notification has been sent to the requester."
    call addLog(intID,projectModuleID,"Notified Prod Mgr")
end sub

sub main
    dim intID
    intID = request("id")

    dim strEmailMessage
    strEmailMessage = Replace(Request.Form("txtNotificationMessage"),"'","''")

    'response.write "<br>Email requester: " & session("requester_email")

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        select case Trim(Request.Form("Action"))
            case "Update"
                call updateBrief
            case "Copy"
                call copyBrief
            case "Update Details"
                call updateBriefDetails
            case "Comment"
                call addComment(intID,projectModuleID)
            case "Notify"
                call notifyRequester(intID,session("requester_email"),strEmailMessage)
            case "Approve"
                call approveProductManager(intID,session("logged_username"),session("project_department"))
            case "Marketing"
                call approveMarketingManager(intID,session("requester_email"),session("logged_username"))
        end select
    end if

    call getBrief(intID)
    call listComments(intID,projectModuleID)
    call listLogs(intID,projectModuleID)
    call getRequesterDetails(session("project_created_by"))
end sub

call main

dim strMessageText
dim strCommentsList
dim strLogsList
dim strMessageType
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
    <!-- #include file="include/header.asp" -->
    <tr>
        <td valign="top" class="maincontent">
            <% if session("project_status") = "1" then %>
            <h2>Cannot update this brief as it has not been submitted yet</h2>
            <% else %>
            <table width="1200" border="0">
                <tr>
                    <td width="15%">
                        <img src="images/backward_arrow.gif" border="0" /> <a href="default.asp">Back to Home</a>
                        <h1>Brief ID: <u><%= session("project_id") %></u></h1>
                    </td>
                    <td width="60%">
                        <table cellpadding="4" cellspacing="0" class="created_table">
                            <tr>
                                <td class="created_column_1"><strong>Department:</strong></td>
                                <td class="created_column_2"><h2><%= session("project_department") %></h2></td>
                                <td class="created_column_3">&nbsp;</td>
                            </tr>
                            <tr>
                                <td><strong>Created by:</strong></td>
                                <td><a href="mailto:<%= session("requester_email") %>"><%= session("project_created_by") %></a></td>
                                <td><%= displayDateFormatted(session("project_date_created")) %></td>
                            </tr>
                            <tr>
                                <td><strong>Last modified by:</strong></td>
                                <td><%= session("project_modified_by") %></td>
                                <td><%= displayDateFormatted(session("project_date_modified")) %></td>
                            </tr>
                        </table>
                    </td>
                    <td valign="top" width="25%">
                        <p>
						<% if (Len(strMessageText)>0) then %>
						<div class="<%= strMessageType %>">
							<%= strMessageText %>						
						</div>
						<% end if%>
						</p>
                    </td>
                </tr>
            </table>
            <br />
            <form action="" method="post" name="form_copy_shipment" id="form_copy_shipment" onsubmit="return copyBrief(this)">
                <p>
                    <input type="hidden" name="Action" />
                    <input type="submit" value="Copy this brief" />
                </p>
            </form>
            <table cellspacing="0" cellpadding="0" width="1200">
                <tr>
                    <td width="45%" style="padding-right:15px;" valign="top">
                        <form action="" method="post" name="form_update_project" id="form_update_project" onsubmit="return validateFormOnSubmit(this)">
                            <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                                <tr>
                                    <td colspan="3" class="item_maintenance_header">Brief</td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Title<span class="mandatory">*</span>:</strong>
                                        <br />
                                        <input name="txtTitle" type="text" id="txtTitle" size="70" maxlength="60" value="<%= session("project_title") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Priority:</strong>
                                        <select name="cboPriority">
                                            <option <% if session("project_priority") = "1" then Response.Write " selected" end if %> value="1">Low</option>
                                            <option <% if session("project_priority") = "2" then Response.Write " selected" end if%> value="2">Medium</option>
                                            <option <% if session("project_priority") = "3" then Response.Write " selected" end if %> value="3">High</option>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="30%">
                                        <strong>Output:</strong>
                                    </td>
                                    <td width="35%">
                                        <input type="checkbox" name="chkOutputPrinted" id="chkOutputPrinted" value="1" <% if session("project_output_printed") = "1" then Response.Write " checked" end if%> /> Printed
                                    </td>
                                    <td width="35%">
                                        <input type="checkbox" name="chkOutputOnline" id="chkOutputOnline" value="1" <% if session("project_output_web") = "1" then Response.Write " checked" end if%> /> Online
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Details<span class="mandatory">*</span>:</strong>
                                        <br />
                                        <input type="text" id="txtOutputDetails" name="txtOutputDetails" maxlength="70" size="80" value="<%= session("project_output_details") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <strong>1<sup>st</sup> draft deadline<span class="mandatory">*</span>:</strong>
                                        <br />
                                        <input type="text" id="txtFirstDeadline" name="txtFirstDeadline" maxlength="10" size="10" value="<%= session("project_first_deadline") %>" />
                                        <span class="mandatory"><em>DD/MM/YYYY</em></span>
                                    </td>
                                    <td>
                                        <strong>2<sup>nd</sup> draft deadline:</strong>
                                        <br />
                                        <input type="text" id="txtSecondDeadline" name="txtSecondDeadline" maxlength="10" size="10" value="<%= session("project_second_deadline") %>" />
                                        <span class="mandatory"><em>DD/MM/YYYY</em></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Printing / publishing deadline<span class="mandatory">*</span>:</strong>
                                        <br />
                                        <input type="text" id="txtDeadline" name="txtDeadline" maxlength="10" size="10" value="<%= session("project_deadline") %>" />
                                        <span class="mandatory"><em>DD/MM/YYYY</em></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Images location:</strong>
                                        <br />
                                        <input type="text" id="txtImagesLocation" name="txtImagesLocation" maxlength="200" size="120" value="<%= session("project_image_location") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Copy location:</strong>
                                        <br />
                                        <input type="text" id="txtCopyLocation" name="txtCopyLocation" maxlength="200" size="120" value="<%= session("project_copy_location") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Most important aspects:</strong>
                                        <br />
                                        1. <input type="text" id="txtAspect1" name="txtAspect1" maxlength="30" size="40" value="<%= Session("project_aspect_1") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        2. <input type="text" id="txtAspect2" name="txtAspect2" maxlength="30" size="40" value="<%= Session("project_aspect_2") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        3. <input type="text" id="txtAspect3" name="txtAspect3" maxlength="30" size="40" value="<%= Session("project_aspect_3") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Up to 3 words that describe the look and feel:</strong>
                                        <br />
                                        <input type="text" id="txtLookFeel" name="txtLookFeel" maxlength="60" size="70" value="<%= Session("project_look_feel") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <strong>Publishing / Printing requirement:</strong>
                                        <br />
										
										<!--
										
                                        <textarea name="txtDescription" id="txtDescription" cols="70" rows="6"><%= session("project_description") %></textarea>
										-->
										  
									    <textarea class="form-control" name="txtDescription" id="txtDescription" cols="100%" rows="10"><%= session("project_description") %></textarea>
										  
                                        <br />
                                      <!-- 
									  onKeyDown="limitText(this.form.txtDescription,this.form.countdown,500);" onKeyUp="limitText(this.form.txtDescription,this.form.countdown,500);"
									  <em>Max: 500 characters</em>
									  -->
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <input type="hidden" name="Action" />
                                        <input type="submit" value="Update Brief" <% if session("project_status") = "0" then Response.Write "disabled" end if%> />
                                    </td>
                                </tr>
                            </table>
                        </form>
                    </td>
                    <td valign="top" width="55%">
                        <form action="" method="post" name="form_update_project_details" id="form_update_project_details" onsubmit="return validateDetailsFormOnSubmit(this)">
                            <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                                <tr>
                                    <td colspan="2" class="item_maintenance_header">Details (GD only)</td>
                                </tr>
                                <tr>
                                    <td width="50%">
                                        <strong>Quote:</strong>
                                        <input type="text" id="txtQuote" name="txtQuote" maxlength="4" size="5" value="<%= session("project_quote") %>" /> hour(s)
                                    </td>
                                    <td width="50%">
                                        <strong>Progress:</strong>
                                        <select name="cboProgress">
                                            <option <% if session("project_progress") = "0" then Response.Write " selected" end if %> value="0">0</option>
                                            <option <% if session("project_progress") = "10" then Response.Write " selected" end if %> value="10">10</option>
                                            <option <% if session("project_progress") = "20" then Response.Write " selected" end if %> value="20">20</option>
                                            <option <% if session("project_progress") = "30" then Response.Write " selected" end if %> value="30">30</option>
                                            <option <% if session("project_progress") = "40" then Response.Write " selected" end if %> value="40">40</option>
                                            <option <% if session("project_progress") = "50" then Response.Write " selected" end if %> value="50">50</option>
                                            <option <% if session("project_progress") = "60" then Response.Write " selected" end if %> value="60">60</option>
                                            <option <% if session("project_progress") = "70" then Response.Write " selected" end if %> value="70">70</option>
                                            <option <% if session("project_progress") = "80" then Response.Write " selected" end if %> value="80">80</option>
                                            <option <% if session("project_progress") = "90" then Response.Write " selected" end if %> value="90">90</option>
                                            <option <% if session("project_progress") = "100" then Response.Write " selected" end if%> value="100">100</option>
                                        </select>
                                        %
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <strong>Actual time spent:</strong>
                                        <input type="text" id="txtActualHours" name="txtActualHours" maxlength="4" size="5" value="<%= session("project_actual_hours") %>" /> hour(s)
                                    </td>
                                </tr>
                                <tr class="status_row">
                                    <td colspan="2">
                                        <strong>Status:</strong>
                                        <select name="cboStatus">
                                            <option <% if session("project_status") = "2" then Response.Write " selected" end if%> value="2" style="color:red">Submitted</option>
                                            <option <% if session("project_status") = "3" then Response.Write " selected" end if%> value="3" style="color:red">Viewed</option>
                                            <option <% if session("project_status") = "4" then Response.Write " selected" end if%> value="4" style="color:orange">Concept</option>
                                            <option <% if session("project_status") = "5" then Response.Write " selected" end if%> value="5" style="color:orange">Draft</option>
                                            <option <% if session("project_status") = "6" then Response.Write " selected" end if%> value="6" style="color:orange">Changes</option>
                                            <option <% if session("project_status") = "7" then Response.Write " selected" end if%> value="7" style="color:green">Pending Approval</option>
                                            <option <% if session("project_status") = "8" then Response.Write " selected" end if%> value="8" style="color:green">On-hold</option>
                                            <option <% if session("project_status") = "1" then Response.Write " disabled" end if %> <% if session("project_status") = "0" then Response.Write " selected" end if%> value="0" style="color:green">Completed</option>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        Job no: <input type="text" id="txtJobNo" name="txtJobNo" maxlength="12" size="15" value="<%= session("project_job_no") %>" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2"><input type="hidden" name="Action" />
                                        <% if session("logged_username") = "jaclynw" or session("logged_username") = "jamieb" then %>
                                        <input type="submit" value="Update Details" <% if session("project_status") = "0" then Response.Write "disabled" end if%> />
                                        <% end if %>
                                    </td>
                                </tr>
                            </table>
                        </form>
                        <h3>Record Log:</h3>
                        <table cellpadding="5" cellspacing="0" border="0" class="item_maintenance_box">
                            <tr>
                                <td><%= session("project_created_by") %></td>
                                <td>Created</td>
                                <td><%= displayDateFormatted(session("project_date_created")) %></td>
                            </tr>
                            <%= strLogsList %>
                        </table>
                        <br />
                        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                            <tr>
                                <td colspan="2" class="item_maintenance_header">Approvals</td>
                            </tr>
                            <tr>
                                <td colspan="2"><strong>Notify <%= session("requester_firstname") %> (Requester) for approval:</strong></td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <form action="" method="post" onsubmit="return submitNotifyRequester(this)">
                                        Message:
                                        <input type="text" id="txtNotificationMessage" name="txtNotificationMessage" maxlength="50" size="50" />
                                        <input type="hidden" name="Action" />
                                        <input type="submit" value="Notify" />
                                    </form>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2"></td>
                            </tr>
                            <tr>
                                <td width="30%" valign="top"><strong>Requester:</strong></td>
                                <td width="70%">
                                    <%
                                        select case session("product_manager_approval")
                                            case 0
                                    %>
                                    <form action="" method="post" onsubmit="return submitProductManagerApproval(this)">
                                        <input type="hidden" name="Action" />
                                        <input type="submit" value="Approve" style="color:green" <% if session("logged_username") <> session("requester_username") then Response.Write("disabled") end if %> />
                                    </form>
                                    <%
                                            case 1
                                    %>
                                    <font color="green">APPROVED</font> by <%= session("product_manager_approval_by") %>
                                    <br />
                                    <%= displayDateFormatted(session("product_manager_approval_date")) %>
                                    <%
                                        end select
                                    %>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2"></td>
                            </tr>
                            <tr>
                                <td valign="top"><strong>Marketing Manager:</strong></td>
                                <td>
                                    <%
                                        select case session("marketing_manager_approval")
                                            case 0
                                                if (session("project_department") = "YMEC" and session("logged_username") = "carolyns") or (session("project_department") <> "YMEC") then
                                    %>
                                    <form action="" method="post" onsubmit="return submitMarketingApproval(this)">
                                        <input type="hidden" name="Action" />
                                        <input type="submit" value="Approve" style="color:green" />
                                    </form>
                                    <%
                                                else
                                    %>
                                    <input type="submit" value="Approve" disabled="disabled" />
                                    <%
                                                end if
                                            case 1
                                    %>
                                    <font color="green">APPROVED</font> by <%= session("marketing_manager_approval_by") %>
                                    <br />
                                    <%= displayDateFormatted(session("marketing_manager_approval_date")) %>
                                    <%
                                        end select
                                    %>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <h3>
                Comments
                <br />
                <img src="images/comment_bar.jpg" border="0" />
            </h3>
            <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
                <%= strCommentsList %>
                <tr>
                    <td>
                        <form action="" name="form_add_comment" id="form_add_comment" method="post" onsubmit="return submitComment(this)">
                            <p>
                                <input type="text" name="txtComment" id="txtComment" maxlength="200" size="150" />
                                <input type="hidden" name="Action" />
                                <input type="submit" value="Add Comment" />
                            </p>
                        </form>
                    </td>
                </tr>
            </table>
          <% end if %>
        </td>
    </tr>
</table>

<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">
    var picker = new Pikaday({
        field: document.getElementById('txtDeadline'),
        firstDay: 1,
        minDate: new Date('2012-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
        format: 'DD/MM/YYYY'
    });
	
	
	 var picker = new Pikaday({
        field: document.getElementById('txtFirstDeadline'),
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
        format: 'DD/MM/YYYY'
    });

    var picker = new Pikaday({
        field: document.getElementById('txtSecondDeadline'),
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
        format: 'DD/MM/YYYY'
    });

    var picker = new Pikaday({
        field: document.getElementById('txtDeadline'),
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
        format: 'DD/MM/YYYY'
    });
	
</script>
</body>
</html>