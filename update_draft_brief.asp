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
<script type="text/javascript" src="../include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtTitle,"Title");
	//reason += validateSpecialCharacters(theForm.txtTitle,"Title");

	reason += validateEmptyField(theForm.txtOutputDetails,"Output details");
	//reason += validateSpecialCharacters(theForm.txtOutputDetails,"Output details");
	
	reason += validateFirstDeadline(theForm.txtFirstDeadline);
	
	reason += validateDeadline(theForm.txtDeadline);
	
	reason += validateEmptyField(theForm.txtGLcode,"GL Code");
	reason += validateSpecialCharacters(theForm.txtGLcode,"GL Code");
	
	reason += validateEmptyField(theForm.txtImagesLocation,"Images Location");
	//reason += validateSpecialCharacters(theForm.txtImagesLocation,"Images Location");
	
	reason += validateEmptyField(theForm.txtCopyLocation,"Copy Location");
	//reason += validateSpecialCharacters(theForm.txtCopyLocation,"Copy Location");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Update';		
		return true;
    }
}

function validateDetailsFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateNumeric(theForm.txtQuote,"Quote");
	reason += validateSpecialCharacters(theForm.txtQuote,"Quote");
	//reason += validateNumeric(theForm.txtActualHours,"Actual hours");
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

	if (blnSubmit == true){
        theForm.Action.value = 'Update Details';
		return true;
    }
}

function submitComment(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtComment,"Comment");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}
	
	if (blnSubmit == true){
		theForm.Action.value = 'Comment';
		
		return true;		
    }
}

function confirmBrief(theForm) {
	if (confirm ("Please click OK to confirm.")){
		theForm.Action.value = 'Confirm';
		return true;
    } else {
		return false;
	}
}
</script>
<%
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
	dim strGLcode
			
	strTitle 				= Replace(Request.Form("txtTitle"),"'","''")
	intPriority 			= Request.Form("cboPriority")
	strEmail  				= Trim(Request.Form("txtEmail"))
	intOutputPrinted 		= Request.Form("chkOutputPrinted")
	intOutputOnline 		= Request.Form("chkOutputOnline")
	strOutputDetails 		= Replace(Request.Form("txtOutputDetails"),"'","''")
	strFirstDeadline		= Trim(Request.Form("txtFirstDeadline"))
	strSecondDeadline		= Trim(Request.Form("txtSecondDeadline"))
	strDeadline				= Trim(Request.Form("txtDeadline"))	
	strImagesLocation 		= Replace(Request.Form("txtImagesLocation"),"'","''")
	strCopyLocation 		= Replace(Request.Form("txtCopyLocation"),"'","''")
	strAspect1 				= Replace(Request.Form("txtAspect1"),"'","''")
	strAspect2 				= Replace(Request.Form("txtAspect2"),"'","''")
	strAspect3 				= Replace(Request.Form("txtAspect3"),"'","''")
	strLookFeel 			= Replace(Request.Form("txtLookFeel"),"'","''")
	strDescription 			= Replace(Request.Form("txtDescription"),"'","''")
	strGLcode 				= Replace(Request.Form("txtGLcode"),"'","''")
	
	Call OpenDataBase()
	
	strSQL = "UPDATE yma_project SET "
	strSQL = strSQL & "project_title = '" & Server.HTMLEncode(strTitle) & "',"
	strSQL = strSQL & "project_priority = '" & intPriority & "',"
	strSQL = strSQL & "project_gl_code = '" & Server.HTMLEncode(strGLcode) & "',"
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
	else
		call addLog(intID,projectModuleID,"Updated plan")
		strMessageText = "The draft has been updated."
	end if
	
	Call CloseDataBase()
end sub

sub main
	dim intID
	intID 	= request("id")
	
	call getBrief(intID)
	call listLogs(intID,projectModuleID)
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		select case Trim(Request.Form("Action"))
			case "Update"
				call updateBrief
				call getBrief(intID)
				call listLogs(intID,projectModuleID)
			case "Confirm"
				response.Redirect("confirm_draft_brief.asp")
		end select
	end if
end sub

call main

dim strMessageText
dim strCommentsList
dim strLogsList
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td valign="top" class="maincontent"><table width="1200" border="0">
        <tr>
          <td width="15%"><img src="images/backward_arrow.gif" border="0" /> <a href="default.asp">Back to Home</a>
            <h2>Update Plan - <%= session("project_department") %></h2></td>
          <td width="60%"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Department:</strong></td>
                <td class="created_column_2"><%= session("project_department") %></td>
                <td class="created_column_3">&nbsp;</td>
              </tr>
              <tr>
                <td class="created_column_1"><strong>Created by:</strong></td>
                <td class="created_column_2"><%= session("project_created_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("project_date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified by:</strong></td>
                <td><%= session("project_modified_by") %></td>
                <td><%= displayDateFormatted(session("project_date_modified")) %></td>
              </tr>
            </table></td>
          <td valign="top" width="25%"><p><font color="red"><%= strMessageText %></font></p></td>
        </tr>
      </table>
      <br />
      <table cellspacing="0" cellpadding="0" width="1200">
        <tr>
          <td width="45%" style="padding-right:15px;"><form action="" method="post" name="form_update_project" id="form_update_project" onsubmit="return validateFormOnSubmit(this)">
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="3" class="item_maintenance_header">Brief</td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Title<span class="mandatory">*</span>:</strong><br />
                    <input name="txtTitle" type="text" id="txtTitle" size="70" maxlength="60" value="<%= session("project_title") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Priority:</strong><select name="cboPriority">
                      <option <% if session("project_priority") = "1" then Response.Write " selected" end if %> value="1">Low</option>
                      <option <% if session("project_priority") = "2" then Response.Write " selected" end if%> value="2">Medium</option>
                      <option <% if session("project_priority") = "3" then Response.Write " selected" end if %> value="3">High</option>
                    </select></td>
                </tr>
                <tr>
                  <td width="30%"><strong>Output:</strong></td>
                  <td width="35%"><input type="checkbox" name="chkOutputPrinted" id="chkOutputPrinted" value="1" <% if session("project_output_printed") = "1" then Response.Write " checked" end if%> />
                    Printed</td>
                  <td width="35%"><input type="checkbox" name="chkOutputOnline" id="chkOutputOnline" value="1" <% if session("project_output_web") = "1" then Response.Write " checked" end if%> />
                    Online</td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Details<span class="mandatory">*</span>:</strong><br />
                    <input type="text" id="txtOutputDetails" name="txtOutputDetails" maxlength="70" size="80" value="<%= session("project_output_details") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2"><strong>1<sup>st</sup> draft deadline<span class="mandatory">*</span>:</strong><br />
                    <input type="text" id="txtFirstDeadline" name="txtFirstDeadline" maxlength="10" size="10" value="<%= session("project_first_deadline") %>" />
                    <span class="mandatory"><em>DD/MM/YYYY</em></span></td>
                  <td><strong>2<sup>nd</sup> draft deadline:</strong><br />
                    <input type="text" id="txtSecondDeadline" name="txtSecondDeadline" maxlength="10" size="10" value="<%= session("project_second_deadline") %>" />
                    <span class="mandatory"><em>DD/MM/YYYY</em></span></td>
                </tr>
                <tr>
                  <td colspan="2"><strong>Printing / publishing deadline<span class="mandatory">*</span>:</strong><br />
                    <input type="text" id="txtDeadline" name="txtDeadline" maxlength="10" size="10" value="<%= session("project_deadline") %>" />
                    <span class="mandatory"><em>DD/MM/YYYY</em></span></td>
                  <td><strong>GL code<span class="mandatory">*</span>:</strong><br />
                    <input type="text" id="txtGLcode" name="txtGLcode" maxlength="12" size="15" value="<%= session("project_gl_code") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Image(s) supplied location<span class="mandatory">*</span>:</strong><br />
                    <input type="text" id="txtImagesLocation" name="txtImagesLocation" maxlength="80" size="90" value="<%= session("project_image_location") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Copy supplied location<span class="mandatory">*</span>:</strong><br />
                    <input type="text" id="txtCopyLocation" name="txtCopyLocation" maxlength="80" size="90" value="<%= session("project_copy_location") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Most important aspects:</strong><br />
                    1.
                    <input type="text" id="txtAspect1" name="txtAspect1" maxlength="30" size="40" value="<%= Session("project_aspect_1") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3">2.
                    <input type="text" id="txtAspect2" name="txtAspect2" maxlength="30" size="40" value="<%= Session("project_aspect_2") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3">3.
                    <input type="text" id="txtAspect3" name="txtAspect3" maxlength="30" size="40" value="<%= Session("project_aspect_3") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Up to 3 words that describe the look and feel:</strong><br />
                    <input type="text" id="txtLookFeel" name="txtLookFeel" maxlength="60" size="70" value="<%= Session("project_look_feel") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3"><strong>Publishing / Printing requirement:</strong><br />
                    <textarea name="txtDescription" id="txtDescription" cols="70" rows="6" onKeyDown="limitText(this.form.txtDescription,this.form.countdown,500);" 
onKeyUp="limitText(this.form.txtDescription,this.form.countdown,500);"><%= session("project_description") %></textarea>
                    <br />
                    <em>Max: 500 characters</em></td>
                </tr>
                <tr>
                  <td colspan="3"><input type="hidden" name="Action" />
                    <input type="submit" value="Update Draft" <% if session("project_status") = "0" then Response.Write "disabled" end if%> /></td>
                </tr>
              </table>
            </form>
            <form action="" method="post" name="form_confirm_brief" id="form_confirm_brief" onsubmit="return confirmBrief(this)">
              <p>
                <input type="hidden" name="Action" />
                <input type="submit" value="Submit" />
              </p>
            </form></td>
          <td valign="top" width="55%"><h3>Record Log:</h3>
            <table cellpadding="5" cellspacing="0" border="0" class="item_maintenance_box">
              <tr>
                <td><%= session("project_created_by") %></td>
                <td>Created plan</td>
                <td><%= displayDateFormatted(session("project_date_created")) %></td>
              </tr>
              <%= strLogsList %>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script> 
<script type="text/javascript" src="include/pikaday.js"></script> 
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDeadline'),		
        firstDay: 1,
        minDate: new Date('2012-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
		format: 'DD/MM/YYYY'
    });			
</script>
</body>
</html>