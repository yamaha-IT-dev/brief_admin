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
<title>Save as Draft Brief</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link rel="stylesheet" href="include/pikaday.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.cboDepartment,"Department");
	
	reason += validateEmptyField(theForm.txtTitle,"Title");
	
	reason += validateEmptyField(theForm.txtOutputDetails,"Output details");
	
	//reason += validateFirstDeadline(theForm.txtFirstDeadline);
	reason += validateDeadline(theForm.txtDeadline);
	
	reason += validateEmptyField(theForm.txtGLcode,"GL Code");
	reason += validateSpecialCharacters(theForm.txtGLcode,"GL Code");
	
	reason += validateEmptyField(theForm.txtImagesLocation,"Images Location");
	
	reason += validateEmptyField(theForm.txtCopyLocation,"Copy Location");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Save';		
		return true;
    }
}

</script>
<%
'-----------------------------------------------
' SAVE AS DRAFT
'-----------------------------------------------
sub draftBrief
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
	dim strGLcode
	
	strDepartment	  		= Trim(Request.Form("cboDepartment"))
	strTitle 				= Replace(Request.Form("txtTitle"),"'","''")
	intPriority 			= request("cboPriority")
	intOutputPrinted 		= request("chkOutputPrinted")
	intOutputOnline 		= request("chkOutputOnline")
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
	strGLcode  				= Trim(Request.Form("txtGLcode"))
		
	call OpenDataBase()
		
	strSQL = "INSERT INTO yma_project ("
	strSQL = strSQL & "	project_department, "
	strSQL = strSQL & " project_title, "
	strSQL = strSQL & " project_priority, "
	strSQL = strSQL & " project_output_printed, "
	strSQL = strSQL & " project_output_web, "
	strSQL = strSQL & " project_output_details, "
	strSQL = strSQL & " project_first_deadline, "
	strSQL = strSQL & " project_second_deadline, "
	strSQL = strSQL & " project_deadline, "
	strSQL = strSQL & " project_image_location, "
	strSQL = strSQL & " project_copy_location, "
	strSQL = strSQL & " project_aspect_1, "
	strSQL = strSQL & " project_aspect_2, "
	strSQL = strSQL & " project_aspect_3, "
	strSQL = strSQL & " project_look_feel, "
	strSQL = strSQL & " project_description, "
	strSQL = strSQL & " project_gl_code, "
	strSQL = strSQL & " project_created_by, "
	strSQL = strSQL & " project_status) VALUES ("
	strSQL = strSQL & "'" & strDepartment & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strTitle) & "',"
	strSQL = strSQL & "'" & intPriority & "',"
	strSQL = strSQL & "'" & intOutputPrinted & "',"
	strSQL = strSQL & "'" & intOutputOnline & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strOutputDetails) & "',"
	strSQL = strSQL & " CONVERT(datetime,'" & strFirstDeadline & "',103),"
	strSQL = strSQL & " CONVERT(datetime,'" & strSecondDeadline & "',103),"
	strSQL = strSQL & " CONVERT(datetime,'" & strDeadline & "',103),"
	strSQL = strSQL & "'" & Server.HTMLEncode(strImagesLocation) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strCopyLocation) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strAspect1) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strAspect2) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strAspect3) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strLookFeel) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strDescription) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strGLcode) & "',"
	strSQL = strSQL & "'" & session("logged_username") & "', 1)"
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		call addLog(intID,projectModuleID,"Created draft")
		'strMessageText = "The brief has been saved."
		Response.Redirect("thank-you_draft.asp")
	end if 

	call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request.Form("Action"))
			case "Save"
				call draftBrief
		end select
	end if
end sub

call main

dim strMessageText
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td valign="top" class="maincontent">
    <img src="images/backward_arrow.gif" border="0" /> <a href="default.asp">Back to Home</a>
    <h2>Plan a brief (Save as draft)</h2>
      <p><font color="red"><%= strMessageText %></font></p>
      <form action="" method="post" name="form_add_project" id="form_add_project" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="3" class="item_maintenance_header">Brief Details</td>
          </tr>
          <tr>
            <td colspan="3"><strong>Department<span class="mandatory">*</span>:</strong><select name="cboDepartment">
                <option value="">...</option>
                <!--<option value="AV">AV</option>
                <option value="CC">CC</option>
                <option value="OP">OP</option>-->
                <option value="MPD">MPD</option>
                <option value="MPD - PRO">PRO</option>
                <option value="MPD - TRAD">TRAD</option>
                <option value="CA">CA</option>
                <option value="YMEC">YMEC</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="3"><strong>Title of Project<span class="mandatory">*</span>:</strong><br />
              <input type="text" id="txtTitle" name="txtTitle" maxlength="60" size="70" /></td>
          </tr>
          <tr>
            <td colspan="3"><strong>Priority<span class="mandatory">*</span>:</strong><select name="cboPriority">
                <option value="1">Low</option>
                <option value="2">Medium</option>
                <option value="3">High</option>
              </select></td>
          </tr>
          <tr>
            <td width="20%"><strong>Output:</strong></td>
            <td width="30%"><input type="checkbox" name="chkOutputPrinted" id="chkOutputPrinted" value="1" />
              Printed</td>
            <td width="50%"><input type="checkbox" name="chkOutputOnline" id="chkOutputOnline" value="1" />
              Online</td>
          </tr>
          <tr>
            <td colspan="3"><strong>Details<span class="mandatory">*</span>: <em>(eg. Qty / Size)</em></strong><em></em><br />
              <input type="text" id="txtOutputDetails" name="txtOutputDetails" maxlength="70" size="80" /></td>
          </tr>
          <tr>
            <td colspan="2"><strong>1<sup>st</sup> draft deadline:</strong><br />
              <input type="text" id="txtFirstDeadline" name="txtFirstDeadline" maxlength="10" size="10" />
              <span class="mandatory"><em>DD/MM/YYYY</em></span></td>
            <td><strong>2<sup>nd</sup> draft deadline:</strong><br />
              <input type="text" id="txtSecondDeadline" name="txtSecondDeadline" maxlength="10" size="10" />
              <span class="mandatory"><em>DD/MM/YYYY</em></span></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Printing / publishing deadline<span class="mandatory">*</span>:</strong><br />
              <input type="text" id="txtDeadline" name="txtDeadline" maxlength="10" size="10" />
              <span class="mandatory"><em>DD/MM/YYYY</em></span></td>
            <td><strong>GL code<span class="mandatory">*</span>: (eg. 31-52100-390)</strong><br />
              <input type="text" id="txtGLcode" name="txtGLcode" maxlength="12" size="15" /></td>
          </tr>
          <tr>
            <td colspan="3"><strong>Image(s) supplied location<span class="mandatory">*</span>: (eg. <u>GD Resource\AV\2012\WIP\Facebook</u>)</strong><br />
              <input type="text" id="txtImagesLocation" name="txtImagesLocation" maxlength="80" size="80" /></td>
          </tr>
          <tr>
            <td colspan="3"><strong>Copy supplied location<span class="mandatory">*</span>: (eg. <u>GD Resource\AV\2012\WIP\Facebook</u>)</strong><br />
              <input type="text" id="txtCopyLocation" name="txtCopyLocation" maxlength="80" size="80" /></td>
          </tr>
          <tr>
            <td colspan="3"><strong>Most important aspects:</strong><br />1. <input type="text" id="txtAspect1" name="txtAspect1" maxlength="30" size="40" /></td>
          </tr>
          <tr>
            <td colspan="3">2. <input type="text" id="txtAspect2" name="txtAspect2" maxlength="30" size="40" /></td>
          </tr>
          <tr>
            <td colspan="3">3. <input type="text" id="txtAspect3" name="txtAspect3" maxlength="30" size="40" /></td>
          </tr>
          <tr>
            <td colspan="3"><strong>Up to 3 words that describe the look and feel:</strong><br /><input type="text" id="txtLookFeel" name="txtLookFeel" maxlength="60" size="70" /></td>
          </tr>
          <tr>
            <td colspan="3"><strong>Publishing / Printing requirement:</strong><br />
              <textarea name="txtDescription" id="txtDescription" cols="60" rows="6" onKeyDown="limitText(this.form.txtDescription,this.form.countdown,500);" 
onKeyUp="limitText(this.form.txtDescription,this.form.countdown,500);"></textarea>
              <br />
              <em>Max: 500 characters</em></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" value="Save as Draft" />
          <input type="reset" name="reset" id="reset" value="Reset" />
        </p>
      </form></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtFirstDeadline'),		
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
		format: 'DD/MM/YYYY'
    });	
	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtSecondDeadline'),		
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
		format: 'DD/MM/YYYY'
    });	
	
	var picker = new Pikaday(
    {
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