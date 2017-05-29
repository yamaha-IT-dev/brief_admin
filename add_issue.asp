<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsExcelIssue.asp" -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>New Excel Issue and Return Log</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link rel="stylesheet" href="include/pikaday.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;	
	
	//reason += validateEmptyField(theForm.txtASC,"ASC");
	//reason += validateSpecialCharacters(theForm.txtASC,"ASC");

	//reason += validateEmptyField(theForm.txtContactName,"ASC Contact Name");
	//reason += validateSpecialCharacters(theForm.txtContactName,"ASC Contact Name");
	
	reason += validateEmptyField(theForm.txtProduct,"Product");
	reason += validateSpecialCharacters(theForm.txtProduct,"Product");
	
	reason += validateEmptyField(theForm.txtReportedFault,"ASC Reported Fault");
	reason += validateSpecialCharacters(theForm.txtReportedFault,"ASC Reported Fault");
	
	reason += validateEmptyField(theForm.txtDiagnosedFault,"Diagnosed Fault");
	reason += validateSpecialCharacters(theForm.txtDiagnosedFault,"Diagnosed Fault");
	
	//reason += validateEmptyField(theForm.txtReason,"Reason for return");
	reason += validateSpecialCharacters(theForm.txtReason,"Reason for return");
	
	reason += validateSpecialCharacters(theForm.txtReturnConnote,"Return con-note");
	
	reason += validateSpecialCharacters(theForm.txtComments,"Comments");
	
	reason += validateSpecialCharacters(theForm.txtSpareParts,"Spare Parts used");
	
	reason += validateSpecialCharacters(theForm.txtDispatchConnote,"Dispatch con-note");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Add';		
		return true;
    }
}
</script>
<%
sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request.Form("Action"))
			case "Add"
				Dim issASC
				Dim issContactName
				Dim issProduct
				Dim issReportedFault
				Dim issDiagnosedFault
				Dim issReason
				Dim issReturnDate
				Dim issReturnConnote
				Dim issComments
				Dim issSpareParts
				Dim issDispatchDate
				Dim issDispatchConnote
				Dim issStatus
				
				issASC				= Trim(Request.Form("txtASC"))
				issContactName		= Trim(Request.Form("txtContactName"))
				issProduct			= Trim(Request.Form("txtProduct"))
				issReportedFault	= Trim(Request.Form("txtReportedFault"))
				issDiagnosedFault 	= Trim(Request.Form("txtDiagnosedFault"))
				issReason  			= Trim(Request.Form("txtReason"))
				issReturnDate		= Trim(Request.Form("txtReturnDate"))
				issReturnConnote	= Trim(Request.Form("txtReturnConnote"))
				issComments			= Trim(Request.Form("txtComments"))
				issSpareParts 		= Trim(Request.Form("txtSpareParts"))
				issDispatchDate  	= Trim(Request.Form("txtDispatchDate"))
				issDispatchConnote  = Trim(Request.Form("txtDispatchConnote"))
				issStatus  			= Trim(Request.Form("cboStatus"))
				
				call addIssue(issASC, issContactName, issProduct, issReportedFault, issDiagnosedFault, issReason, issReturnDate, issReturnConnote, issComments, issSpareParts, issDispatchDate, issDispatchConnote, issStatus, session("logged_username"))
		end select
	end if
end sub

call main

dim strMessageText
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <tr>
    <td valign="top" class="maincontent"><img src="images/backward_arrow.gif" border="0" /> <a href="list_issues.asp">Back to List</a>
      <h2>New Excel Issue &amp; Return Log</h2>
      <p><font color="red"><%= strMessageText %></font></p>
      <form action="" method="post" name="form_add_issue" id="form_add_issue" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td width="50%"><strong>ASC:</strong><br />
              <input type="text" id="txtASC" name="txtASC" maxlength="20" size="30" /></td>
            <td width="50%"><strong>ASC Contact name:</strong><br />
              <input type="text" id="txtContactName" name="txtContactName" maxlength="20" size="30" /></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Product<span class="mandatory">*</span>:</strong><br />
              <input type="text" id="txtProduct" name="txtProduct" maxlength="30" size="40" /></td>
          </tr>
          <tr>
            <td colspan="2"><strong>ASC Reported Fault<span class="mandatory">*</span>:</strong><br />
              <input type="text" id="txtReportedFault" name="txtReportedFault" maxlength="100" size="100" /></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Excel Diagnosed Fault<span class="mandatory">*</span>:</strong><br />
              <input type="text" id="txtDiagnosedFault" name="txtDiagnosedFault" maxlength="100" size="100" /></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Reason for Return:</strong><br />
              <input type="text" id="txtReason" name="txtReason" maxlength="50" size="60" /></td>
          </tr>
          <tr>
            <td><strong>Return date:</strong><br />
              <input type="text" id="txtReturnDate" name="txtReturnDate" maxlength="15" size="20" /></td>
            <td><strong>Return con-note:</strong><br />
              <input type="text" id="txtReturnConnote" name="txtReturnConnote" maxlength="15" size="20" /></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Comments:</strong><br />
              <textarea name="txtComments" id="txtComments" cols="50" rows="4"></textarea></td>
          </tr>
          <tr>
            <td colspan="2"><strong>Spare Parts used:</strong><br />
              <input type="text" id="txtSpareParts" name="txtSpareParts" maxlength="30" size="35" /></td>
          </tr>
          <tr>
            <td><strong>Dispatch date:</strong><br />
              <input type="text" id="txtDispatchDate" name="txtDispatchDate" maxlength="15" size="20" /></td>
            <td><strong>Dispatch con-note:</strong><br />
              <input type="text" id="txtDispatchConnote" name="txtDispatchConnote" maxlength="15" size="20" /></td>
          </tr>
          <tr>
            <td colspan="2"><input type="hidden" name="Action" />
              <input type="submit" value="Submit" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script> 
<script type="text/javascript" src="include/pikaday.js"></script> 
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtReturnDate'),		
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
		format: 'DD/MM/YYYY'
    });
	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtDispatchDate'),		
        firstDay: 1,
        minDate: new Date('2013-01-01'),
        maxDate: new Date('2030-12-31'),
        yearRange: [2013,2030],
		format: 'DD/MM/YYYY'
    });				
</script>
</body>
</html>