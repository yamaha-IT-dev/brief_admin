<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp" -->
<!--#include file="class/clsBrief.asp" -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!--#include file="class/clsLog.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>New Customer Service Ticket</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link rel="stylesheet" href="include/pikaday.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;	
		
	reason += validateEmptyField(theForm.cboDepartment,"Department");
	
	reason += validateEmptyField(theForm.txtTitle,"Title");
	reason += validateSpecialCharacters(theForm.txtTitle,"Title");

	reason += validateEmptyField(theForm.txtOutputDetails,"Output details");
	reason += validateSpecialCharacters(theForm.txtOutputDetails,"Output details");
	
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
        theForm.Action.value = 'Add';		
		return true;
    }
}
</script>
<%
sub setSessionVariables
	Session("new_project_department")		= Request.Form("cboDepartment")
	Session("new_project_title")			= Trim(Request.Form("txtTitle"))
	Session("new_project_priority")			= Request.Form("cboPriority")
	Session("new_project_output_printed") 	= Request.Form("chkOutputPrinted")
	Session("new_project_output_web") 		= Request.Form("chkOutputOnline")
	Session("new_project_output_details")  	= Trim(Request.Form("txtOutputDetails"))
	Session("new_project_first_deadline")	= Trim(Request.Form("txtFirstDeadline"))
	Session("new_project_second_deadline")	= Trim(Request.Form("txtSecondDeadline"))
	Session("new_project_deadline")			= Trim(Request.Form("txtDeadline"))
	Session("new_project_image_location") 	= Trim(Request.Form("txtImagesLocation"))
	Session("new_project_copy_location") 	= Trim(Request.Form("txtCopyLocation"))
	Session("new_project_aspect_1")  		= Trim(Request.Form("txtAspect1"))
	Session("new_project_aspect_2")  		= Trim(Request.Form("txtAspect2"))
	Session("new_project_aspect_3")  		= Trim(Request.Form("txtAspect3"))
	Session("new_project_look_feel")  		= Trim(Request.Form("txtLookFeel"))
	Session("new_project_description")  	= Trim(Request.Form("txtDescription"))
	Session("new_project_gl_code")  		= Trim(Request.Form("txtGLcode"))
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request.Form("Action"))
			case "Add"
				call setSessionVariables
				response.Redirect("confirm_brief.asp")
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
    <td valign="top" class="maincontent"><img src="images/backward_arrow.gif" border="0" /> <a href="default.asp">Back to Home</a>
      <h2>New Customer Service Ticket</h2>
      <p><font color="red"><%= strMessageText %></font></p>
      <form action="" method="post" name="form_add_project" id="form_add_project" onsubmit="return validateFormOnSubmit(this)">
      <table width="1024">
          <tr>
            <td valign="top"><table cellpadding="5" cellspacing="0" class="form_box">
                <tr>
                  <td colspan="3" class="form_header">Customer Details</td>
                </tr>
                <tr>
                  <td>First name<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtFirstName" name="txtFirstName" maxlength="30" size="35" value="<%= session("new_firstname") %>" /></td>
                  <td colspan="2">Last name<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtLastName" name="txtLastName" maxlength="30" size="35" value="<%= session("new_lastname") %>" /></td>
                </tr>
                
                <tr>
                  <td colspan="3">Address<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtAddress" name="txtAddress" maxlength="50" size="60" value="<%= session("new_address") %>" /></td>
                </tr>
                <tr>
                  <td width="50%">City<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtCity" name="txtCity" maxlength="30" size="35" value="<%= session("new_city") %>" /></td>
                  <td width="20%">State<span class="mandatory">*</span>:<br />
                    <select name="cboState">
                      <%= strStateList %>
                    </select></td>
                  <td width="30%">Postcode<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtPostcode" name="txtPostcode" maxlength="4" size="5" value="<%= session("new_postcode") %>" /></td>
                </tr>
                <tr>
                  <td colspan="3">Email<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtEmail" name="txtEmail" maxlength="60" size="70" value="<%= session("new_email") %>" /></td>
                </tr>
                <tr>
                  <td>Phone no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtPhone" name="txtPhone" maxlength="12" size="15" value="<%= session("new_phone") %>" /></td>
                  <td colspan="2">Company:<br />
                    <input type="text" id="txtCompany" name="txtCompany" maxlength="30" size="35" value="<%= session("new_company") %>" /></td>
                </tr>
              </table></td>
            <td valign="top"><table cellpadding="5" cellspacing="0" class="form_box">
                <tr>
                  <td colspan="2" class="form_header">Purchase Details</td>
                </tr>
                <tr>
                  <td width="50%">Model no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtModelNo" name="txtModelNo" maxlength="25" size="35" value="<%= session("new_model_no") %>" /></td>
                  <td width="50%">Serial no<span class="mandatory">*</span>:<br />
                  <input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="15" size="20" value="<%= session("new_serial_no") %>" /></td>
                </tr>
                <tr>
                  <td>Invoice no<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="15" size="20" value="<%= session("new_invoice_no") %>" /></td>
                  <td>Date purchased<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtDatePurchased" name="txtDatePurchased" maxlength="10" size="10" value="<%= session("new_date_purchased") %>" />
                    <em>DD/MM/YYYY</em></td>
                </tr>
                <tr>
                  <td colspan="2">Dealer name<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtDealer" name="txtDealer" maxlength="25" size="35" value="<%= session("new_dealer") %>" /></td>
                </tr>
                
                <tr>
                  <td>Dealer contact name:<br />
                    <input type="text" id="txtAccessories" name="txtAccessories" maxlength="30" size="30" value="<%= session("new_accessories") %>" /></td>
                  <td>Dealer phone:<br />
                  <input type="text" id="txtAccessories3" name="txtAccessories3" maxlength="12" size="15" value="<%= session("new_accessories") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Fault<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtFault" name="txtFault" maxlength="30" size="50" value="<%= session("new_fault") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Comments:<br />
                  <textarea name="txtDescription2" id="txtDescription2" cols="60" rows="6" onkeydown="limitText(this.form.txtDescription,this.form.countdown,500);" 
onkeyup="limitText(this.form.txtDescription,this.form.countdown,500);"><%= Session("new_project_description") %></textarea></td>
                </tr>
                <tr>
                  <td colspan="2">Status:<br />
                    <select name="cboStatus">
                      <%= strJobStatusList %>
                    </select></td>
                </tr>
            </table></td>
            <td valign="top"><table cellpadding="5" cellspacing="0" class="form_box">
                <tr>
                  <td width="100%" colspan="2" class="form_header">ASC Details</td>
                </tr>
                <tr>
                  <td colspan="2">ASC name<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtDealer" name="txtDealer" maxlength="25" size="35" value="<%= session("new_dealer") %>" /></td>
                </tr>
                <tr>
                  <td width="100%">ASC contact name<span class="mandatory">*</span>:<br />
                  <input type="text" id="txtFault2" name="txtFault2" maxlength="30" size="50" value="<%= session("new_fault") %>" /></td>
                  <td width="100%">ASC phone:<br />
                  <input type="text" id="txtAccessories2" name="txtAccessories2" maxlength="12" size="15" value="<%= session("new_accessories") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Comments:<br />
                  <textarea name="txtDescription2" id="txtDescription2" cols="60" rows="6" onkeydown="limitText(this.form.txtDescription,this.form.countdown,500);" 
onkeyup="limitText(this.form.txtDescription,this.form.countdown,500);"><%= Session("new_project_description") %></textarea></td>
                </tr>
                <tr>
                  <td colspan="2">Status:<br />
                    <select name="cboStatus">
                      <%= strJobStatusList %>
                    </select></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td colspan="3" align="center"><br /><input type="hidden" name="Action" />
          <input type="submit" value="Submit" /></td>
            </tr>
        </table>
        <br />
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" name="confirm_button" value="Confirm" />
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