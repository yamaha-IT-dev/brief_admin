<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Staff Launch Survey</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link rel="stylesheet" href="include/pikaday.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateRadioButton(fld,errormsg) {
    var error = "";
	
    if ((fld[0].checked == false) && (fld[1].checked == false) && (fld[2].checked == false) && (fld[3].checked == false) && (fld[4].checked == false)) {
       error = "- " + errormsg + " is empty.\n"
    } 
    return error;
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;	
		
	reason += validateRadioButton(theForm.radQuestion1,"Question 1");
	reason += validateRadioButton(theForm.radQuestion2,"Question 2");
	reason += validateRadioButton(theForm.radQuestion3,"Question 3");
	
	//reason += validateEmptyField(theForm.txtFutureComments,"Future Staff Launch?");	
	
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
sub addSurvey
	dim intOverall
	dim strOverallComments
	dim intDemo
	dim strDemoComments
	dim intFocus
	dim strFocusComments
	dim strFutureComments
	
	intOverall 			= Request.Form("radQuestion1")
	strOverallComments 	= Replace(Request.Form("txtOverallComments"),"'","''")
	intDemo 			= Request.Form("radQuestion2")
	strDemoComments 	= Replace(Request.Form("txtDemoComments"),"'","''")
	intFocus 			= Request.Form("radQuestion3")
	strFocusComments 	= Replace(Request.Form("txtFocusComments"),"'","''")
	strFutureComments 	= Replace(Request.Form("txtFutureComments"),"'","''")	
	
	call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_survey ("
	strSQL = strSQL & "	overall, "
	strSQL = strSQL & " overall_comments, "
	strSQL = strSQL & " demo, "
	strSQL = strSQL & " demo_comments, "
	strSQL = strSQL & " focus, "
	strSQL = strSQL & " focus_comments, "
	strSQL = strSQL & " future_comments, "
	strSQL = strSQL & " created_by "	
	strSQL = strSQL & ") VALUES ("
	strSQL = strSQL & "'" & Trim(intOverall) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strOverallComments) & "',"
	strSQL = strSQL & "'" & Trim(intDemo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strDemoComments) & "',"
	strSQL = strSQL & "'" & Trim(intFocus) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strFocusComments) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strFutureComments) & "',"	
	strSQL = strSQL & "'" & session("logged_username") & "')"
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else		
		Response.Redirect("survey_confirm.asp")
	end if	

	call CloseDataBase()
end sub

sub main
	session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		select case Trim(Request.Form("Action"))
			case "Add"
				call addSurvey	
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
    <td valign="top" class="maincontent">
      <h1>Staff Launch Survey</h1>
      <p><em>Logged in as: <%= session("logged_username") %></em></p>
      <p><font color="red"><%= strMessageText %></font></p>
      <form action="" method="post" name="form_add_survey" id="form_add_survey" onsubmit="return validateFormOnSubmit(this)">
      <table width="800" border="0" cellspacing="1" cellpadding="4">
          <tr>
            <td>&nbsp;</td>
            <td align="center"><em>Poor</em></td>
            <td align="center">&nbsp;</td>
            <td align="center"><em>Good</em></td>
            <td align="center">&nbsp;</td>
            <td align="center"><em>Excellent</em></td>
          </tr>
          <tr>
            <td width="60%">&nbsp;</td>
            <td width="8%" align="center"><strong>1</strong></td>
            <td width="8%" align="center"><strong>2</strong></td>
            <td width="8%" align="center"><strong>3</strong></td>
            <td width="8%" align="center"><strong>4</strong></td>
            <td width="8%" align="center"><strong>5</strong></td>
          </tr>
          <tr>
            <td class="form_column">1. How would you rate the overall day?</td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion1" id="radQuestion1" value="1" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion1" id="radQuestion1" value="2" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion1" id="radQuestion1" value="3" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion1" id="radQuestion1" value="4" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion1" id="radQuestion1" value="5" /></td>
          </tr>
          <tr>
            <td colspan="6" class="survey_column">Any comments?<br />
            <textarea name="txtOverallComments" id="txtOverallComments" cols="55" rows="4" onkeydown="limitText(this.form.txtOverallComments,this.form.countdown,400);" 
onkeyup="limitText(this.form.txtOverallComments,this.form.countdown,400);"></textarea></td>
          </tr>
          <tr>
            <td class="form_column">2. How would you rate the product demonstrations?</td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion2" id="radQuestion2" value="1" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion2" id="radQuestion2" value="2" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion2" id="radQuestion2" value="3" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion2" id="radQuestion2" value="4" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion2" id="radQuestion2" value="5" /></td>
          </tr>
          <tr>
            <td colspan="6" class="survey_column">Any suggestions for the next launch?<br />
            <textarea name="txtDemoComments" id="txtDemoComments" cols="55" rows="4" onkeydown="limitText(this.form.txtDemoComments,this.form.countdown,400);" 
onkeyup="limitText(this.form.txtDemoComments,this.form.countdown,400);"></textarea>
            </td>
          </tr>
          <tr>
            <td class="form_column">3. How would you rate the Group Ideas Session?</td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion3" id="radQuestion3" value="1" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion3" id="radQuestion3" value="2" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion3" id="radQuestion3" value="3" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion3" id="radQuestion3" value="4" /></td>
            <td align="center" class="form_column"><input type="radio" name="radQuestion3" id="radQuestion3" value="5" /></td>
          </tr>
          <tr>
            <td colspan="6" class="survey_column">Any suggestions for the next launch?<br />
            <textarea name="txtFocusComments" id="txtFocusComments" cols="55" rows="4" onkeydown="limitText(this.form.txtFocusComments,this.form.countdown,400);" 
onkeyup="limitText(this.form.txtFocusComments,this.form.countdown,400);"></textarea></td>
          </tr>
          <tr>
            <td colspan="6" class="form_column">What would you like to see in future staff launch events?</td>
          </tr>
          <tr>
            <td colspan="6" class="survey_column">
            <textarea name="txtFutureComments" id="txtFutureComments" cols="55" rows="4" onkeydown="limitText(this.form.txtFutureComments,this.form.countdown,400);" 
onkeyup="limitText(this.form.txtFutureComments,this.form.countdown,400);"></textarea></td>
          </tr>
        </table>
        <p>
          <input type="hidden" name="Action" />
          <input type="submit" name="submit" value="Submit" />
        </p>
      </form></td>
  </tr>
</table>
</body>
</html>