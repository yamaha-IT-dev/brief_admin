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
<title>New Brief</title>
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
    reason += validateFirstDeadline(theForm.txtFirstDeadline);
    reason += validateDeadline(theForm.txtDeadline);
    reason += validateEmptyField(theForm.txtImagesLocation,"Images Location");
    reason += validateEmptyField(theForm.txtCopyLocation,"Copy Location");

    if (reason != "") {
        alert("Some fields need correction:\n" + reason);
        blnSubmit = false;
        return false;
    }

    if (blnSubmit == true) {
        theForm.Action.value = 'Add';
        return true;
    }
}
</script>
<%
sub setSessionVariables
    Session("new_project_department")       = Request.Form("cboDepartment")
    Session("new_project_title")            = Trim(Request.Form("txtTitle"))
    Session("new_project_priority")         = Request.Form("cboPriority")
    Session("new_project_output_printed")   = Request.Form("chkOutputPrinted")
    Session("new_project_output_web")       = Request.Form("chkOutputOnline")
    Session("new_project_output_details")   = Trim(Request.Form("txtOutputDetails"))
    Session("new_project_first_deadline")   = Trim(Request.Form("txtFirstDeadline"))
    Session("new_project_second_deadline")  = Trim(Request.Form("txtSecondDeadline"))
    Session("new_project_deadline")         = Trim(Request.Form("txtDeadline"))
    Session("new_project_image_location")   = Trim(Request.Form("txtImagesLocation"))
    Session("new_project_copy_location")    = Trim(Request.Form("txtCopyLocation"))
    Session("new_project_aspect_1")         = Trim(Request.Form("txtAspect1"))
    Session("new_project_aspect_2")         = Trim(Request.Form("txtAspect2"))
    Session("new_project_aspect_3")         = Trim(Request.Form("txtAspect3"))
    Session("new_project_look_feel")        = Trim(Request.Form("txtLookFeel"))
    Session("new_project_description")      = Trim(Request.Form("txtDescription"))
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
        <td valign="top" class="maincontent">
            <img src="images/backward_arrow.gif" border="0" /><a href="default.asp">Back to Home</a>
            <h2>New Brief</h2>
            <p><font color="red"><%= strMessageText %></font></p>
            <form action="" method="post" name="form_add_project" id="form_add_project" onsubmit="return validateFormOnSubmit(this)">
                <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                    <tr>
                        <td colspan="3" class="item_maintenance_header">Brief Details</td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Department<span class="mandatory">*</span>:</strong>
                            <select name="cboDepartment">
                                <option value="">...</option>
                                <option <% if Session("new_project_department") = "MPD" then Response.Write " selected" end if %> value="MPD">MPD</option>
                                <option <% if Session("new_project_department") = "MPD - PRO" then Response.Write " selected" end if %> value="MPD - PRO">PRO</option>
                                <option <% if Session("new_project_department") = "MPD - TRAD" then Response.Write " selected" end if %> value="MPD - TRAD">TRAD</option>
                                <option <% if Session("new_project_department") = "CA" then Response.Write " selected" end if %> value="CA">CA</option>
                                <option <% if Session("new_project_department") = "YMEC" then Response.Write " selected" end if %> value="YMEC">YMEC</option>
                                <option <% if Session("new_project_department") = "AV" then Response.Write " selected" end if %> value="AV">AV</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Title of Project<span class="mandatory">*</span>:</strong>
                            <br />
                            <input type="text" id="txtTitle" name="txtTitle" maxlength="60" size="70" value="<%= Session("new_project_title") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Priority<span class="mandatory">*</span>:</strong>
                            <select name="cboPriority">
                                <option <% if Session("new_project_priority") = "1" then Response.Write " selected" end if %> value="1">Low</option>
                                <option <% if Session("new_project_priority") = "2" then Response.Write " selected" end if%> value="2">Medium</option>
                                <option <% if Session("new_project_priority") = "3" then Response.Write " selected" end if %> value="3">High</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td width="20%">
                            <strong>Output:</strong>
                        </td>
                        <td width="30%">
                            <input type="checkbox" name="chkOutputPrinted" id="chkOutputPrinted" value="1" <% if Session("new_project_output_printed") = "1" then Response.Write " checked" end if%> /> Printed
                        </td>
                        <td width="50%">
                            <input type="checkbox" name="chkOutputOnline" id="chkOutputOnline" value="1" <% if Session("new_project_output_web") = "1" then Response.Write " checked" end if%> /> Online
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Details<span class="mandatory">*</span>: <em>(eg. Qty / Size)</em></strong>
                            <br />
                            <input type="text" id="txtOutputDetails" name="txtOutputDetails" maxlength="70" size="80" value="<%= Session("new_project_output_details") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <strong>1<sup>st</sup> draft deadline<span class="mandatory">*</span>:</strong>
                            <br />
                            <input type="text" id="txtFirstDeadline" name="txtFirstDeadline" maxlength="10" size="10" value="<%= Session("new_project_first_deadline") %>" />
                            <span class="mandatory"><em>DD/MM/YYYY</em></span>
                        </td>
                        <td>
                            <strong>2<sup>nd</sup> draft deadline:</strong>
                            <br />
                            <input type="text" id="txtSecondDeadline" name="txtSecondDeadline" maxlength="10" size="10" value="<%= Session("new_project_second_deadline") %>" />
                            <span class="mandatory"><em>DD/MM/YYYY</em></span>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Printing / publishing deadline<span class="mandatory">*</span>:</strong>
                            <br />
                            <input type="text" id="txtDeadline" name="txtDeadline" maxlength="10" size="10" value="<%= Session("new_project_deadline") %>" />
                            <span class="mandatory"><em>DD/MM/YYYY</em></span>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Images location<span class="mandatory">*</span>: (e.g. <u>GD Resource\MPD\2013\WIP\Facebook</u>)</strong>
                            <br />
                            <input type="text" id="txtImagesLocation" name="txtImagesLocation" maxlength="200" size="120" value="<%= Session("new_project_image_location") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Copy location<span class="mandatory">*</span>: (e.g. <u>GD Resource\MPD\2013\WIP\Facebook</u>)</strong>
                            <br />
                            <input type="text" id="txtCopyLocation" name="txtCopyLocation" maxlength="200" size="120" value="<%= Session("new_project_copy_location") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Most important aspects:</strong>
                            <br />
                            1. <input type="text" id="txtAspect1" name="txtAspect1" maxlength="30" size="40" value="<%= Session("new_project_aspect_1") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            2. <input type="text" id="txtAspect2" name="txtAspect2" maxlength="30" size="40" value="<%= Session("new_project_aspect_2") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            3. <input type="text" id="txtAspect3" name="txtAspect3" maxlength="30" size="40" value="<%= Session("new_project_aspect_3") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <strong>Up to 3 words that describe the look and feel:</strong>
                            <br />
                            <input type="text" id="txtLookFeel" name="txtLookFeel" maxlength="60" size="70" value="<%= Session("new_project_look_feel") %>" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="5">
                           


						   <strong>Publishing / Printing requirement:</strong>
                            <br />
							<!--
							onKeyDown="limitText(this.form.txtDescription,this.form.countdown,500);" onKeyUp="limitText(this.form.txtDescription,this.form.countdown,500);"
							-->
                            
							<textarea class="form-control" name="txtDescription" id="txtDescription" cols="100%" rows="10"><%= Session("new_project_description") %></textarea>
							
                            <br />
                          <!-- 
						  <em>Max: 500 characters</em>
						  -->
                        </td>
						
						
                    </tr>
                </table>
                <p>
                    <input type="hidden" name="Action" />
                    <input type="submit" name="confirm_button" value="Confirm" />
                    <input type="reset" name="reset" id="reset" value="Reset" />
                </p>
            </form>
        </td>
    </tr>
</table>

<script type="text/javascript" src="include/moment.js"></script>
<script type="text/javascript" src="include/pikaday.js"></script>
<script type="text/javascript">
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