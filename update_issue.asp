<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsExcelIssue.asp " -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<% strSection = "" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Warehouse Return</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="../include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
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
        theForm.Action.value = 'Update';

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
</script>
<%
function displayDateFormatted(strDateInput)	
	if IsNull(strDateInput) or strDateInput = "01/01/1900" or strDateInput = "1/1/1900"  then 
		Response.Write "N/A"
	else
		Response.Write "" & WeekDayName(WeekDay(strDateInput)) & ", " & FormatDateTime(strDateInput,1) & " at " & FormatDateTime(strDateInput,3)	
	end if
end function

sub main
	dim intID
	intID 	= request("id")
	
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

	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateIssue				
				call getReasonCode
			case "Comment"
				call addComment(intID,warehouseReturnModuleID)
		end select
	end if
	
	call getIssue(intID)
	call listComments(intID,warehouseReturnModuleID)
end sub

call main

dim strMessageText
dim strCommentsList
dim strReasonCodeList
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <tr>
    <td valign="top" class="maincontent"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_issues.asp">Back to List</a>
            <h2>Update Excel Issues</h2>
            <font color="red"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("issCreatedBy") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("issDateCreated")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("issModifiedBy") %></td>
                <td><%= displayDateFormatted(session("issDateModified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_quarantine" id="form_update_quarantine" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td width="50%">ASC<span class="mandatory">*</span>:<br />
              <input type="text" id="txtASC" name="txtASC" maxlength="20" size="30" value="<%= Server.HTMLEncode(session("issASC")) %>" /></td>
            <td width="50%">ASC Contact name<span class="mandatory">*</span>:<br />
              <input type="text" id="txtContactName" name="txtContactName" maxlength="20" size="30" value="<%= Server.HTMLEncode(session("issContactName")) %>" /></td>
          </tr>
          <tr>
            <td colspan="2">Product<span class="mandatory">*</span>:<br />
              <input type="text" id="txtProduct" name="txtProduct" maxlength="30" size="40" value="<%= session("issProduct") %>" /></td>
          </tr>
          <tr>
            <td colspan="2">ASC Reported Fault<span class="mandatory">*</span>:<br />
              <input type="text" id="txtReportedFault" name="txtReportedFault" maxlength="100" size="100" value="<%= session("issReportedFault") %>" /></td>
          </tr>
          <tr>
            <td colspan="2">Excel Diagnosed Fault<span class="mandatory">*</span>:<br />
              <input type="text" id="txtDiagnosedFault" name="txtDiagnosedFault" maxlength="100" size="100" value="<%= session("issDiagnosedFault") %>" /></td>
          </tr>
          <tr>
            <td colspan="2">Reason for Return<span class="mandatory">*</span>:<br />
              <input type="text" id="txtReason" name="txtReason" maxlength="50" size="60" value="<%= session("issReason") %>" /></td>
          </tr>
          <tr>
            <td>Return date<span class="mandatory">*</span>:<br />
              <input type="text" id="txtReturnDate" name="txtReturnDate" maxlength="15" size="20" value="<%= session("issReturnDate") %>" /></td>
            <td>Return con-note<span class="mandatory">*</span>:<br />
              <input type="text" id="txtReturnConnote" name="txtReturnConnote" maxlength="15" size="20" value="<%= session("issReturnConnote") %>" /></td>
          </tr>
          <tr>
            <td colspan="2">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="50" rows="4"><%= session("issComments") %></textarea></td>
          </tr>
          <tr>
            <td colspan="2">Spare Parts used:<br />
              <input type="text" id="txtSpareParts" name="txtSpareParts" maxlength="30" size="35" value="<%= session("issSpareParts") %>" /></td>
          </tr>
          <tr>
            <td>Dispatch date:<br />
              <input type="text" id="txtDispatchDate" name="txtDispatchDate" maxlength="15" size="20" value="<%= session("issDispatchDate") %>" /></td>
            <td>Dispatch con-note:<br />
              <input type="text" id="txtDispatchConnote" name="txtDispatchConnote" maxlength="15" size="20" value="<%= session("issDispatchConnote") %>" /></td>
          </tr>
          <tr>
            <td colspan="2"><input type="hidden" name="Action" />
              <input type="submit" value="Update" /></td>
          </tr>
        </table>
      </form>
      <h2>Comments <br />
        <img src="images/comment_bar.jpg" border="0" /></h2>
      <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
        <%= strCommentsList %>
        <tr>
          <td><form action="" name="form_add_comment" id="form_add_comment" method="post" onsubmit="return submitComment(this)">
              <p>
                <input type="text" name="txtComment" id="txtComment" maxlength="60" size="65" />
                <input type="hidden" name="Action" />
                <input type="submit" value="Add Comment" />
              </p>
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>