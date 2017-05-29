<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsLoanRenewal.asp " -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<% strSection = "" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Loan Stock</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="../include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtLocation,"Location");
	reason += validateSpecialCharacters(theForm.txtLocation,"Location");
	
	reason += validateSpecialCharacters(theForm.txtComments,"Comments");

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
							
	loanAccount			= Trim(Request("account"))
	loanProduct			= Trim(Request("product"))
	loanSerialNo		= Trim(Request("serial"))
	loanQty				= Trim(Request("qty"))
	loanLIC				= Trim(Request("lic"))
	loanDate 			= Trim(Request("date"))
	loanFirstExpiryDate = DateAdd("m", 3, loanDate)
	loanFinalExpiryDate = DateAdd("m", 6, loanDate)
	loanOrderNo			= Trim(Request("order"))
	loanShipmentNo		= Trim(Request("ship"))	

	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateIssue				
				call getReasonCode
			case "Comment"
				call addComment(intID,warehouseReturnModuleID)
		end select
	end if
	
	'call getIssue(intID)
	'call listComments(intID,warehouseReturnModuleID)
end sub

call main

Dim loanAccount
Dim loanProduct
Dim loanSerialNo
Dim loanQty
Dim loanLIC
Dim loanDate
Dim loanFirstExpiryDate
Dim loanFinalExpiryDate
Dim loanOrderNo
Dim loanShipmentNo

dim strMessageText
dim strCommentsList
dim strReasonCodeList
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <tr>
    <td valign="top" class="maincontent"><p><img src="images/backward_arrow.gif" border="0" /> <a href="loan.asp">Back to List</a></p> <font color="red"><%= strMessageText %></font>
      <form action="" method="post" name="form_update_loan" id="form_update_loan" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td width="30%"><strong>Loan account:</strong></td>
            <td width="70%"><%= loanAccount %></td>
          </tr>
          <tr>
            <td><strong>Product:</strong></td>
            <td><%= loanProduct %></td>
          </tr>
          <tr>
            <td><strong>Serial no:</strong></td>
            <td><%= loanSerialNo %></td>
          </tr>
          <tr>
            <td><strong>Qty:</strong></td>
            <td><%= loanQty %></td>
          </tr>
          <tr>
            <td><strong>LIC:</strong></td>
            <td>$<%= FormatNumber(loanLIC) %></td>
          </tr>
          <tr>
            <td><strong>Loan date:</strong></td>
            <td><%= FormatDateTime(loanDate,1) %></td>
          </tr>
          <tr>
            <td><strong>First expiry date:</strong></td>
            <td><%= FormatDateTime(loanFirstExpiryDate,1) %></td>
          </tr>
          <tr>
            <td><strong>Final expiry date:</strong></td>
            <td><%= FormatDateTime(loanFinalExpiryDate,1) %></td>
          </tr>
          <tr>
            <td><strong>Order no:</strong></td>
            <td><%= loanOrderNo %></td>
          </tr>
          <tr>
            <td><strong>Shipment no:</strong></td>
            <td><%= loanShipmentNo %></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>Location<span class="mandatory">*</span>:<br /></td>
            <td><input type="text" id="txtLocation" name="txtLocation" maxlength="20" size="25" value="<%= session("loanLocation") %>" /></td>
          </tr>
          <tr>
            <td>Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="35" rows="4"><%= session("loanComments") %></textarea></td>
          </tr>          
          <tr>
            <td></td>
            <td><input type="hidden" name="Action" />
              <input type="submit" value="Update Loan Stock" /></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>