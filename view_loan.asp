<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="include/loan_functions.asp " -->
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!--#include file="class/clsLoanBase.asp " -->
<!--#include file="class/clsLoanInfo.asp " -->
<!--#include file="class/clsLoanRenewal.asp " -->
<!--#include file="class/clsLoanRenewalApproval.asp " -->
<% strSection = "" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>View Loan Stock</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="../include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateRenewalLength(fld,message) {
    var error = "";
	
    if ((fld[0].checked == false) && (fld[1].checked == false)) {
       
        error = "- Please select the length.\n"
    } 
    return error;
}

function validateAddForm(theForm) {
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
        theForm.Action.value = 'Add';

		return true;
    }
}

function validateUpdateStockInfoForm(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtUpdateLocation,"Location");
	reason += validateSpecialCharacters(theForm.txtUpdateLocation,"Location");
	
	reason += validateEmptyField(theForm.txtUpdateSerialNo,"Serial no");
	reason += validateSpecialCharacters(theForm.txtUpdateSerialNo,"Serial no");
	
	reason += validateSpecialCharacters(theForm.txtUpdateComments,"Comments");

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

function validateRenewForm(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	//reason += validateRenewalLength(theForm.radLength);
	
	reason += validateEmptyField(theForm.txtRenewComments,"Comments");
	reason += validateSpecialCharacters(theForm.txtRenewComments,"Comments");		

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Renew';

		return true;
    }
}

function validateBasketStockInfoForm(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Basket';

		return true;
    }
}

function submitApproval(theForm) {
	var blnSubmit = true;

	if (blnSubmit == true){
		theForm.Action.value = 'Approve';
		
		return true;
		
    }
}

function submitRejection(theForm) {
	var blnSubmit = true;

	if (blnSubmit == true){
		theForm.Action.value = 'Reject';
		
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
	session("loanLocation") 	= ""
	session("loanComments") 	= ""
	session("loan_serial_no") 	= ""
	session("loan_qty") 		= ""
	
	stockID			= Request("stockID")		
	
	loanOrderNo		= Trim(Request("order"))
	loanOrderLine	= Trim(Request("line"))
	
	loanLocation 	= Trim(Replace(Request.Form("txtLocation"),"'","''"))
	loanSerialNo 	= Trim(Replace(Request.Form("txtSerial"),"'","''"))
	loanComments 	= Trim(Replace(Request.Form("txtComments"),"'","''"))
	
	serialNo 		= Trim(Replace(Request.Form("txtSerialNo"),"'","''"))
	
	'UPDATE LOAN STOCK INFO
	updateLocation 	= Trim(Replace(Request.Form("txtUpdateLocation"),"'","''"))
	updateSerialNo 	= Trim(Replace(Request.Form("txtUpdateSerialNo"),"'","''"))
	updateComments 	= Trim(Replace(Request.Form("txtUpdateComments"),"'","''"))	
	
	'BASKET
	stockLocation	= Request("stockLocation")
	stockSerialNo	= Request("stockSerialNo")
	
	'RENEW
	renLength 	= Trim(Request("radLength"))
	renComments = Trim(Replace(Request.Form("txtRenewComments"),"'","''"))	
	
	'APPROVAL
	renewalID 			= Trim(Request("renID"))
	renCreatedByEmail 	= Trim(Request("renCreatedByEmail"))
	
	call getLoanStock(loanOrderNo, loanOrderLine)
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Add"
			    call addLoanStockInfo(loanOrderNo, loanOrderLine, loanLocation, loanSerialNo, loanComments, session("logged_username"))
			case "Update"
				call updateLoanStockInfo(stockID, updateLocation, updateSerialNo, updateComments, session("logged_username"))
			case "Renew"
				call addLoanStockRenewal(loanOrderNo, loanOrderLine, renComments, session("logged_username"), session("emp_email"))
			case "Basket"
				call addLoanStockToBasket(session("loan_product"), stockSerialNo, session("loan_lic"), stockLocation, session("loan_account"))
			case "Approve"
				call approveRenewal(renewalID, renCreatedByEmail, session("logged_username"))
			case "Reject"
				call rejectRenewal(renewalID, renCreatedByEmail, session("logged_username"))			
		end select
	end if	
	
	'call getLoanStockRenewal(loanOrderNo, loanOrderLine)
	call listLoanStockInfo(loanOrderNo, loanOrderLine)
	call listLoanRenewal(loanOrderNo, loanOrderLine)
	intSerialNo = Trim(session("loan_serial_no"))
	intQty 		= Trim(session("loan_qty"))	
	
 	'intDayDiff = DateDiff("d",session("loan_date"), strTodayDate)
      		
	'response.write session("loan_date_diff")
	'response.write "intQty: " & intQty
	'response.write "loan_serial_no: " & intSerialNo
end sub

call main

Dim stockID
Dim loanOrderNo
Dim loanOrderLine
Dim serialNo

Dim strLoanStockInfoList

Dim strMessageText
Dim intSerialNo
Dim intQty

Dim updateLocation
Dim	updateSerialNo
Dim	updateComments

Dim stockLocation
Dim stockSerialNo

dim strTodayDate
strTodayDate = FormatDateTime(Date())

'Dim intDayDiff
Dim renLength
Dim renComments

Dim renewalID
Dim renCreatedByEmail

Dim strLoanRenewalList
%>
</head>
<body>
<p align="center"><a href="loan_summary.asp">Loan Summary</a> <img src="images/forward_arrow.gif" /> <a href="loan_user.asp?account=<%= session("loan_user_account") %>">Loan Stock</a> <img src="images/forward_arrow.gif" /> View Loan Stock Info</p>
<table border="0" cellspacing="0" cellpadding="5" align="center">
  <tr>
    <td valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
        <tr>
          <td colspan="2" class="item_maintenance_header">Loan Stock (from BASE)</td>
        </tr>
        <tr>
          <td width="30%" align="right"><strong>Loan account:</strong></td>
          <td width="70%"><%= session("loan_account") %></td>
        </tr>
        <tr>
          <td align="right"><strong>Product:</strong></td>
          <td><%= session("loan_product") %></td>
        </tr>
        <tr>
          <td align="right"><strong>Serial no:</strong></td>
          <td><%= session("loan_serial_no") %></td>
        </tr>
        <tr>
          <td align="right"><strong>Qty:</strong></td>
          <td><%= session("loan_qty") %></td>
        </tr>
        <tr>
          <td align="right"><strong>LIC:</strong></td>
          <td>$<%= FormatNumber(session("loan_lic")) %></td>
        </tr>
        <tr>
          <td align="right"><strong>Loan date:</strong></td>
          <td><%= FormatDateTime(session("loan_date"),1) %></td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td align="right"><strong>First expiry date:</strong></td>
          <td><% if DateDiff("d",session("loan_first_expiry"), strTodayDate) > 0 then %>
            <font color="red">
            <% end if %>
            <%= FormatDateTime(session("loan_first_expiry"),1) %> </font></td>
        </tr>
        <tr bgcolor="#CCCCCC">
          <td align="right"><strong>Final expiry date:</strong></td>
          <td><% if DateDiff("d",session("loan_final_expiry"), strTodayDate) > 0 then %>
            <font color="red">
            <% end if %>
            <%= FormatDateTime(session("loan_final_expiry"),1) %> </font></td>
        </tr>
        <tr>
          <td align="right"><strong>Order no:</strong></td>
          <td><%= loanOrderNo %></td>
        </tr>
        <tr>
          <td align="right"><strong>Order line:</strong></td>
          <td><%= loanOrderLine %></td>
        </tr>
        <tr>
          <td align="right"><strong>Shipment no:</strong></td>
          <td><%= session("loan_shipment_no") %></td>
        </tr>
      </table>
      <p align="center"><%= strMessageText %></p></td>
    <td valign="top"><% if Trim(session("loanstock_info_record_count")) <> Trim(intQty) then %>
      <form action="" method="post" name="form_add_loan" id="form_add_loan" onsubmit="return validateAddForm(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="item_maintenance_header">Loan Stock Info</td>
          </tr>
          <tr>
            <td width="20%" align="right">Location<span class="mandatory">*</span>:<br /></td>
            <td width="80%"><input type="text" id="txtLocation" name="txtLocation" maxlength="20" size="25" value="" /></td>
          </tr>
          <tr>
            <td align="right">Serial no:</td>
            <td><input type="text" id="txtSerial" name="txtSerial" maxlength="20" size="25" value="" /></td>
          </tr>
          <tr>
            <td align="right">Comments:</td>
            <td><textarea name="txtComments" id="txtComments" cols="35" rows="4"></textarea></td>
          </tr>
          <tr>
            <td></td>
            <td><input type="hidden" name="Action" />
              <input type="submit" value="Update Loan Stock Info" /></td>
          </tr>
        </table>
      </form>
      <br />
      <% end if %>
      <table cellpadding="5" cellspacing="0" border="0" class="form_box_nowidth" width="700">
        <tr>
          <td class="item_maintenance_header">Location</td>
          <td class="item_maintenance_header">Serial #</td>
          <td class="item_maintenance_header">Comments</td>
          <td class="item_maintenance_header"></td>
          <td class="item_maintenance_header"></td>
          <td class="item_maintenance_header"></td>
        </tr>
        <%= strLoanStockInfoList %>
      </table>
      <p align="center"><small>Record count: <%= session("loanstock_info_record_count") %></small></p>
      <br />
      <% if session("loan_date_diff") < 180 then %>
      <form action="" method="post" name="form_renew_loan" id="form_renew_loan" onsubmit="return validateRenewForm(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="item_maintenance_header">3 Months Renewal - NO LIMIT</td>
          </tr>          
          <tr>
            <td width="20%" align="right" valign="top">Comments:</td>
            <td width="80%"><textarea name="txtRenewComments" id="txtRenewComments" cols="35" rows="4"></textarea></td>
          </tr>
          <tr>
            <td></td>
            <td><input type="hidden" name="Action" />
              <input type="submit" value="Renew" /></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <br />
      <table cellpadding="5" cellspacing="0" border="0" class="form_box_nowidth" width="500">
        <tr>
          <td class="item_maintenance_header">Submitted</td>
          <td class="item_maintenance_header">Expiry Date</td>
          <td class="item_maintenance_header">Comments</td>
          <td class="item_maintenance_header">GM Approval</td>
        </tr>
        <%= strLoanRenewalList %>
      </table></td>
  </tr>
</table>
</body>
</html>