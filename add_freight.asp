<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsFreight.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>New MPD Freight Request</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="javascript" type="text/javascript">
function pickupHO(form) {
	if (form.chkPickupHO.checked) {
		form.txtPickupName.value 		= "Yamaha Music Australia"
		form.txtPickupAddress.value 	= "Level 1 / 99 Queensbridge St"
		form.txtPickupContact.value 	= "Matthew Madden"
		form.txtPickupPhone.value 		= "03 9693 5206"
		form.txtPickupCity.value 		= "Southbank"
		form.txtPickupPostcode.value 	= "3006"
	} else {
		form.txtPickupName.value 		= ""
		form.txtPickupAddress.value 	= ""
		form.txtPickupContact.value 	= ""
		form.txtPickupPhone.value 		= ""
		form.txtPickupCity.value 		= ""
		form.txtPickupPostcode.value 	= ""
	}
}

function receiverHO(form) {
	if (form.chkReceiverHO.checked) {	
		form.txtReceiverName.value 		= "Yamaha Music Australia"
		form.txtReceiverAddress.value 	= "Level 1 / 99 Queensbridge St"
		form.txtReceiverContact.value 	= "Matthew Madden"
		form.txtReceiverPhone.value 	= "03 9693 5206"
		form.txtReceiverCity.value 		= "Southbank"
		form.txtReceiverPostcode.value 	= "3006"		
	} else {	
		form.txtReceiverName.value 		= ""
		form.txtReceiverAddress.value 	= ""
		form.txtReceiverContact.value 	= ""
		form.txtReceiverPhone.value 	= ""
		form.txtReceiverCity.value 		= ""
		form.txtReceiverPostcode.value 	= ""
	}
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;	
	
	reason += validateEmptyField(theForm.txtPickupName,"Pickup Name");
	reason += validateSpecialCharacters(theForm.txtPickupName,"Pickup Name");

	//reason += validateEmptyField(theForm.txtPickupContact,"Pickup Contact");
	reason += validateSpecialCharacters(theForm.txtPickupContact,"Pickup Contact");
	
	reason += validateEmptyField(theForm.txtPickupPhone,"Pickup Phone");
	reason += validateSpecialCharacters(theForm.txtPickupPhone,"Pickup Phone");
	
	reason += validateEmptyField(theForm.txtPickupAddress,"Pickup Address");
	reason += validateSpecialCharacters(theForm.txtPickupAddress,"Pickup Address");
	
	reason += validateEmptyField(theForm.txtPickupCity,"Pickup City");
	reason += validateSpecialCharacters(theForm.txtPickupCity,"Pickup City");
	
	reason += validateNumeric(theForm.txtPickupPostcode,"Pickup Postcode");
	
	reason += validateEmptyField(theForm.txtReceiverName,"Receiver Name");
	reason += validateSpecialCharacters(theForm.txtReceiverName,"Receiver Name");

	//reason += validateEmptyField(theForm.txtReceiverContact,"Receiver Contact");
	reason += validateSpecialCharacters(theForm.txtReceiverContact,"Receiver Contact");
	
	reason += validateEmptyField(theForm.txtReceiverPhone,"Receiver Phone");
	reason += validateSpecialCharacters(theForm.txtReceiverPhone,"Receiver Phone");
	
	reason += validateEmptyField(theForm.txtReceiverAddress,"Receiver Address");
	reason += validateSpecialCharacters(theForm.txtReceiverAddress,"Receiver Address");
	
	reason += validateEmptyField(theForm.txtReceiverCity,"Receiver City");
	reason += validateSpecialCharacters(theForm.txtReceiverCity,"Receiver City");
	
	reason += validateNumeric(theForm.txtReceiverPostcode,"Receiver Postcode");
	
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
				call setFreightSessionVariables
				
				strPickupName		= Replace(Trim(Request.Form("txtPickupName")),"'","''")
				strPickupContact	= Replace(Trim(Request.Form("txtPickupContact")),"'","''")
				strPickupPhone		= Replace(Trim(Request.Form("txtPickupPhone")),"'","''")
				strPickupAddress	= Replace(Trim(Request.Form("txtPickupAddress")),"'","''")
				strPickupCity		= Replace(Trim(Request.Form("txtPickupCity")),"'","''")
				strPickupState		= Trim(Request.Form("cboPickupState"))
				intPickupPostcode	= Trim(Request.Form("txtPickupPostcode"))
				strPickupComments	= Replace(Trim(Request.Form("txtPickupComments")),"'","''")
				
				strReceiverName		= Replace(Trim(Request.Form("txtReceiverName")),"'","''")
				strReceiverContact	= Replace(Trim(Request.Form("txtReceiverContact")),"'","''")
				strReceiverPhone	= Replace(Trim(Request.Form("txtReceiverPhone")),"'","''")
				strReceiverAddress	= Replace(Trim(Request.Form("txtReceiverAddress")),"'","''")
				strReceiverCity		= Replace(Trim(Request.Form("txtReceiverCity")),"'","''")
				strReceiverState	= Trim(Request.Form("cboReceiverState"))
				intReceiverPostcode	= Trim(Request.Form("txtReceiverPostcode"))
				strReceiverComments	= Replace(Trim(Request.Form("txtReceiverComments")),"'","''")
				
				call addFreight(strPickupName, strPickupContact, strPickupPhone, strPickupAddress, strPickupCity, strPickupState, intPickupPostcode, strPickupComments, strReceiverName, strReceiverContact, strReceiverPhone, strReceiverAddress, strReceiverCity, strReceiverState, intReceiverPostcode, strReceiverComments, session("logged_username"))
		end select
	end if
end sub

call main

dim strMessageText
dim strPickupName, strPickupContact, strPickupPhone, strPickupAddress, strPickupCity, strPickupState, intPickupPostcode, strPickupComments, strReceiverName, strReceiverContact, strReceiverPhone, strReceiverAddress, strReceiverCity, strReceiverState, intReceiverPostcode, strReceiverComments
%>
</head>
<body>
<div class="main"> 
  <!-- #include file="include/header.asp" --> 
  <img src="images/backward_arrow.gif" border="0" /> <a href="list_freight.asp">Back to List</a>
  <h2>New Freight Request</h2>
  <p><font color="red"><%= strMessageText %></font></p>
  <form action="" method="post" name="form_add_freight" id="form_add_freight" onsubmit="return validateFormOnSubmit(this)">
    <table>
      <tr>
        <td><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
            <tr>
              <td colspan="2" bgcolor="#f0f0f0"><strong>Pickup Details</strong></td>
            </tr>
            <tr>
              <td colspan="2"><input type="checkbox" name="chkPickupHO" id="chkPickupHO" onclick="pickupHO(this.form)" />
                <em>Tick here to select YMA Head Office as the Pickup</em></td>
            </tr>
            <tr>
              <td width="30%">Name<span class="mandatory">*</span>:</td>
              <td width="70%"><input type="text" id="txtPickupName" name="txtPickupName" maxlength="40" size="45" /></td>
            </tr>
            <tr>
              <td>Contact Person:</td>
              <td><input type="text" id="txtPickupContact" name="txtPickupContact" maxlength="30" size="40" /></td>
            </tr>
            <tr>
              <td>Phone<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtPickupPhone" name="txtPickupPhone" maxlength="30" size="30" /></td>
            </tr>
            <tr>
              <td>Address<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtPickupAddress" name="txtPickupAddress" maxlength="40" size="45" /></td>
            </tr>
            <tr>
              <td>City<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtPickupCity" name="txtPickupCity" maxlength="30" size="35" /></td>
            </tr>
            <tr>
              <td>State:</td>
              <td><select name="cboPickupState">
                  <option value="VIC">VIC</option>
                  <option value="NSW">NSW</option>
                  <option value="ACT">ACT</option>
                  <option value="QLD">QLD</option>
                  <option value="NT">NT</option>
                  <option value="WA">WA</option>
                  <option value="SA">SA</option>
                  <option value="TAS">TAS</option>
                  <option value="Other">Other</option>
                </select></td>
            </tr>
            <tr>
              <td>Postcode<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtPickupPostcode" name="txtPickupPostcode" maxlength="4" size="8" /></td>
            </tr>
            <tr>
              <td valign="top">Comments:</td>
              <td><textarea name="txtPickupComments" id="txtPickupComments" cols="35" rows="3" onkeydown="limitText(this.form.txtPickupComments,this.form.countdown,120);" 
onkeyup="limitText(this.form.txtPickupComments,this.form.countdown,120);"></textarea></td>
            </tr>
          </table></td>
        <td><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
            <tr>
              <td colspan="2" bgcolor="#f0f0f0"><strong>Receiver Details</strong></td>
            </tr>
            <tr>
              <td colspan="2"><input type="checkbox" name="chkReceiverHO" id="chkReceiverHO" onclick="receiverHO(this.form)" />
                <em>Tick here to select YMA Head Office as the Receiver</em></td>
            </tr>
            <tr>
              <td width="30%">Name<span class="mandatory">*</span>:</td>
              <td width="70%"><input type="text" id="txtReceiverName" name="txtReceiverName" maxlength="40" size="45" /></td>
            </tr>
            <tr>
              <td>Contact Person:</td>
              <td><input type="text" id="txtReceiverContact" name="txtReceiverContact" maxlength="30" size="40" /></td>
            </tr>
            <tr>
              <td>Phone<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtReceiverPhone" name="txtReceiverPhone" maxlength="30" size="30" /></td>
            </tr>
            <tr>
              <td>Address<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtReceiverAddress" name="txtReceiverAddress" maxlength="40" size="45" /></td>
            </tr>
            <tr>
              <td>City<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtReceiverCity" name="txtReceiverCity" maxlength="30" size="35" /></td>
            </tr>
            <tr>
              <td>State:</td>
              <td><select name="cboReceiverState">
                  <option value="VIC">VIC</option>
                  <option value="NSW">NSW</option>
                  <option value="ACT">ACT</option>
                  <option value="QLD">QLD</option>
                  <option value="NT">NT</option>
                  <option value="WA">WA</option>
                  <option value="SA">SA</option>
                  <option value="TAS">TAS</option>
                  <option value="Other">Other</option>
                </select></td>
            </tr>
            <tr>
              <td>Postcode<span class="mandatory">*</span>:</td>
              <td><input type="text" id="txtReceiverPostcode" name="txtReceiverPostcode" maxlength="4" size="8" /></td>
            </tr>
            <tr>
              <td valign="top">Comments:</td>
              <td><textarea name="txtReceiverComments" id="txtReceiverComments" cols="35" rows="3" onkeydown="limitText(this.form.txtReceiverComments,this.form.countdown,120);" 

onkeyup="limitText(this.form.txtReceiverComments,this.form.countdown,120);"></textarea></td>
            </tr>
          </table></td>
      </tr>
    </table>
    <p>
      <input type="hidden" name="Action" />
      <input type="submit" name="confirm_button" value="Confirm" />
      <input type="reset" name="reset" id="reset" value="Reset" />
    </p>
  </form>
</div>
</body>
</html>