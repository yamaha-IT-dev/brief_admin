<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!--#include file="class/clsWarehouseReturn.asp " -->
<% strSection = "quarantine" %>
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
	
	reason += validateEmptyField(theForm.cboReturnType,"Type");
	reason += validateEmptyField(theForm.txtItemCode,"Item Code");
	reason += validateSpecialCharacters(theForm.txtItemCode,"Item Code");
	reason += validateEmptyField(theForm.txtDescription,"Item Description");
	reason += validateSpecialCharacters(theForm.txtDescription,"Item Description");
	reason += validateSpecialCharacters(theForm.txtDealer,"Dealer");
	reason += validateEmptyField(theForm.txtShipmentNo,"Shipment No");
	reason += validateSpecialCharacters(theForm.txtShipmentNo,"Shipment No");
	reason += validateEmptyField(theForm.txtQty,"Qty");
	reason += validateSpecialCharacters(theForm.txtQty,"Qty");

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

Sub getQuarantine
	dim strSQL
	dim intID
	intID = request("id")

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_quarantines WHERE quarantine_id = " & intID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("return_type") 		= rs("return_type")
		session("department") 		= rs("department")
		session("item_code") 		= rs("item_code")
		session("shipment_no") 		= rs("shipment_no")
		session("qty") 				= rs("qty")
		session("description") 		= rs("description")
		session("photos") 			= rs("photos")
		session("gra") 				= rs("gra")
		session("return_carrier") 	= rs("return_carrier")
		session("return_connote")	= rs("return_connote")
		session("original_connote") = rs("original_connote")
		session("dealer") 			= rs("dealer")
		session("reason_code") 		= rs("reason_code")
		session("instruction") 		= rs("instruction")
		session("serial_no") 		= rs("serial_no")
		session("stock_type") 		= rs("stock_type")
		session("date_received") 	= rs("date_received")
		session("comments") 		= rs("comments")
		session("date_created")		= rs("date_created")
		session("created_by") 		= rs("created_by")
		session("date_modified")	= rs("date_modified")
		session("modified_by") 		= rs("modified_by")
		session("status") 			= rs("status")
    end if

	session("days_quarantine") = DateDiff("d",rs("date_created"), strTodayDate)

    call CloseDataBase()

end sub

sub updateReturn
	dim strSQL
	dim intID
	intID = request("id")
	
	dim intReturnType
	dim strDepartment
	dim strItemCode
	dim strShipmentNo
	dim intQty
	dim strDescription
	dim intPhotos
	dim strGRA
	dim strReturnCarrier
	dim strReturnConnote
	dim strOriginalConnote
	dim strDealer
	dim intReasonCode
	dim intInstruction
	dim strSerialNo
	dim intStockType
	dim strDateReceived
	dim strComments
	dim intStatus
	
	intReturnType 		= Request.Form("cboReturnType")
	strDepartment 		= Request.Form("cboDepartment")
	strItemCode 		= Request.Form("txtItemCode")
	strShipmentNo 		= Request.Form("txtShipmentNo")
	intQty 				= Request.Form("txtQty")
	strDescription 		= Request.Form("txtDescription")
	intPhotos 			= Request.Form("cboPhotos")
	strGRA 				= Request.Form("txtGRA")
	strReturnCarrier 	= Request.Form("cboReturnCarrier")
	strReturnConnote 	= Request.Form("txtReturnConnote")
	strOriginalConnote 	= Request.Form("txtOriginalConnote")
	strDealer 			= Request.Form("txtDealer")
	intReasonCode 		= Request.Form("cboReasonCode")
	intInstruction 		= Request.Form("cboInstruction")
	strSerialNo 		= Request.Form("txtSerialNo")
	intStockType		= Request.Form("cboStockType")
	strDateReceived		= Request.Form("txtDateReceived")
	strComments 		= Request.Form("txtComments")
	intStatus 			= Request.Form("cboStatus")

	call OpenDataBase()

	strSQL = "UPDATE yma_quarantines SET "
	strSQL = strSQL & "return_type = '" & intReturnType & "',"
	strSQL = strSQL & "department = '" & strDepartment & "',"
	strSQL = strSQL & "item_code = '" & Server.HTMLEncode(strItemCode) & "',"
	strSQL = strSQL & "shipment_no = '" & Server.HTMLEncode(strShipmentNo) & "',"
	strSQL = strSQL & "qty = '" & intQty & "',"
	strSQL = strSQL & "description = '" & Server.HTMLEncode(strDescription) & "',"
	strSQL = strSQL & "photos = '" & intPhotos & "',"
	strSQL = strSQL & "gra = '" & Server.HTMLEncode(strGRA) & "',"
	strSQL = strSQL & "return_carrier = '" & strReturnCarrier & "',"
	strSQL = strSQL & "return_connote = '" & Server.HTMLEncode(strReturnConnote) & "',"
	strSQL = strSQL & "original_connote = '" & Server.HTMLEncode(strOriginalConnote) & "',"
	strSQL = strSQL & "dealer = '" & Server.HTMLEncode(strDealer) & "',"
	strSQL = strSQL & "reason_code = '" & intReasonCode & "',"
	strSQL = strSQL & "instruction = '" & intInstruction & "',"
	strSQL = strSQL & "serial_no = '" & Server.HTMLEncode(strSerialNo) & "',"
	strSQL = strSQL & "date_received = CONVERT(datetime,'" & strDateReceived & "',103),"
	strSQL = strSQL & "stock_type = '" & intStockType & "',"
	strSQL = strSQL & "comments = '" & Server.HTMLEncode(strComments) & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & session("UsrUserName") & "',"
	strSQL = strSQL & "status = '" & intStatus & "' WHERE quarantine_id = " & intID

	'response.Write strSQL

	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if

	call CloseDataBase()
end sub

sub main
	call getReasonCode
	
	dim intID
	intID 	= request("id")
	
	call getQuarantine
	call listComments(intID,warehouseReturnModuleID)
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateReturn
				call getQuarantine
				call getReasonCode
			case "Comment"
				call addComment(intID,warehouseReturnModuleID)
				call listComments(intID,warehouseReturnModuleID)
		end select
	end if
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
          <td><a href="list_damage.asp"><img src="images/icon_gra.jpg" border="0" alt="Warehouse Return" /></a></td>
          <td valign="top"><img src="images/backward_arrow.gif" border="0" /> <a href="list_return.asp">Back to List</a>
            <h2>Update Warehouse Return</h2>
            <font color="red"><%= strMessageText %></font></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("created_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("modified_by") %></td>
                <td><%= displayDateFormatted(session("date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <form action="" method="post" name="form_update_quarantine" id="form_update_quarantine" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="3" align="center" class="item_maintenance_header"><%= session("days_quarantine") %> days in Warehouse</td>
          </tr>
          <tr>
            <td colspan="3">Return type:
              <% 	select case session("return_type")
						case "1"
							response.write "Managed (GRA)"
						case "0"
							response.write "Un-managed"
						case "2"
							response.write "Un-addressed"	
					end select
			  %></td>
          </tr>
          <tr>
            <td>Department<span class="mandatory">*</span>:<br />
              <select name="cboDepartment">
                <option <% if session("department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                <option <% if session("department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                <option <% if session("department") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
              </select></td>
            <td colspan="2">Item code<span class="mandatory">*</span>:<br />
              <input type="text" id="txtItemCode" name="txtItemCode" maxlength="20" size="30" value="<%= Server.HTMLEncode(session("item_code")) %>" /></td>
          </tr>
          <tr>
            <td colspan="3">Item description<span class="mandatory">*</span>:<br />
              <input type="text" id="txtDescription" name="txtDescription" maxlength="50" size="60" value="<%= session("description") %>" /></td>
          </tr>
          <tr>
            <td>Return con-note<span class="mandatory">*</span>:<br />
              <input type="text" id="txtReturnConnote" name="txtReturnConnote" maxlength="15" size="20" value="<%= session("return_connote") %>" /></td>
            <td colspan="2">Original con-note:<br />
              <input type="text" id="txtOriginalConnote" name="txtOriginalConnote" maxlength="15" size="20" value="<%= session("original_connote") %>" /></td>
          </tr>
          <tr>
            <td width="50%">Dealer:<br />
              <input type="text" id="txtDealer" name="txtDealer" maxlength="30" size="35" value="<%= session("dealer") %>" /></td>
            <td width="30%">Shipment no<span class="mandatory">*</span>:<br />
              <input type="text" id="txtShipmentNo" name="txtShipmentNo" maxlength="10" size="12" value="<%= session("shipment_no") %>" /></td>
            <td width="20%">Qty<span class="mandatory">*</span>:<br />
              <input type="text" id="txtQty" name="txtQty" maxlength="4" size="5" value="<%= session("qty") %>" /></td>
          </tr>
          <tr>
            <td colspan="3"><input type="checkbox" name="chkPhotos" id="chkPhotos" value="1" <% if session("photos") = "1" then Response.Write " checked" end if%> />
              <img src="images/camera_icon.gif" alt="Photo" border="0" /> <a href="file:\\YAMMAS22\quarantine\<%= intID %>" target="_blank">Directory</a> <small>(Rename main photo to <u>1.jpg</u>)</small></td>
          </tr>
          <tr rel="unmanaged">
            <td colspan="3" bgcolor="#66CCFF"><table width="100%" border="0" cellspacing="0" cellpadding="4">
                <tr>
                  <td>Return carrier:<br />
                    <select name="cboReturnCarrier">
                      <option <% if session("return_carrier") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("return_carrier") = "Cope" then Response.Write " selected" end if%> value="Cope">Cope</option>
                      <option <% if session("return_carrier") = "StarTrack" then Response.Write " selected" end if%> value="StarTrack">StarTrack</option>
                      <option <% if session("return_carrier") = "Schenker" then Response.Write " selected" end if%> value="Schenker">Schenker</option>
                      <option <% if session("return_carrier") = "Kings" then Response.Write " selected" end if%> value="Kings">Kings</option>
                    </select></td>
                  <td>Serial no(s): (If damaged)<br />
                    <input type="text" id="txtSerialNo" name="txtSerialNo" maxlength="50" size="55" value="<%= session("serial_no") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2" align="right">Stock: <img src="images/icon_new.gif" border="0" align="top" />
                    <select name="cboStockType">
                      <option <% if session("stock_type") = "" then Response.Write " selected" end if%> value="">...</option>
                      <option <% if session("stock_type") = "1" then Response.Write " selected" end if%> value="1">Damaged</option>
                      <option <% if session("stock_type") = "2" then Response.Write " selected" end if%> value="2">Partial</option>
                    </select></td>
                </tr>
                <tr>
                  <td colspan="2">Date stock received:<br />
                    <input type="text" id="txtDateReceived" name="txtDateReceived" maxlength="10" size="15" value="<%= session("date_received") %>" />
                    <span class="mandatory"><em>DD/MM/YYYY</em></span></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td>Reason: <img src="images/icon_new.gif" border="0" /><br />
              <select name="cboReasonCode">                
                <%= strReasonCodeList %>
              </select></td>
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td>Instruction:<br />
              <select name="cboInstruction">
                <option <% if session("instruction") = "" then Response.Write " selected" end if%> value="">...</option>
                <option <% if session("instruction") = "1" then Response.Write " selected" end if%> value="1">Return to good stock 3T</option>
                <option <% if session("instruction") = "2" then Response.Write " selected" end if%> value="2">Transfer to Excel 3XL</option>
                <option <% if session("instruction") = "3" then Response.Write " selected" end if%> value="3">Resend to customer</option>
                <option <% if session("instruction") = "4" then Response.Write " selected" end if%> value="4">Damaged item to Excel - good stock to 3T</option>
              </select></td>
            <td colspan="2">GRA:<br />
              <input type="text" id="txtGRA" name="txtGRA" maxlength="7" size="10" value="<%= session("gra") %>" /></td>
          </tr>
          <tr>
            <td colspan="3">Comments:<br />
              <textarea name="txtComments" id="txtComments" cols="50" rows="4"><%= session("comments") %></textarea></td>
          </tr>
          <tr class="status_row">
            <td colspan="3">Status:
              <select name="cboStatus">
                <option <% if session("status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                <option <% if session("status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
              </select></td>
          </tr>
          <tr>
            <td colspan="3"><input type="hidden" name="Action" />
          <input type="submit" value="Update Warehouse Return" /></td>
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