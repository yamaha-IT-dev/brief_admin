<!--#include file="../include/connection_auction.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>#View Winner</title>
<link REL="stylesheet" HREF="include/stylesheet.css" TYPE="text/css">
<script type="text/javascript" src="../include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtSalesOrderNo,"Sales Order No");
	//reason += validateSpecialCharacters(theForm.txtSalesOrderNo,"Sales Order No");
	//reason += validateEmptyField(theForm.txtInvoiceNo,"Invoice No");
	//reason += validateSpecialCharacters(theForm.txtInvoiceNo,"Invoice No");
	reason += validateNumeric(theForm.txtDelivery,"Delivery fee");
	
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
</script>
<%

Function getUser(intID)
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tblQARegistration WHERE regID = " & intID

	rs.Open strSQL, conn
	
    if not DB_RecSetIsEmpty(rs) Then
		session("winner_username") 			= rs("regUsername")
		session("winner_name") 				= rs("regName")
		session("winner_email") 			= rs("regEmail")
		session("winner_sales_order_no") 	= rs("regSalesOrderNo")
		session("winner_invoice_no") 		= rs("regInvoiceNo")
		session("winner_delivery") 			= rs("regDelivery")
		session("winner_date_modified") 	= rs("regDateModified")
		session("winner_modified_by") 		= rs("regModifiedBy")
    end if

    call CloseDataBase()
end function

sub updateWinner
	dim strSQL
	dim intID
	intID = request("id")
	
	dim strSalesOrderNo
	dim strInvoiceNo	
	dim intDelivery
	
	strSalesOrderNo	= Replace(Request.Form("txtSalesOrderNo"),"'","''")
	strInvoiceNo	= Replace(Request.Form("txtInvoiceNo"),"'","''")	
	intDelivery 	= Request.Form("txtDelivery")
	
	Call OpenDataBase()

	strSQL = "UPDATE tblQARegistration SET "
	strSQL = strSQL & "regSalesOrderNo = '" & strSalesOrderNo & "',"
	strSQL = strSQL & "regInvoiceNo = '" & strInvoiceNo & "',"
	strSQL = strSQL & "regDelivery = CONVERT(money," & intDelivery & "),"	
	strSQL = strSQL & "regDateModified = getdate(),"
	strSQL = strSQL & "regModifiedBy = '" & session("logged_username") & "'"
	strSQL = strSQL & "	WHERE regID = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The winner has been updated."
	end if

	Call CloseDataBase()
end sub

sub listWinningItems
	dim iRecordCount
	iRecordCount = 0
	
	dim intTotal
	intTotal = 0
	
	dim intGrandTotal
	intGrandTotal = 0
	
    dim strSortBy
	dim strSortItem
    dim strSQL
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount	
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 200
	
	strSQL = "SELECT A.*, C.catname "
	strSQL = strSQL & "	FROM tblQAAuctions A "
	strSQL = strSQL & "		INNER JOIN tblQACategories C ON A.aucCategoryID = C.catID "
	'strSQL = strSQL & "		LEFT JOIN tblQARegistration R ON A.aucCurrentBidder = " & session("winner_id")& ""
	strSQL = strSQL & " WHERE aucCurrentBidder = " & session("winner_id")& ""
	strSQL = strSQL & " 		AND aucEnded = 'Y' AND aucFiresale <> '1' "
	strSQL = strSQL & "	ORDER BY aucLotNo"
			
	'response.write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
		For intRecord = 1 To rs.PageSize 			
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucLotNo") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("catName") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucItemTitle") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucBaseCode") & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucLIC") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">$" & rs("aucStartingBid") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">$" & rs("aucCurrentBid") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			intTotal = intTotal + rs("aucCurrentBid")
			rs.movenext
			iRecordCount = iRecordCount + 1
			
			If rs.EOF Then Exit For 
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""6"" align=""center"">No items found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""6"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "Purchased: " & intRecordCount & " item(s)"
	strDisplayList = strDisplayList & "<br>Total: $" & FormatNumber(intTotal) & ""
	strDisplayList = strDisplayList & "<br>Delivery Fee: $" & session("winner_delivery") & ""
	
	if isnull(session("winner_delivery")) then
		strDisplayList = strDisplayList & "<br><strong>Grant Total: <u>$" & FormatNumber(intTotal) & "</u></strong>"
	else 
		intGrandTotal = intTotal + session("winner_delivery")
		strDisplayList = strDisplayList & "<br><strong>Grand Total: <u>$" & FormatNumber(intGrandTotal) & "</u></strong>"
	end if
		
	'strDisplayList = strDisplayList & "<br><strong>Grand Total: <u>$" & intGrandTotal & "</u></strong>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main	
	session("winner_id") = trim(request("id"))
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		select case Trim(Request("Action"))
			case "Update"
				call updateWinner
		end select
	end if
	
	call getUser(session("winner_id"))
	call listWinningItems
end sub

call main

dim strMessageText
dim strDisplayList
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="800">
  <tr>
    <td class="first_content"><h2>Update Winner</h2>
      <p><img src="images/backward_arrow.gif" /> <a href="list_winners.asp">Back to Winner List</a></p>
      <table cellpadding="4" cellspacing="0" class="created_table">
        <tr>
          <td width="20%"><strong>Last modified by:</strong></td>
          <td width="30%"><%= session("winner_modified_by") %></td>
          <td width="50%"><%= session("winner_date_modified") %></td>
        </tr>
      </table>
      <br />
      <font color="red"><%= strMessageText %></font>
      <form action="" method="post" name="form_update_changeover" id="form_update_changeover" onsubmit="return validateFormOnSubmit(this)">
        <table cellpadding="5" cellspacing="0" class="item_maintenance_box" bgcolor="#FFFFFF">
          <tr>
            <td width="25%">Name:</td>
            <td width="75%"><a href="mailto:<%= session("winner_email") %>"><%= session("winner_name") %></a></td>
          </tr>
          <tr>
            <td>Sales order no:</td>
            <td><input type="text" id="txtSalesOrderNo" name="txtSalesOrderNo" maxlength="6" size="10" value="<%= Session("winner_sales_order_no") %>" /></td>
          </tr>
          <tr>
            <td>Invoice no:</td>
            <td><input type="text" id="txtInvoiceNo" name="txtInvoiceNo" maxlength="7" size="10" value="<%= Session("winner_invoice_no") %>" /></td>
          </tr>
          <tr>
            <td>Delivery fee:</td>
            <td>$ <input type="text" id="txtDelivery" name="txtDelivery" maxlength="4" size="6" value="<%= Session("winner_delivery") %>" /></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><input type="hidden" name="Action" />
              <input type="submit" value="Update" /></td>
          </tr>
        </table>
        <br />
        <table cellspacing="0" cellpadding="4" class="database_records" width="100%">
          <tr class="innerdoctitle" align="center">
            <td>Lot</td>
            <td>Category</td>
            <td>Product</td>
            <td>Component</td>
            <td>Reserve</td>
            <td>Winning Bid</td>
          </tr>
          <%= strDisplayList %>
         
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>