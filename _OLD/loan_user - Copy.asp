<%
'setup for Australian Date/Time
session.lcid = 2057
session.timeout = 420

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="include/loan_functions.asp " -->
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsAuction.asp" -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!--#include file="class/clsLoanAccount.asp " -->
<!--#include file="class/clsLoanBase.asp " -->
<!--#include file="class/clsLoanBasket.asp " -->
<!--#include file="class/clsLoanLocation.asp " -->
<!--#include file="class/clsLoanRenewal.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Cache-control" content="no-store">
<title>Loan Stock User</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link rel="stylesheet" href="css/sticky-navigation.css" />
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"></script>
<script type="text/javascript" src="../include/generic_form_validations.js"></script>
<script>
$(function() {
	var sticky_navigation_offset_top = $('#sticky_navigation').offset().top;
	
	var sticky_navigation = function(){
		var scroll_top = $(window).scrollTop();

		if (scroll_top > sticky_navigation_offset_top) { 
			$('#sticky_navigation').css({ 'position': 'fixed', 'top':0, 'left':0 });
		} else {
			$('#sticky_navigation').css({ 'position': 'relative' }); 
		}   
	};
	
	sticky_navigation();
	
	$(window).scroll(function() {
		 sticky_navigation();
	});
	
	$('a[href="#"]').click(function(event){ 
		event.preventDefault(); 
	});
	
});

function searchStock(){    
    var strSearch	= document.forms[0].txtSearch.value;
	var strSort	  	= document.forms[0].cboSort.value;
    document.location.href = 'loan_user.asp?type=search&txtSearch=' + strSearch + '&sort=' + strSort;
}
    
function resetSearch(){
	document.location.href = 'loan_user.asp?type=reset';    
}  

function submitRenewalForm(theForm) {
    theForm.Action.value = 'Renew';

	return true;
}

function validateAddLocation(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtLocation,"Location");
	reason += validateSpecialCharacters(theForm.txtLocation,"Location");		

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

function validateUpdateLocation(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtUpdateLocation,"Location");
	reason += validateSpecialCharacters(theForm.txtUpdateLocation,"Location");		

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
</head>
<body>
<%
sub setSearch	
	Select case trim(request("type"))
		case "reset"
			session("loan_user_search") 		= ""		
			session("loan_user_sort") 			= ""
			session("loan_user_initial_page") 	= 1
		case "search"
			session("loan_user_search") 		= trim(request("txtSearch"))
			session("loan_user_sort") 			= trim(request("sort"))
			session("loan_user_initial_page") 	= 1
	end Select
end sub

sub displayLoanStock
	dim iRecordCount
	iRecordCount = 0
    dim intDays
	dim intNewDays
	dim intMonths
    dim strSQL
	
	dim intTotalLIC
	dim intTotalQty
	
	intTotalLIC = 0
	intTotalQty = 0
	
	dim intRecordCount
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 5000
	
	if session("loan_user_sort") = "" then
		session("loan_user_sort") = "oldest"
	end if
	
	strSQL = strSQL & "SELECT DISTINCT order_no, order_line, department, account_code, account_code_ext, account_name, item_code, serial_no,  "
	strSQL = strSQL & "		B9AHEN, B9SKSU, lic, qty, loan_year, loan_month, loan_day, loan_date, stockLocation, stockRenewalCounter, stockID, renStatus, renActive, renExpiryDate, renDateCreated "
	strSQL = strSQL & ", product_code, aucItemCode "
	strSQL = strSQL & " FROM OPENQUERY "
	strSQL = strSQL & " (AS400, 'SELECT B9JUNO AS order_no, B9JUGY AS order_line, Y1REGN AS department, B9URKC AS account_code, B9JURC AS account_code_ext, "
	strSQL = strSQL & "			Y1KOM1 AS account_name, B9SOSC AS item_code, B9SIBN AS serial_no, B9SKJY AS loan_year, B9SKJM AS loan_month, B9SKJD AS loan_day, "
	strSQL = strSQL & "			B9SKSU - B9AHEN AS qty, B9AHEN, B9SKSU, "
	strSQL = strSQL & " 		RIGHT(''0'' || B9SKJD,2)|| ''/'' || RIGHT (''0'' || B9SKJM,2) || ''/'' || B9SKJY AS loan_date, "
 	strSQL = strSQL & " 		CASE WHEN B9STJN <> ''00'' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) "
	strSQL = strSQL & " 			ELSE 0 "
	strSQL = strSQL & " 		END AS lic "	
	strSQL = strSQL & " 	FROM BF9EP "
	strSQL = strSQL & "			INNER JOIN EF2SP ON B9SOSC = E2SOSC "
	strSQL = strSQL & " 		INNER JOIN YF1MP ON CONCAT(B9URKC,B9JURC) = Y1KOKC "
	strSQL = strSQL & " WHERE Y1SKKI <> ''D'' AND E2NGTY = "
	strSQL = strSQL & "	(SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)"
	strSQL = strSQL & " 	AND E2NGTM = "
	strSQL = strSQL & "	(SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)"
	strSQL = strSQL & "	')"
	strSQL = strSQL & "		LEFT JOIN tbl_loan_location ON stockOrderNo = order_no AND stockOrderLine = order_line "
	strSQL = strSQL & "		LEFT JOIN (SELECT renOrderNo, renOrderLine, renStatus, renActive, renDateCreated, renExpiryDate FROM tbl_loan_renewal GROUP BY renOrderNo, renOrderLine, renStatus, renActive, renDateCreated, renExpiryDate) AS RENEWAL ON order_no = renOrderNo AND order_line = renOrderLine "
	strSQL = strSQL & "		LEFT JOIN yamaha_workflow..workflow_loan_return_item_list ON yamaha_workflow..workflow_loan_return_item_list.order_number = order_no AND yamaha_workflow..workflow_loan_return_item_list.order_lines = order_line "
	strSQL = strSQL & "		LEFT JOIN tbl_auction ON aucOrderNo = order_no AND aucOrderLine = order_line "
	strSQL = strSQL & " WHERE B9AHEN < B9SKSU and (renActive = 0 or renOrderNo is null) "
	strSQL = strSQL & "		AND (item_code LIKE '%" & UCASE(trim(session("loan_user_search"))) & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & UCASE(trim(session("loan_user_search"))) & "%')"
	strSQL = strSQL & "		AND account_code = '" & UCASE(session("loan_user_account")) & "' "
	strSQL = strSQL & " ORDER BY "
	select case session("loan_user_sort")
		case "oldest"
			strSQL = strSQL & "	loan_year ASC, loan_month ASC, loan_day ASC"
		case "latest"
			strSQL = strSQL & "	loan_year DESC, loan_month DESC, loan_day DESC"
		case "product"
			strSQL = strSQL & "	item_code"
		case "expensive"
			strSQL = strSQL & "	lic DESC"
		case "cheapest"
			strSQL = strSQL & "	lic"
		case "serial"
			strSQL = strSQL & "	serial_no"
	end select
	
	'response.write strSQL
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
	    rs.AbsolutePage = session("loan_user_initial_page")
	
		For intRecord = 1 To rs.PageSize
			if rs("renStatus") = 0 then
				intDays = DateDiff("d",rs("renDateCreated"), strTodayDate)
			else
				intDays = DateDiff("d",rs("loan_date"), strTodayDate)
			end if
			
			session("account_code_ext") = rs("account_code_ext")
			
			intTotalLIC = intTotalLIC + Cint(rs("lic"))
			intTotalQty = intTotalQty + Cint(rs("qty"))
			
			'dim strOriginalExpiryDate
			'strOriginalExpiryDate = DateAdd("m", 3, rs("loan_date"))
			
			if intDays > 90 then
				strDisplayList = strDisplayList & "<tr class=""overdue_row"">"
			else	
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			end if
			
			strDisplayList = strDisplayList & "<td align=""center"">"
			if IsNull(rs("product_code")) then
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_basket"" id=""form_basket"" onsubmit=""return submitBasketForm(this)"">"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""action"" value=""Basket"">"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""item_code"" value=""" & Trim(rs("item_code")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""serial_no"" value=""" & Trim(rs("serial_no")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""lic"" value=""" & Trim(rs("lic")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""location"" value=""" & Trim(rs("stockLocation")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""account_code"" value=""" & Trim(rs("account_code")) & "" & Trim(rs("account_code_ext")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_no"" value=""" & Trim(rs("order_no")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_line"" value=""" & Trim(rs("order_line")) & """>"
				strDisplayList = strDisplayList & "		<input type=""submit"" value=""Add to workflow"" />"
				strDisplayList = strDisplayList & "	</form>"
			else
				strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"					
			end if
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center"">"
			if IsNull(rs("aucItemCode")) then
			'if IsNull(rs("product_code")) then
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_auction"" id=""form_auction"" onsubmit=""return submitAuctionForm(this)"">"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""action"" value=""Auction"">"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""item_code"" value=""" & Trim(rs("item_code")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""serial_no"" value=""" & Trim(rs("serial_no")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""lic"" value=""" & Trim(rs("lic")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""location"" value=""" & Trim(rs("stockLocation")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""account_code"" value=""" & Trim(rs("account_code")) & "" & Trim(rs("account_code_ext")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""account_name"" value=""" & Trim(rs("account_name")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_no"" value=""" & Trim(rs("order_no")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_line"" value=""" & Trim(rs("order_line")) & """>"
				strDisplayList = strDisplayList & "		<input type=""submit"" value=""Add to auction"" />"
				strDisplayList = strDisplayList & "	</form>"
			else
				strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"					
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("item_code") & "</td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("serial_no") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("qty") & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"">" & FormatNumber(rs("lic")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"" nowrap>" & FormatDateTime(rs("loan_date"),1) & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>"
			if intDays > 90 then
				strDisplayList = strDisplayList & " <span style=""color:red"">"
			end if
			strDisplayList = strDisplayList & intDays & "</span>"
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""left"">"
			if rs("stockRenewalCounter") >= 1 then
				strDisplayList = strDisplayList & rs("stockRenewalCounter")
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if intDays > 85 and not IsNull(rs("stockLocation")) and IsNull(rs("product_code")) then
			'if intDays > 85 and not IsNull(rs("stockLocation")) then
				'if not isNull(rs("renStatus")) and rs("renStatus") <> 2 then
					strDisplayList = strDisplayList & "	<form method=""post"" name=""form_renew"" id=""form_renew"" onsubmit=""return submitRenewalForm(this)"">"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""action"" value=""Renew"">"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_no"" value=""" & Trim(rs("order_no")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_line"" value=""" & Trim(rs("order_line")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""account_code"" value=""" & Trim(rs("account_code")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""account_name"" value=""" & Trim(rs("account_name")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""department"" value=""" & Trim(rs("department")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""item_code"" value=""" & Trim(rs("item_code")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""location"" value=""" & Trim(rs("stockLocation")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""serial_no"" value=""" & Trim(rs("serial_no")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""lic"" value=""" & Trim(rs("lic")) & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""loan_date"" value=""" & Trim(rs("loan_date")) & """>"
					strDisplayList = strDisplayList & "		<input type=""submit"" value=""Renew"" />"
					strDisplayList = strDisplayList & "	</form>"
				'end if
			else
				strDisplayList = strDisplayList & "-"
			end if
			
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>"
			if not IsNull(rs("renDateCreated")) then
				strDisplayList = strDisplayList & FormatDateTime(rs("renDateCreated"),1)
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>"
			Select Case	rs("renStatus")
				case 1
					strDisplayList = strDisplayList & "<font color=""blue""><em>Pending Approval</em>"
				case 2
					strDisplayList = strDisplayList & "<font color=""red""><img src=""images/cross.gif""> Rejected"
				case 0
					strDisplayList = strDisplayList & "<font color=""green""><img src=""images/tick.gif""> Approved"
			end select
			strDisplayList = strDisplayList & "</font></td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if not IsNull(rs("renExpiryDate")) then
				strDisplayList = strDisplayList & FormatDateTime(rs("renExpiryDate"),1)
			end if
			strDisplayList = strDisplayList & "</td>"	
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>"
			if IsNull(rs("stockLocation")) then
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_add_location"" id=""form_add_location"" onsubmit=""return validateAddLocation(this)"">"
				strDisplayList = strDisplayList & "	<table>"
				strDisplayList = strDisplayList & "		<tr>"
				strDisplayList = strDisplayList & "			<td align=""center"">"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""action"" value=""Add"">"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""order_no"" value=""" & Trim(rs("order_no")) & """>"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""order_line"" value=""" & Trim(rs("order_line")) & """>"
				strDisplayList = strDisplayList & "				<input type=""text"" id=""txtLocation"" name=""txtLocation"" maxlength=""100"" size=""18"">"
				strDisplayList = strDisplayList & "			</td>"
				strDisplayList = strDisplayList & "			<td align=""center""><input type=""submit"" value=""Save"" /></td>"
				strDisplayList = strDisplayList & "		</tr>"
				strDisplayList = strDisplayList & "	</table>"
				strDisplayList = strDisplayList & "	</form>"
			else
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_update_location"" id=""form_update_location"" onsubmit=""return validateAddLocation(this)"">"
				strDisplayList = strDisplayList & "	<table>"
				strDisplayList = strDisplayList & "		<tr>"
				strDisplayList = strDisplayList & "			<td align=""center"">"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""action"" value=""Update"">"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""stock_id"" value=""" & Trim(rs("stockID")) & """>"
				strDisplayList = strDisplayList & "				<input type=""text"" id=""txtLocation"" name=""txtLocation"" maxlength=""100"" size=""18"" value=""" & rs("stockLocation") & """ >"
				strDisplayList = strDisplayList & "			</td>"
				strDisplayList = strDisplayList & "			<td align=""center""><input type=""submit"" value=""Save"" /></td>"
				strDisplayList = strDisplayList & "		</tr>"
				strDisplayList = strDisplayList & "	</table>"
				strDisplayList = strDisplayList & "	</form>"
			end if
			strDisplayList = strDisplayList & "</td>"		
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""14"" class=""bottom_grid"">No stocks found.</td></tr>"
	end if
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""14"" class=""bottom_grid"">"
	strDisplayList = strDisplayList & "<p>Total Value: $" & FormatNumber(intTotalLIC) & ""
	strDisplayList = strDisplayList & "<br>Total Loan Items: " & intTotalQty & "</p>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		strOrderNo 		= Request("order_no")
		strOrderLine	= Request("order_line")
		strAccountCode	= Request("account_code")
		strAccountName	= Request("account_name")
		strDepartment	= Request("department")
		strItemCode		= Request("item_code")
		strSerialNo		= Request("serial_no")
		strLIC			= FormatNumber(Request("lic"))
		strLoanDate		= Request("loan_date")
		stockLocation 	= Replace(Request("location"),"'","''")
		strLocation 	= Replace(Request("txtLocation"),"'","''")
		intStockID		= Request("stock_id")		
		
		select case Trim(Request("Action"))
			case "Renew"
				call addRenewal(strOrderNo, strOrderLine, strAccountCode, strAccountName, strDepartment, strItemCode, strSerialNo, stockLocation, strLIC, strLoanDate, session("logged_username"), session("emp_email"))
			case "Basket"
				call addItemBasket(strItemCode, strSerialNo, strLIC, strAccountCode, strOrderNo, strOrderLine)
			case "Auction"
				call addAuction(strItemCode, strSerialNo, strLIC, strAccountCode, strAccountName, strOrderNo, strOrderLine, session("logged_username"))
			case "Add"
				call addLoanLocation(strOrderNo, strOrderLine, strLocation, session("logged_username"))
			case "Update"
				call updateLocation(intStockID, strLocation, session("logged_username"))
			case else
				call displayLoanStock	
		end select
	else
		session("acc_department") = ""
		
		strAccountString = Trim(Request("account"))
		session("loan_account_code") = strAccountString
		
		call getEmployeeDetails(session("logged_username"))		
		call getLoanAccountDepartment(strAccountString)

		if trim(session("loan_user_initial_page")) = "" then
			session("loan_user_initial_page") = 1
		end if
			
		if len(request("account")) > 1 then
			session("loan_user_account") = trim(request("account"))
		end if  
			
		if len(request("name")) > 1 then
			session("loan_user_account_name") = trim(request("name"))
		end if			
		
		if session("emp_admin") <> 1 then
			if session("emp_admin") = 2 then
				if session("acc_department") = session("emp_department") then
					'response.write "YES this is your dept"
					call displayLoanStock
				else
					'response.write "NOT your dept!"
					Response.redirect "error.html"
				end if
			else
				if Trim(Request("account")) = session("emp_initial") then
					'response.write "YES this is your account only"
					call displayLoanStock
				else 
					'response.write "NOT your account!"
					Response.redirect "error.html"
				end if
			end if
		else
			'response.write "Admin only"	
			call displayLoanStock
		end if
								
	end if
end sub

call main

Dim strDisplayList, strMessageText, strAccountString
Dim strOrderNo, strOrderLine, strAccountCode, strAccountName, strDepartment, strItemCode, strSerialNo, strLIC, strLoanDate, strLocation, intStockID
%>
<div id="sticky_navigation_wrapper">
  <div id="sticky_navigation">
    <div class="demo_container">
      <form name="frmSearch" id="frmSearch" method="post" action="loan_user.asp?type=search" onsubmit="searchStock()">
        <div class="float_left"><%= session("loan_user_account") & "-" & session("account_code_ext") & " " & session("loan_user_account_name") & "" %></div>
        <div class="float_right"> Item / Serial
          <input type="text" name="txtSearch" maxlength="15" size="20" value="<%= request("txtSearch") %>" />
          <select name="cboSort" onchange="searchStock()">
            <option <% if session("loan_user_sort") = "oldest" then Response.Write " selected" end if%> value="oldest">Sort by: Loan (Old - New)</option>
            <option <% if session("loan_user_sort") = "latest" then Response.Write " selected" end if%> value="latest">Sort by: Loan (New - Old)</option>
            <option <% if session("loan_user_sort") = "product" then Response.Write " selected" end if%> value="product">Sort by: Item code (A-Z)</option>
            <option <% if session("loan_user_sort") = "expensive" then Response.Write " selected" end if%> value="expensive">Sort by: Value (High - Low)</option>
            <option <% if session("loan_user_sort") = "cheapest" then Response.Write " selected" end if%> value="cheapest">Sort by: Value (Low - High)</option>
            <option <% if session("loan_user_sort") = "serial" then Response.Write " selected" end if%> value="serial">Sort by: Serial</option>
          </select>
          <input type="button" name="btnSearch" value="Search" onclick="searchStock()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </div>
      </form>
    </div>
  </div>
</div>
<div align="center"> <br />
  <div style="padding-left:10px; float:left"><a href="loan_summary.asp">Loan Summary</a> <img src="images/forward_arrow.gif" /> Loan Stock</div>
  <div style="padding-right:10px; float:right;"><a href="view_basket.asp?account=<%= session("loan_user_account") & "" & session("account_code_ext") %>" target="_blank" style="font-weight:bold;">Workflow Summary</a> | <a href="list_auction.asp" target="_blank">Auction Stuff</a> | <a href="loan_renewal.asp" target="_blank">Renewal Approval</a> | <img src="images/icon_excel.jpg" /> <a href="export_loan_user.asp?search=<%= request("txtSearch") %>&account=<%= session("loan_user_account") %>&year=<%= session("loan_user_year") %>&month=<%= session("loan_user_month") %>&sort=<%= session("loan_user_sort") %>">Export</a>  | <img src="images/icon_excel.jpg" /> <a href="export_loan_user_location.asp?account=<%= session("loan_user_account") %>">Export with Location</a></div>
  <p><%= strMessageText %></p>
  <p><em>*Please note that Location is mandatory before renewing</em></p>
  <table cellspacing="0" cellpadding="5" class="loan_table" border="0">
    <tr class="loan_header_row">
      <td align="left" width="5%" nowrap="nowrap">Loan Sale / Return</td>
      <td align="left" width="5%" nowrap="nowrap">Auction</td>	 
      <td align="left" width="10%" nowrap="nowrap">Item Code</td>
      <td align="left" width="5%">Serial</td>      
      <td align="center" width="5%">Qty</td>
      <td align="right" width="5%" nowrap="nowrap">LIC $</td>
      <td align="right" width="10%" nowrap="nowrap">Loan Date</td>
      <td align="center" width="10%">Day Count</td>
      <td align="center" width="10%" nowrap="nowrap">Renewal Count</td>
      <td align="center" width="5%"></td>
      <td align="center" width="5%">Renewal</td>
      <td align="center" width="5%">Status</td>
      <td align="center" width="10%" nowrap="nowrap">Expiry Date</td>
      <td align="center" width="10%">Location*</td>
    </tr>
    <%= strDisplayList %>
  </table>
  <p><small><%= session("emp_department") %> - Logged in as: <%= session("logged_username") %> (<%= UCASE(trim(session("emp_initial"))) %>) - Admin: <%= session("emp_admin") %></small></p>
</div>
</body>
</html>