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
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!--#include file="class/clsLoanAccount.asp " -->
<!--#include file="class/clsLoanBase.asp " -->
<!--#include file="class/clsLoanLocation.asp " -->
<!--#include file="class/clsLoanTransfer.asp " -->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if lt IE 9]>
	<script src="//oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
	<script src="//oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
<![endif]-->
<title>Loan Stock User</title>
<link rel="stylesheet" href="css/sticky-navigation.css">
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/header.css">
<link href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css" rel="stylesheet">
<script src="../include/generic_form_validations.js"></script>
<script>
function searchStock(){    
    var strSearch	= document.forms[0].txtSearch.value;
	var strSort	  	= document.forms[0].cboSort.value;
    document.location.href = 'loan-user.asp?type=search&txtSearch=' + strSearch + '&sort=' + strSort;
}
    
function resetSearch(){
	document.location.href = 'loan-user.asp?type=reset';    
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
	strSQL = strSQL & "		B9AHEN, B9SKSU, lic, qty, loan_year, loan_month, loan_day, loan_date, "
	strSQL = strSQL & " 	stockLocation, stockRenewalCounter, stockID, stockDateCreated, stockCreatedBy, stockDateModified, stockModifiedBy, traModelNo, traStatus "
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
	strSQL = strSQL & "		LEFT JOIN tbl_loan_transfer ON traOrderNo = order_no AND traOrderLine = order_line "
	strSQL = strSQL & " WHERE B9AHEN < B9SKSU "
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
			intDays = DateDiff("d",rs("loan_date"), strTodayDate)	
			session("account_code_ext") = rs("account_code_ext")
			
			intTotalLIC = intTotalLIC + Cint(rs("lic"))
			intTotalQty = intTotalQty + Cint(rs("qty"))			

			strDisplayList = strDisplayList & "<tr>"			
			strDisplayList = strDisplayList & "<td>"
			if IsNull(rs("traModelNo")) or rs("traStatus") = 2 then
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_auction"" id=""form_auction"" onsubmit=""return submitAuctionForm(this)"">"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""action"" value=""Auction"">"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""account_code"" value=""" & Trim(rs("account_code")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""item_code"" value=""" & Trim(rs("item_code")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""serial_no"" value=""" & Trim(rs("serial_no")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_no"" value=""" & Trim(rs("order_no")) & """>"
				strDisplayList = strDisplayList & "		<input type=""hidden"" name=""order_line"" value=""" & Trim(rs("order_line")) & """>"	
				strDisplayList = strDisplayList & "		<input type=""submit"" value=""Transfer"" class=""btn btn-info"" />"
				strDisplayList = strDisplayList & "	</form>"
			else
				select case rs("traStatus")
					case 1
						strDisplayList = strDisplayList & "<font color=blue>" & " In-progress</font>"	
					case 2
						strDisplayList = strDisplayList & "<font color=red>" & " Rejected</font>"
					case 0
						strDisplayList = strDisplayList & "<font color=green>" & " Completed</font>"	
				end select				
				'strDisplayList = strDisplayList & rs("traStatus")
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td><a name=""" & Trim(rs("order_no")) & "" & Trim(rs("order_line")) & """></a><strong>" & rs("item_code") & "</strong></td>"
			strDisplayList = strDisplayList & "<td>" & rs("serial_no") & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("qty") & "</td>"
			strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("lic")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & FormatDateTime(rs("loan_date"),1) & "</td>"			
			strDisplayList = strDisplayList & "<td>"
			if intDays > 90 then
				strDisplayList = strDisplayList & " <span style=""color:red"">"
			end if
			strDisplayList = strDisplayList & intDays & "</span>"
			strDisplayList = strDisplayList & "</td>"			
			strDisplayList = strDisplayList & "<td>"
			if IsNull(rs("stockLocation")) then
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_add_location"" id=""form_add_location"" onsubmit=""return validateAddLocation(this)"">"
				strDisplayList = strDisplayList & "	<table>"
				strDisplayList = strDisplayList & "		<tr>"
				strDisplayList = strDisplayList & "			<td align=""center"">"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""action"" value=""Add"">"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""order_no"" value=""" & Trim(rs("order_no")) & """>"
				strDisplayList = strDisplayList & "				<input type=""hidden"" name=""order_line"" value=""" & Trim(rs("order_line")) & """>"
				strDisplayList = strDisplayList & "				<input type=""text"" id=""txtLocation"" name=""txtLocation"" class=""form-control"" maxlength=""100"" size=""30"" required>"
				strDisplayList = strDisplayList & "			</td>"
				strDisplayList = strDisplayList & "			<td class=""save-column""><input type=""submit"" value=""Save"" class=""btn btn-primary"" /></td>"
				strDisplayList = strDisplayList & "		</tr>"
				strDisplayList = strDisplayList & "	</table>"
				strDisplayList = strDisplayList & "	</form>"
			else
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_update_location"" id=""form_update_location"" onsubmit=""return validateAddLocation(this)"">"
				strDisplayList = strDisplayList & "	<table>"
				strDisplayList = strDisplayList & "	<tr>"
				strDisplayList = strDisplayList & "		<td align=""center"">"
				strDisplayList = strDisplayList & "			<input type=""hidden"" name=""action"" value=""Update"">"
				strDisplayList = strDisplayList & "			<input type=""hidden"" name=""stock_id"" value=""" & Trim(rs("stockID")) & """>"
				strDisplayList = strDisplayList & "<input type=""text"" id=""txtLocation"" name=""txtLocation"" class=""form-control"" maxlength=""100"" size=""30"" value=""" & rs("stockLocation") & """ required>"
				strDisplayList = strDisplayList & "		</td>"
				strDisplayList = strDisplayList & "		<td class=""save-column""><input type=""submit"" value=""Save"" class=""btn btn-primary"" /></td>"
				strDisplayList = strDisplayList & "	</tr>"
				strDisplayList = strDisplayList & "	</table>"
				strDisplayList = strDisplayList & "	</form>"
			end if
			strDisplayList = strDisplayList & "</td>"	
			
			if Len(rs("stockModifiedBy")) > 1 then
				strDisplayList = strDisplayList & "<td><strong>" & rs("stockModifiedBy") & "</strong><br>" & rs("stockDateModified") & "</td>"
			else
				strDisplayList = strDisplayList & "<td><strong>" & rs("stockCreatedBy") & "</strong><br>" & rs("stockDateCreated") & "</td>"
			end if
			
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr><td colspan=""9"" align=""center"">No stocks found.</td></tr>"
	end if
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""9"" align=""center"">"
	strDisplayList = strDisplayList & "<h2>Total Value: $" & FormatNumber(intTotalLIC) & "</h2>"
	strDisplayList = strDisplayList & "<h3>Total Items: " & intTotalQty & "</h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	
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
			case "Auction"				
				call newLoanTransfer(strAccountCode, strItemCode, strSerialNo, strOrderNo, strOrderLine)
			case "Add"
				call addLoanLocation(strOrderNo, strOrderLine, strLocation, session("logged_username"))
			case "Update"
				call updateLocation(intStockID, strLocation, session("logged_username"))
			'case else
			'	call displayLoanStock	
		end select
	
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
				
		call setSearch
		call displayLoanStock
end sub

call main

Dim strDisplayList, strMessageText, strAccountString
Dim strOrderNo, strOrderLine, strAccountCode, strAccountName, strDepartment, strItemCode, strSerialNo, strLIC, strLoanDate, strLocation, intStockID
%>
<div class="blog-masthead">
  <div class="container">
    <nav class="blog-nav"> <a class="blog-nav-item active"><i class="fa fa-home fa-lg"></i></a> <a class="blog-nav-item" href="loan-transfer.asp">Transfer</a> </nav>
  </div>
</div>
<div class="container"> <br>
  <ol class="breadcrumb">
    <li><a href="loan_summary.asp">Loan Summary</a></li>
    <li class="active">View Loan Stock</li>
  </ol>
  <h1 class="page-header"><i class="fa fa-search-plus"></i> <%= session("loan_user_account") & "-" & session("account_code_ext") & "" %></h1>
  <form name="frmSearch" id="frmSearch" method="post" action="loan-user.asp?type=search" onsubmit="searchStock()">
    <div class="row">
      <div class="form-group col-lg-4">
        <input type="text" class="form-control" name="txtSearch" maxlength="20" size="45" placeholder="Search Item code / Serial no" value="<%= request("txtSearch") %>" />
      </div>
      <div class="form-group col-lg-3">
        <select name="cboSort" class="form-control" onchange="searchStock()">
          <option <% if session("loan_user_sort") = "oldest" then Response.Write " selected" end if%> value="oldest">Sort by: Loan (Old - New)</option>
          <option <% if session("loan_user_sort") = "latest" then Response.Write " selected" end if%> value="latest">Sort by: Loan (New - Old)</option>
          <option <% if session("loan_user_sort") = "product" then Response.Write " selected" end if%> value="product">Sort by: Item code (A-Z)</option>
          <option <% if session("loan_user_sort") = "expensive" then Response.Write " selected" end if%> value="expensive">Sort by: Value (High - Low)</option>
          <option <% if session("loan_user_sort") = "cheapest" then Response.Write " selected" end if%> value="cheapest">Sort by: Value (Low - High)</option>
          <option <% if session("loan_user_sort") = "serial" then Response.Write " selected" end if%> value="serial">Sort by: Serial no</option>
        </select>
      </div>
      <div class="form-group col-lg-5">
        <input type="button" class="btn btn-primary" name="btnSearch" value="Search &raquo;" onclick="searchStock()" />
        <input type="button" class="btn btn-primary" name="btnReset" value="Reset" onclick="resetSearch()" />
      </div>
    </div>
  </form>
  <p><%= strMessageText %></p>
  <div class="table-responsive">
    <table class="table table-striped">
      <thead>
        <tr>
          <td>Transfer</td>
          <td>Item Code</td>
          <td>Serial</td>
          <td>Qty</td>
          <td>LIC $</td>
          <td>Loan Date</td>
          <td>Day</td>
          <td>Location</td>
          <td></td>
        </tr>
      </thead>
      <tbody>
        <%= strDisplayList %>
      </tbody>
    </table>
  </div>
  <p align="center"> <a href="export_loan_user.asp?search=<%= request("txtSearch") %>&account=<%= session("loan_user_account") %>&year=<%= session("loan_user_year") %>&month=<%= session("loan_user_month") %>&sort=<%= session("loan_user_sort") %>">
    <button type="button" class="btn btn-success"><i class="fa fa-download"></i> Export</button>
    </a> <a href="export_loan_user_location.asp?account=<%= session("loan_user_account") %>">
    <button type="button" class="btn btn-success"><i class="fa fa-download"></i> Export with Location</button>
    </a></p>
  <p align="center"><small><%= session("emp_department") %> - Logged in as: <%= session("logged_username") %> (<%= UCASE(trim(session("emp_initial"))) %>) - Admin: <%= session("emp_admin") %></small></p>
</div>
<script src="//code.jquery.com/jquery.js"></script> 
<script src="bootstrap/js/bootstrap.min.js"></script>
</body>
</html>