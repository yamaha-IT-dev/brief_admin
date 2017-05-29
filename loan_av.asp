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
<!--#include file="../include/connection_base.asp " -->
<!--#include file="../include/connection_local.asp " -->
<!--#include file="class/clsEmployee.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Cache-control" content="no-store">
<title>Loan Stock AV</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<link rel="stylesheet" href="css/sticky-navigation.css" />
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"></script>
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
	var strUser 	= document.forms[0].cboUser.value;	
	var strYear 	= document.forms[0].cboYear.value;
	var strMonth	= document.forms[0].cboMonth.value;
	var strSort	  	= document.forms[0].cboSort.value;
    document.location.href = 'loan_av.asp?type=search&txtSearch=' + strSearch + '&user=' + strUser + '&year=' + strYear + '&month=' + strMonth + '&sort=' + strSort;
}
    
function resetSearch(){
	document.location.href = 'loan_av.asp?type=reset';    
}  
</script>
</head>
<body>
<%
session.lcid = 2057
sub setSearch	
	Select case trim(request("type"))
		case "reset"
			session("loan_av_search") 		= ""
			session("loan_av_user") 		= ""			
			session("loan_av_year") 		= ""
			session("loan_av_month") 		= ""
			session("loan_av_sort") 		= ""
			session("loan_av_initial_page") = 1
		case "search"
			session("loan_av_search") 		= trim(request("txtSearch"))
			session("loan_av_user") 		= trim(request("user"))			
			session("loan_av_year") 		= trim(request("year"))
			session("loan_av_month") 		= trim(request("month"))
			session("loan_av_sort") 		= trim(request("sort"))
			session("loan_av_initial_page") = 1
	end Select
end sub

sub displayLoanStock
	dim iRecordCount
	iRecordCount = 0
    dim strSortBy
	dim strSortItem
    dim strDays
    dim strSQL
	
	dim intTotalLIC
	dim intTotalQty
	
	intTotalLIC = 0
	intTotalQty = 0
	
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())
	
	'strSearchTxt = trim(Request("txtSearch"))
	
    call OpenBaseDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 5000
	
	if session("loan_av_qty") = "" then
		session("loan_av_qty") = "1"
	end if
	
	if session("loan_av_sort") = "" then
		session("loan_av_sort") = "oldest"
	end if
	
	strSQL = strSQL & "SELECT B9JUNO AS order_no, B9JUGY AS order_line, "
	strSQL = strSQL & " B9SKNO AS ship_number, B9SKGY AS ship_line, B9URKC AS dealer_code, B9SCSS AS warehouse, "
	strSQL = strSQL & " B9GREG AS item_group, B9SOSC AS item_code, B9SKSU - B9AHEN AS loan_qty, Y1KOM1, "
	strSQL = strSQL & "	B9SKJY, "
	strSQL = strSQL & "	B9SKJM, "
	strSQL = strSQL & "	B9SKJD, "
	strSQL = strSQL & " RIGHT('0' || B9SKJD,2)|| '/' || RIGHT ('0' || B9SKJM,2) || '/' || B9SKJY as loan_date, "
	strSQL = strSQL & " case "
	strSQL = strSQL & " 	when B9STJN <> '00' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) "
	strSQL = strSQL & " 		else 0 "
	strSQL = strSQL & " 	end AS lic, B9SIBN AS serial_number, "
	strSQL = strSQL & " B9ASFN AS comment "
	strSQL = strSQL & " FROM BF9EP "
	strSQL = strSQL & " INNER JOIN EF2SP ON B9SOSC = E2SOSC "
	strSQL = strSQL & " INNER JOIN YF1MP ON CONCAT(B9URKC,B9HSRC) = Y1KOKC "
	strSQL = strSQL & " WHERE B9AHEN < B9SKSU "
	strSQL = strSQL & "				AND (B9SOSC LIKE '%" & UCASE(trim(session("loan_av_search"))) & "%' "
	strSQL = strSQL & "					OR B9JUNO LIKE '%" & UCASE(trim(session("loan_av_search"))) & "%' "
	strSQL = strSQL & "					OR B9SIBN LIKE '%" & UCASE(trim(session("loan_av_search"))) & "%' "
	strSQL = strSQL & "					OR B9SKNO LIKE '%" & UCASE(trim(session("loan_av_search"))) & "%') "
	if session("loan_av_user") = "" then
		strSQL = strSQL & "			AND (B9URKC LIKE '%DM%' OR B9URKC LIKE '%DH%' OR B9URKC LIKE '%DT%' OR B9URKC LIKE '%GN%' OR B9URKC LIKE '%JD%' OR B9URKC LIKE '%ML%' OR B9URKC LIKE '%RW%' OR B9URKC LIKE '%SG%' OR B9URKC LIKE '%SM%' OR B9URKC LIKE '%WF%') "
	else
		strSQL = strSQL & "			AND B9URKC LIKE '%" & UCASE(trim(session("loan_av_user"))) & "%' "
	end if	
	strSQL = strSQL & "				AND B9SKJY LIKE '%" & trim(session("loan_av_year")) & "%' "
	strSQL = strSQL & "				AND B9SKJM LIKE '%" & trim(session("loan_av_month")) & "%' "
	strSQL = strSQL & " AND (E2NGTY = "
	strSQL = strSQL & "	 (SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only) "
	strSQL = strSQL & "	AND E2NGTM = "
	strSQL = strSQL & "	 (SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only))"
	strSQL = strSQL & "		ORDER BY "
	
	select case session("loan_av_sort")
		case "oldest"
			strSQL = strSQL & "		B9SKJY ASC, B9SKJM ASC, B9SKJD ASC"
			'strSQL = strSQL & "		transfer_year ASC, transfer_month ASC, transfer_day ASC"
		case "latest"
			strSQL = strSQL & "		B9SKJY DESC, B9SKJM DESC, B9SKJD DESC"
			'strSQL = strSQL & "		transfer_year DESC, transfer_month DESC, transfer_day DESC"
		case "product"
			strSQL = strSQL & "		item_code"
		case "expensive"
			strSQL = strSQL & "		lic DESC"
		case "cheapest"
			strSQL = strSQL & "		lic"
		case "serial"
			strSQL = strSQL & "		serial_number"
		case "order"
			strSQL = strSQL & "		order_no"
		case "shipment"
			strSQL = strSQL & "		ship_number"
	end select
	
	'response.write strSQL
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
		
    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("loan_av_initial_page")
	
		For intRecord = 1 To rs.PageSize 
			strDays = DateDiff("d",rs("loan_date"), strTodayDate)
			
			intTotalLIC = intTotalLIC + Cint(rs("lic"))
			intTotalQty = intTotalQty + Cint(rs("loan_qty"))
			
			dim strFirstExpiryDate
			strFirstExpiryDate = DateAdd("m", 3, rs("loan_date"))
			
			dim strFinalExpiryDate
			strFinalExpiryDate = DateAdd("m", 6, rs("loan_date"))
			
			if DateDiff("d",strFirstExpiryDate, strTodayDate) > 0 then
				strDisplayList = strDisplayList & "<tr class=""overdue_row"">"		
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			end if
			
			strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""view_loan.asp?order=" & Trim(rs("order_no")) & "&line=" & Trim(rs("order_line")) & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("dealer_code") & "</td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("Y1KOM1") & "</td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("item_code") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("serial_number") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("loan_qty") & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"">" & FormatNumber(rs("lic")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"">" & FormatDateTime(rs("loan_date"),1) & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>"
			if DateDiff("d",strFirstExpiryDate, strTodayDate) > 0 then
				strDisplayList = strDisplayList & " <span style=""color:red"">"
			end if
			strDisplayList = strDisplayList & strDays & "</span></td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""9"" align=""center"">No stocks found.</td></tr>"
	end if
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""9"" bgcolor=""white"" align=""center"">"
	strDisplayList = strDisplayList & "<h1>Total LIC: $" & FormatNumber(intTotalLIC) & "</h1>"
	strDisplayList = strDisplayList & "<h1>Total Qty: <u>" & intTotalQty & "</u> stocks</h1>"
	strDisplayList = strDisplayList & "<h3>Search results: " & intRecordCount & " records</h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseBaseDataBase()
end sub

sub main
	call getEmployeeDetails(session("logged_username"))
	call setSearch 
	
	if trim(session("loan_av_initial_page"))  = "" then
    	session("loan_av_initial_page") = 1
	end if		
    
    call displayLoanStock
end sub

call main

dim strDisplayList
%>
<div id="sticky_navigation_wrapper">
  <div id="sticky_navigation">
    <div class="demo_container">
      <form name="frmSearch" id="frmSearch" method="post" action="loan_av.asp?type=search" onsubmit="searchStock()">
        <font size="+2"><strong>Search AV Loan Stock:</strong> Item / Serial</font>
        <input type="text" name="txtSearch" maxlength="15" size="15" value="<%= request("txtSearch") %>" />
        <select name="cboUser" onchange="searchStock()">
          <option <% if session("loan_av_user") = "" then Response.Write " selected" end if%> value="">All AV Users</option>
          <option <% if session("loan_av_user") = "7DM001" then Response.Write " selected" end if%> value="7DM001">Dale Moore</option>
          <option <% if session("loan_av_user") = "7DMRS0" then Response.Write " selected" end if%> value="7DMRS0">Dale Moore (RS)</option>
          <option <% if session("loan_av_user") = "7DH000" then Response.Write " selected" end if%> value="7DH000">Damien Henderson</option>
          <option <% if session("loan_av_user") = "7DH002" then Response.Write " selected" end if%> value="7DH002">Damien Henderson (RS)</option>
          <option <% if session("loan_av_user") = "7DT001" then Response.Write " selected" end if%> value="7DT001">Dave Thwaites</option>
          <option <% if session("loan_av_user") = "7GN001" then Response.Write " selected" end if%> value="7GN001">George Nasr</option>
          <option <% if session("loan_av_user") = "7JD001" then Response.Write " selected" end if%> value="7JD001">Justin D'offay</option>
          <option <% if session("loan_av_user") = "7JD002" then Response.Write " selected" end if%> value="7JD002">Justin D'offay (RS)</option>
          <option <% if session("loan_av_user") = "7ML001" then Response.Write " selected" end if%> value="7ML001">Mark Loey</option>
          <option <% if session("loan_av_user") = "7ML002" then Response.Write " selected" end if%> value="7ML002">Mark Lapthorne</option>
          <option <% if session("loan_av_user") = "7RW001" then Response.Write " selected" end if%> value="7RW001">Russell Wykes</option>
          <option <% if session("loan_av_user") = "7SG001" then Response.Write " selected" end if%> value="7SG001">Simon Goldsworthy</option>
          <option <% if session("loan_av_user") = "7SMC01" then Response.Write " selected" end if%> value="7SMC01">Shaun McMahon</option>
          <option <% if session("loan_av_user") = "7WF001" then Response.Write " selected" end if%> value="7WF001">Wesley Fischer</option>
        </select>
        <select name="cboYear" onchange="searchStock()">
          <option <% if session("loan_av_year") = "" then Response.Write " selected" end if%> value="">All years</option>
          <option <% if session("loan_av_year") = "2013" then Response.Write " selected" end if%> value="2013">2013 only</option>
          <option <% if session("loan_av_year") = "2012" then Response.Write " selected" end if%> value="2012">2012 only</option>
          <option <% if session("loan_av_year") = "2011" then Response.Write " selected" end if%> value="2011">2011 only</option>
          <option <% if session("loan_av_year") = "2010" then Response.Write " selected" end if%> value="2010">2010 only</option>
          <option <% if session("loan_av_year") = "2009" then Response.Write " selected" end if%> value="2009">2009 only</option>
          <option <% if session("loan_av_year") = "2008" then Response.Write " selected" end if%> value="2008">2008 only</option>
        </select>
        <select name="cboMonth" onchange="searchStock()">
          <option <% if session("loan_av_month") = "" then Response.Write " selected" end if%> value="">All months</option>
          <option <% if session("loan_av_month") = "1" then Response.Write " selected" end if%> value="1">January</option>
          <option <% if session("loan_av_month") = "2" then Response.Write " selected" end if%> value="2">February</option>
          <option <% if session("loan_av_month") = "3" then Response.Write " selected" end if%> value="3">March</option>
          <option <% if session("loan_av_month") = "4" then Response.Write " selected" end if%> value="4">April</option>
          <option <% if session("loan_av_month") = "5" then Response.Write " selected" end if%> value="5">May</option>
          <option <% if session("loan_av_month") = "6" then Response.Write " selected" end if%> value="6">June</option>
          <option <% if session("loan_av_month") = "7" then Response.Write " selected" end if%> value="7">July</option>
          <option <% if session("loan_av_month") = "8" then Response.Write " selected" end if%> value="8">August</option>
          <option <% if session("loan_av_month") = "9" then Response.Write " selected" end if%> value="9">September</option>
          <option <% if session("loan_av_month") = "10" then Response.Write " selected" end if%> value="10">October</option>
          <option <% if session("loan_av_month") = "11" then Response.Write " selected" end if%> value="11">November</option>
          <option <% if session("loan_av_month") = "12" then Response.Write " selected" end if%> value="12">December</option>
        </select>
        <select name="cboSort" onchange="searchStock()">
          <option <% if session("loan_av_sort") = "oldest" then Response.Write " selected" end if%> value="oldest">Sort: Oldest loan date</option>
          <option <% if session("loan_av_sort") = "latest" then Response.Write " selected" end if%> value="latest">Sort: Latest loan date</option>
          <option <% if session("loan_av_sort") = "product" then Response.Write " selected" end if%> value="product">Sort: Item code (A-Z)</option>
          <option <% if session("loan_av_sort") = "expensive" then Response.Write " selected" end if%> value="expensive">Sort: Most expensive</option>
          <option <% if session("loan_av_sort") = "cheapest" then Response.Write " selected" end if%> value="cheapest">Sort: Cheapest</option>
          <option <% if session("loan_av_sort") = "serial" then Response.Write " selected" end if%> value="serial">Sort: Serial no</option>
          <option <% if session("loan_av_sort") = "order" then Response.Write " selected" end if%> value="order">Sort: Order no</option>
          <option <% if session("loan_av_sort") = "shipment" then Response.Write " selected" end if%> value="shipment">Sort: Shipment no</option>
        </select>
        <input type="button" name="btnSearch" value="Search" onclick="searchStock()" />
        <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
      </form>
    </div>
  </div>
</div>
<p align="center"><small>You are logged in as: <%= session("logged_username") %> (<%= UCASE(trim(session("emp_initial"))) %>)</small></p>
<p align="center"><img src="images/icon_excel.jpg" /> <a href="export_loan_av.asp?search=<%= request("txtSearch") %>&user=<%= session("loan_av_user") %>&year=<%= session("loan_av_year") %>&month=<%= session("loan_av_month") %>&sort=<%= session("loan_av_sort") %>">Export</a></p>
<table cellspacing="0" cellpadding="5" align="center" class="loan_table">
  <tr class="loan_header_row">
    <td align="left" width="5%"></td>
    <td align="left" width="10%">Account</td>
    <td align="left" width="20%">Name</td>
    <td align="left" width="10%">Item code</td>
    <td align="left" width="10%">Serial</td>
    <td align="center" width="5%">Qty</td>
    <td align="right" width="10%">LIC $</td>
    <td align="right" width="15%">Loan date</td>
    <td align="center" width="15%">Day count</td>
  </tr>
  <%= strDisplayList %>
</table>
</body>
</html>