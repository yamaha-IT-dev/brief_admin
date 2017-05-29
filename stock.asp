<!--#include file="../include/connection_base.asp " -->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Warehouse Inventory Movement</title>
<link rel="stylesheet" href="include/stylesheet.css">
<link rel="stylesheet" href="css/sticky-navigation.css">
<script src="//code.jquery.com/jquery.js"></script>
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
    var strSearch 		= document.forms[0].txtSearch.value;
	var strUser  		= document.forms[0].cboUser.value;
	var intMonth		= document.forms[0].cboMonth.value;
	var strVendor		= document.forms[0].cboVendor.value;
	var strWarehouse	= document.forms[0].cboWarehouse.value;
	var strSort  		= document.forms[0].cboSort.value;
	
    document.location.href = 'stock.asp?type=search&txtSearch=' + strSearch + '&user=' + strUser + '&month=' + intMonth + '&vendor=' + strVendor + '&warehouse=' + strWarehouse + '&sort=' + strSort;
}

function resetSearch(){
	document.location.href = 'stock.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			Session("inventory_search") 		= ""
			Session("inventory_user") 			= ""
			Session("inventory_month") 			= ""
			Session("inventory_vendor") 		= ""
			Session("inventory_warehouse") 		= ""
			Session("inventory_sort") 			= ""
			Session("inventory_initial_page") 	= 1
		case "search"
			Session("inventory_search") 		= Trim(Request("txtSearch"))
			Session("inventory_user") 			= Trim(Request("user"))
			Session("inventory_month") 			= Trim(Request("month"))
			Session("inventory_vendor")			= Trim(Request("vendor"))
			Session("inventory_warehouse")		= Trim(Request("warehouse"))
			Session("inventory_sort") 			= Trim(Request("sort"))
			Session("inventory_initial_page") 	= 1
	end select
end sub

sub displayStock
	dim strSQL
	
	dim intRecordCount
	
	dim iRecordCount
	iRecordCount = 0    
    		
	dim strTodayDate	
	strTodayDate = FormatDateTime(Date())

    call OpenBaseDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100
	
	strSQL = "SELECT E1OPEC AS OP, E1NSKY, E1NSKM, E1NSKD, E1NSKY * 10000 +  E1NSKM * 100 + E1NSKD AS INV_MOV_DATE, E1SOSC AS PRODUCT, E1SKNO AS SHIPMENT, "
	strSQL = strSQL & "	((E1AKKB*-2)+1)* E1NSKS AS QTY, E1SISC AS VENDOR_CODE, E1SOCD AS WAREHOUSE "
	strSQL = strSQL & "	FROM EF1BP "
	strSQL = strSQL & "	WHERE ("
	strSQL = strSQL & "			E1SOSC LIKE '%" & Ucase(Session("inventory_search")) & "%' "
	strSQL = strSQL & "			OR E1SKNO LIKE '%" & Session("inventory_search") & "%')"
	if Session("inventory_month") <> "" then
		strSQL = strSQL & " AND E1NSKM = '" & Trim(Session("inventory_month")) & "' "
	end if
	strSQL = strSQL & "		AND E1TRTI = 'AH'"
	strSQL = strSQL & "		AND E1SKKI <> 'D'"
	strSQL = strSQL & "		AND E1NSKY >= 2014"	
	strSQL = strSQL & "		AND E1OPEC LIKE '%" & Session("inventory_user") & "%' "
	strSQL = strSQL & "		AND E1SISC LIKE '%" & Session("inventory_vendor") & "%' "
	strSQL = strSQL & "		AND E1SOCD LIKE '%" & Session("inventory_warehouse") & "%' "
	strSQL = strSQL & "	ORDER BY "
			
	select case Session("inventory_sort")
		case "oldest"
			strSQL = strSQL & "INV_MOV_DATE"
		case "operator"
			strSQL = strSQL & "OP"
		case "product"
			strSQL = strSQL & "PRODUCT"
		case "shipment"
			strSQL = strSQL & "SHIPMENT"
		case "qty"
			strSQL = strSQL & "QTY DESC"
		case "vendor"
			strSQL = strSQL & "VENDOR_CODE"
		case "warehouse"
			strSQL = strSQL & "WAREHOUSE"	
		case else
			strSQL = strSQL & "INV_MOV_DATE DESC"
	end select
	
	rs.Open strSQL, conn

	'Response.Write strSQL

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			Session("inventory_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			Session("inventory_initial_page") = intpage

			if Session("inventory_initial_page") < 1 then Session("inventory_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			Session("inventory_initial_page") = intpage

			if Session("inventory_initial_page") > intPageCount then Session("inventory_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			Session("inventory_initial_page") = intpage
    end select
	
    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = Session("inventory_initial_page")

		For intRecord = 1 To rs.PageSize
			if rs("OP") <> "CD" and rs("OP") <> "KT" and rs("OP") <> "JS" and ((rs("VENDOR_CODE") <> "3OL" and rs("WAREHOUSE") <> "3T") or (rs("VENDOR_CODE") <> "3T" and rs("WAREHOUSE") <> "3OL")) then
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			end if
			
			strDisplayList = strDisplayList & "<td>" & rs("OP") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("E1NSKD") & " "			
			Select Case trim(rs("E1NSKM"))				
				case 1
					strDisplayList = strDisplayList & "Jan"
				case 2
					strDisplayList = strDisplayList & "Feb"
				case 3
					strDisplayList = strDisplayList & "Mar"
				case 4
					strDisplayList = strDisplayList & "April"
				case 5
					strDisplayList = strDisplayList & "May"
				case 6
					strDisplayList = strDisplayList & "June"
				case 7
					strDisplayList = strDisplayList & "July"
				case 8
					strDisplayList = strDisplayList & "Aug"	
				case 9
					strDisplayList = strDisplayList & "Sep"
				case 10
					strDisplayList = strDisplayList & "Oct"
				case 11
					strDisplayList = strDisplayList & "Nov"
				case 12
					strDisplayList = strDisplayList & "Dec"	
			end select
			strDisplayList = strDisplayList & " " & rs("E1NSKY") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("PRODUCT") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("SHIPMENT") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("QTY") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("VENDOR_CODE") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("WAREHOUSE") & "</td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""7"">No record found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""7"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""stock.asp"" method=""post"">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & Session("inventory_initial_page") & ">"

	if Session("inventory_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if Session("inventory_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
	
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strSort & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & Session("inventory_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & ""
    strDisplayList = strDisplayList & "</h2></form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseBaseDataBase()
end sub

sub main
	call setSearch
	
	if trim(Session("inventory_initial_page")) = "" then
    	Session("inventory_initial_page") = 1
	end if
	
    call displayStock
end sub

call main

dim rs, intPageCount, intpage, intRecord, strDisplayList

session("username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)

if session("username") = "harsonos" or session("username") = "gandig" or session("username") = "tasih" or session("username") = "craigd" or session("username") = "kurtt" or session("username") = "johannas" then
%>
<div id="sticky_navigation_wrapper">
  <div id="sticky_navigation">
    <div class="demo_container" style="padding-left:15px;">
      <form name="frmSearch" id="frmSearch" action="stock.asp?type=search" method="post" onsubmit="searchStock()">
        Search:
        <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" placeholder="Product / Shipment no" />
        <select name="cboUser" onchange="searchStock()">
          <option value="">All Operators</option>
          <option <% if Session("inventory_user") = "CD" then Response.Write " selected" end if%> value="CD">Craig</option>
          <option <% if Session("inventory_user") = "JS" then Response.Write " selected" end if%> value="JS">Johanna</option>
          <option <% if Session("inventory_user") = "KT" then Response.Write " selected" end if%> value="KT">Kurt</option>
          <option <% if Session("inventory_user") = "SC" then Response.Write " selected" end if%> value="SC">Shane</option>
          <option <% if Session("inventory_user") = "TK" then Response.Write " selected" end if%> value="TK">Tony</option>
        </select>
        <select name="cboMonth" onchange="searchStock()">
          <option <% if Session("inventory_month") = "" then Response.Write " selected" end if%> value="">All Months</option>
          <option <% if Session("inventory_month") = "1" then Response.Write " selected" end if%> value="1">January</option>
          <option <% if Session("inventory_month") = "2" then Response.Write " selected" end if%> value="2">February</option>
          <option <% if Session("inventory_month") = "3" then Response.Write " selected" end if%> value="3">March</option>
          <option <% if Session("inventory_month") = "4" then Response.Write " selected" end if%> value="4">April</option>
          <option <% if Session("inventory_month") = "5" then Response.Write " selected" end if%> value="5">May</option>
          <option <% if Session("inventory_month") = "6" then Response.Write " selected" end if%> value="6">June</option>
          <option <% if Session("inventory_month") = "7" then Response.Write " selected" end if%> value="7">July</option>
          <option <% if Session("inventory_month") = "8" then Response.Write " selected" end if%> value="8">August</option>
          <option <% if Session("inventory_month") = "9" then Response.Write " selected" end if%> value="9">September</option>
          <option <% if Session("inventory_month") = "10" then Response.Write " selected" end if%> value="10">October</option>
          <option <% if Session("inventory_month") = "11" then Response.Write " selected" end if%> value="11">November</option>
          <option <% if Session("inventory_month") = "12" then Response.Write " selected" end if%> value="12">December</option>
        </select>
        <select name="cboVendor" onchange="searchStock()">
          <option <% if Session("inventory_vendor") = "" then Response.Write " selected" end if%> value="">All Vendors (Source)</option>
          <option <% if Session("inventory_vendor") = "3K" then Response.Write " selected" end if%> value="3K">3K</option>
          <option <% if Session("inventory_vendor") = "3L" then Response.Write " selected" end if%> value="3L">3L</option>
          <option <% if Session("inventory_vendor") = "3ND" then Response.Write " selected" end if%> value="3ND">3ND</option>
          <option <% if Session("inventory_vendor") = "3OL" then Response.Write " selected" end if%> value="3OL">3OL</option>
          <option <% if Session("inventory_vendor") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
          <option <% if Session("inventory_vendor") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
          <option <% if Session("inventory_vendor") = "3TH" then Response.Write " selected" end if%> value="3TH">3TH</option>
          <option <% if Session("inventory_vendor") = "3XL" then Response.Write " selected" end if%> value="3XL">3XL</option>
        </select>
        <select name="cboWarehouse" onchange="searchStock()">
          <option <% if Session("inventory_warehouse") = "" then Response.Write " selected" end if%> value="">All Warehouses (Destination)</option>
          <option <% if Session("inventory_warehouse") = "3K" then Response.Write " selected" end if%> value="3K">3K</option>
          <option <% if Session("inventory_warehouse") = "3L" then Response.Write " selected" end if%> value="3L">3L</option>
          <option <% if Session("inventory_warehouse") = "3ND" then Response.Write " selected" end if%> value="3ND">3ND</option>
          <option <% if Session("inventory_warehouse") = "3OL" then Response.Write " selected" end if%> value="3OL">3OL</option>
          <option <% if Session("inventory_warehouse") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
          <option <% if Session("inventory_warehouse") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
          <option <% if Session("inventory_warehouse") = "3TH" then Response.Write " selected" end if%> value="3TH">3TH</option>
          <option <% if Session("inventory_warehouse") = "3XL" then Response.Write " selected" end if%> value="3XL">3XL</option>
        </select>
        <select name="cboSort" onchange="searchStock()">
          <option <% if Session("inventory_sort") = "latest" then Response.Write " selected" end if%> value="latest">Sort by: Latest</option>
          <option <% if Session("inventory_sort") = "oldest" then Response.Write " selected" end if%> value="oldest">Sort by: Oldest</option>
          <option <% if Session("inventory_sort") = "operator" then Response.Write " selected" end if%> value="operator">Sort by: Operator</option>
          <option <% if Session("inventory_sort") = "product" then Response.Write " selected" end if%> value="product">Sort by: Product</option>
          <option <% if Session("inventory_sort") = "shipment" then Response.Write " selected" end if%> value="shipment">Sort by: Shipment</option>
          <option <% if Session("inventory_sort") = "qty" then Response.Write " selected" end if%> value="qty">Sort by: Qty</option>
          <option <% if Session("inventory_sort") = "vendor" then Response.Write " selected" end if%> value="vendor">Sort by: Vendor (Source)</option>
          <option <% if Session("inventory_sort") = "warehouse" then Response.Write " selected" end if%> value="warehouse">Sort by: Warehouse (Destination)</option>
        </select>
        <input type="button" name="btnSearch" value="Search" onclick="searchStock()" />
        <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
      </form>
    </div>
  </div>
</div>
<div style="padding-top:10px;padding-left:15px;">
  <h1>Warehouse Inventory Movement</h1>
  <h3><a href="export_stock.asp">Export this</a></h3>
  <table cellspacing="0" cellpadding="5" width="900" border="0">
    <tr class="innerdoctitle">
      <td>Operator</td>
      <td>Invoice Move Date</td>
      <td>Product</td>
      <td>Shipment</td>
      <td>Qty</td>
      <td>Vendor</td>
      <td>Warehouse</td>
    </tr>
    <%= strDisplayList %>
  </table>
</div>
<% else %>
<h1>You are not authorised to view this.</h1>
<% end if %>
</body>
</html>