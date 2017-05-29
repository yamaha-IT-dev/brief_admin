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
<!--#include file="../include/connection_base.asp " -->
<!--#include file="../include/connection_local.asp " -->
<!--#include file="class/clsEmployee.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Loan Stock</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script language="JavaScript" type="text/javascript">
function searchStock(){    
    var strSearch	= document.forms[0].txtSearch.value;
	var strUser 	= document.forms[0].cboUser.value;	
	var strYear 	= document.forms[0].cboYear.value;
	var strMonth	= document.forms[0].cboMonth.value;
	var strSort	  	= document.forms[0].cboSort.value;
    document.location.href = 'loanstock.asp?type=search&txtSearch=' + strSearch + '&user=' + strUser + '&year=' + strYear + '&month=' + strMonth + '&sort=' + strSort;
}
    
function resetSearch(){
	document.location.href = 'loanstock.asp?type=reset';    
}  
</script>
</head>
<body>
<%
session.lcid = 2057
sub setSearch	
	Select case trim(request("type"))
		case "reset"
			session("admin_loanstock_search") 		= ""
			session("admin_loanstock_user") 			= ""			
			session("admin_loanstock_year") 			= ""
			session("admin_loanstock_month") 			= ""
			session("admin_loanstock_sort") 			= ""
			session("admin_loanstock_initial_page") 	= 1
		case "search"
			session("admin_loanstock_search") 		= trim(request("txtSearch"))
			session("admin_loanstock_user") 			= trim(request("user"))			
			session("admin_loanstock_year") 			= trim(request("year"))
			session("admin_loanstock_month") 			= trim(request("month"))
			session("admin_loanstock_sort") 			= trim(request("sort"))
			session("admin_loanstock_initial_page") 	= 1
	end Select
end sub

sub displayLoanStock
	dim iRecordCount
	iRecordCount = 0
    dim strSortBy
	dim strSortItem
    'dim strSearchTxt
    dim strSQL
	
	dim intOnHandsTotal
	dim intOnHandsTotalReserved
	dim intOnHandsTotalAllocated
	dim intOnHandsTotalAvailable
	dim intInTransitTotal
	dim intInTransitTotalReserved
	dim intInTransitTotalAllocated
	dim intInTransitTotalAvailable
	dim intTotalBackorder
	
	intOnHandsTotal = 0
	intOnHandsTotalReserved = 0
	intOnHandsTotalAllocated = 0
	intOnHandsTotalAvailable = 0
	intInTransitTotal = 0
	intInTransitTotalReserved = 0
	intInTransitTotalAllocated = 0
	intInTransitTotalAvailable = 0
	intTotalBackorder = 0
	
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
	
	if session("admin_loanstock_qty") = "" then
		session("admin_loanstock_qty") = "1"
	end if
	
	if session("admin_loanstock_sort") = "" then
		session("admin_loanstock_sort") = "oldest"
	end if
	
	strSQL = strSQL & "SELECT B9JUNO AS order_number, B9JUGY AS order_line, "
	strSQL = strSQL & " B9SKNO AS ship_number, B9SKGY AS ship_line, B9URKC AS dealer_code, B9SCSS AS warehouse, "
	strSQL = strSQL & " B9GREG AS item_group, B9SOSC AS item_code, B9SKSU - B9AHEN AS loan_qty, "
	strSQL = strSQL & "	B9SKJY, "
	strSQL = strSQL & "	B9SKJM, "
	strSQL = strSQL & "	B9SKJD, "
	strSQL = strSQL & " RIGHT('0' || B9SKJD,2)|| '/' || RIGHT ('0' || B9SKJM,2) || '/' || B9SKJY as transfer_date, "
	strSQL = strSQL & " case "
	strSQL = strSQL & " 	when B9STJN <> '00' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) "
	strSQL = strSQL & " 		else 0 "
	strSQL = strSQL & " 	end AS lic, B9SIBN AS serial_number, "
	strSQL = strSQL & " B9ASFN AS comment "
	strSQL = strSQL & " FROM BF9EP "
	strSQL = strSQL & " INNER JOIN EF2SP ON B9SOSC = E2SOSC "
	strSQL = strSQL & " WHERE B9AHEN < B9SKSU "
	strSQL = strSQL & "				AND (B9SOSC LIKE '%" & UCASE(trim(session("admin_loanstock_search"))) & "%' "
	strSQL = strSQL & "					OR B9JUNO LIKE '%" & UCASE(trim(session("admin_loanstock_search"))) & "%' "
	strSQL = strSQL & "					OR B9SIBN LIKE '%" & UCASE(trim(session("admin_loanstock_search"))) & "%' "
	strSQL = strSQL & "					OR B9SKNO LIKE '%" & UCASE(trim(session("admin_loanstock_search"))) & "%') "
	strSQL = strSQL & "				AND B9URKC LIKE '%" & UCASE(trim(session("admin_loanstock_user"))) & "%' "
	strSQL = strSQL & "				AND B9SKJY LIKE '%" & trim(session("admin_loanstock_year")) & "%' "
	strSQL = strSQL & "				AND B9SKJM LIKE '%" & trim(session("admin_loanstock_month")) & "%' "
	strSQL = strSQL & " AND (E2NGTY = "
	strSQL = strSQL & "	 (SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only) "
	strSQL = strSQL & "	AND E2NGTM = "
	strSQL = strSQL & "	 (SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only))"
	strSQL = strSQL & "		ORDER BY "
	
	select case session("admin_loanstock_sort")
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
			strSQL = strSQL & "		order_number"
		case "shipment"
			strSQL = strSQL & "		ship_number"
	end select
	
	'response.write strSQL
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
		
    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("admin_loanstock_initial_page")
	
		For intRecord = 1 To rs.PageSize 
			intOnHandsTotal = intOnHandsTotal + Cint(rs("lic"))
			
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if
			strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""update_loanstock.asp?owner=" & rs("dealer_code") & "&product=" & rs("item_code") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("dealer_code") & "</td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("item_code") & "</td>"
			strDisplayList = strDisplayList & "<td align=""left"">" & rs("serial_number") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("loan_qty") & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"">" & FormatNumber(rs("lic")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"">" & FormatDateTime(rs("transfer_date"),1) & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"">" & rs("order_number") & "</td>"
			strDisplayList = strDisplayList & "<td align=""right"">" & rs("ship_number") & "</td>"
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
	strDisplayList = strDisplayList & "<h1>Total LIC: $" & FormatNumber(intOnHandsTotal) & "</h1>"
	strDisplayList = strDisplayList & "<h2>Search results: <u>" & intRecordCount & "</u> stocks.</h2>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseBaseDataBase()
end sub

sub main
	call getEmployeeDetails(session("logged_username"))
	call setSearch 
	
	if trim(session("admin_loanstock_initial_page"))  = "" then
    	session("admin_loanstock_initial_page") = 1
	end if		
    
    call displayLoanStock
	
	response.Write "Logged in as: " & session("emp_username")
end sub

call main

dim strDisplayList
%>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
  <!-- #include file="include/header_loan.asp" -->
  <tr>
    <td valign="top" class="maincontent"><table width="1250" border="0">
        <tr>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" method="post" action="loanstock.asp?type=search" onsubmit="searchStock()">               
                <strong>Search:</strong> Product / Serial # / Order # / Ship #
                <input type="text" name="txtSearch" maxlength="15" size="20" value="<%= request("txtSearch") %>" />
                <select name="cboUser" onchange="searchStock()">
                  <option <% if session("admin_loanstock_user") = "" then Response.Write " selected" end if%> value="">All Users</option>
                  <option <% if session("admin_loanstock_user") = "7CAM01" then Response.Write " selected" end if%> value="7CAM01">Cameron Tait</option>
                  <option <% if session("admin_loanstock_user") = "7CH001" then Response.Write " selected" end if%> value="7CH001">Chris Herring</option>
                  <option <% if session("admin_loanstock_user") = "7DM001" then Response.Write " selected" end if%> value="7DM001">Dale Moore</option>
                  <option <% if session("admin_loanstock_user") = "7DMRS0" then Response.Write " selected" end if%> value="7DMRS0">Dale Moore (Roadshow)</option>
                  <option <% if session("admin_loanstock_user") = "7DH000" then Response.Write " selected" end if%> value="7DH000">Damien Henderson</option>
                  <option <% if session("admin_loanstock_user") = "7DH002" then Response.Write " selected" end if%> value="7DH002">Damien Henderson (Roadshow)</option>
                  <option <% if session("admin_loanstock_user") = "7DT001" then Response.Write " selected" end if%> value="7DT001">Dave Thwaites</option>
                  <option <% if session("admin_loanstock_user") = "7FED00" then Response.Write " selected" end if%> value="7FED00">Felix Elliot-Dedman</option>
                  <option <% if session("admin_loanstock_user") = "7EM001" then Response.Write " selected" end if%> value="7EM001">Euan McInnes</option>
                  <option <% if session("admin_loanstock_user") = "7EM002" then Response.Write " selected" end if%> value="7EM002">Euan McInnes (VOX Pedal)</option>
                  <option <% if session("admin_loanstock_user") = "7GN001" then Response.Write " selected" end if%> value="7GN001">George Nasr</option>
                  <option <% if session("admin_loanstock_user") = "7GL000" then Response.Write " selected" end if%> value="7GL000">Grant Lane</option>
                  <option <% if session("admin_loanstock_user") = "7JW001" then Response.Write " selected" end if%> value="7JW001">Jaclyn Williams</option>
                  <option <% if session("admin_loanstock_user") = "7JG000" then Response.Write " selected" end if%> value="7JG000">Jamie Goff</option>
                  <option <% if session("admin_loanstock_user") = "7AUDW0" then Response.Write " selected" end if%> value="7AUDW0">Jamie Goff (AUDW)</option>
                  <option <% if session("admin_loanstock_user") = "7JS001" then Response.Write " selected" end if%> value="7JS001">John Saccaro</option>
                  <option <% if session("admin_loanstock_user") = "7JP001" then Response.Write " selected" end if%> value="7JP001">Joseph Pantalleresco</option>
                  <option <% if session("admin_loanstock_user") = "7JD001" then Response.Write " selected" end if%> value="7JD001">Justin D'offay</option>
                  <option <% if session("admin_loanstock_user") = "7JD002" then Response.Write " selected" end if%> value="7JD002">Justin D'offay (Roadshow)</option>
                  <option <% if session("admin_loanstock_user") = "7KJ001" then Response.Write " selected" end if%> value="7KJ001">Kevin Johnson</option>
                  <option <% if session("admin_loanstock_user") = "7LB000" then Response.Write " selected" end if%> value="7LB000">Leon Blaher</option>
                  <option <% if session("admin_loanstock_user") = "7MC001" then Response.Write " selected" end if%> value="7MC001">Mark Condon</option>
                  <option <% if session("admin_loanstock_user") = "7ML001" then Response.Write " selected" end if%> value="7ML001">Mark Loey</option>
                  <option <% if session("admin_loanstock_user") = "7MT000" then Response.Write " selected" end if%> value="7MT000">Mathew Taylor</option>
                  <option <% if session("admin_loanstock_user") = "7MH000" then Response.Write " selected" end if%> value="7MH000">Mick Hughes</option>
                  <option <% if session("admin_loanstock_user") = "7NB001" then Response.Write " selected" end if%> value="7NB001">Nathan Biggin</option>
                  <option <% if session("admin_loanstock_user") = "7PW000" then Response.Write " selected" end if%> value="7PW000">Paul Wheeler</option>
                  <option <% if session("admin_loanstock_user") = "7BP000" then Response.Write " selected" end if%> value="7BP000">Peter Beveridge</option>
                  <option <% if session("admin_loanstock_user") = "7RW001" then Response.Write " selected" end if%> value="7RW001">Russell Wykes</option>
                  <option <% if session("admin_loanstock_user") = "7SG001" then Response.Write " selected" end if%> value="7SG001">Simon Goldsworthy</option>
                  <option <% if session("admin_loanstock_user") = "7SL001" then Response.Write " selected" end if%> value="7SL001">Steve Legg</option>
                  <option <% if session("admin_loanstock_user") = "7SVR01" then Response.Write " selected" end if%> value="7SVR01">Steven Vranch</option>
                  <option <% if session("admin_loanstock_user") = "7TM000" then Response.Write " selected" end if%> value="7TM000">Terry McMahon</option>
                  <option <% if session("admin_loanstock_user") = "7SMC01" then Response.Write " selected" end if%> value="7SMC01">Shaun McMahon</option>
                  <option <% if session("admin_loanstock_user") = "7WF001" then Response.Write " selected" end if%> value="7WF001">Wesley Fischer</option>
                  <option <% if session("admin_loanstock_user") = "7YME01" then Response.Write " selected" end if%> value="7YME01">YMEC Altona</option>
                  <option <% if session("admin_loanstock_user") = "7YME02" then Response.Write " selected" end if%> value="7YME02">YMEC Balwyn</option>
                  <option <% if session("admin_loanstock_user") = "7YME09" then Response.Write " selected" end if%> value="7YME09">YMEC Baulkham Hills</option>
                  <option <% if session("admin_loanstock_user") = "7YME05" then Response.Write " selected" end if%> value="7YME05">YMEC Carnegie</option>
                  <option <% if session("admin_loanstock_user") = "7YME12" then Response.Write " selected" end if%> value="7YME12">YMEC Morley</option>
                  <option <% if session("admin_loanstock_user") = "7YME08" then Response.Write " selected" end if%> value="7YME08">YMEC Strathmore</option>
                </select>
                <select name="cboYear" onchange="searchStock()">
                  <option <% if session("admin_loanstock_year") = "" then Response.Write " selected" end if%> value="">All years</option>
                  <option <% if session("admin_loanstock_year") = "2013" then Response.Write " selected" end if%> value="2013">2013 only</option>
                  <option <% if session("admin_loanstock_year") = "2012" then Response.Write " selected" end if%> value="2012">2012 only</option>
                  <option <% if session("admin_loanstock_year") = "2011" then Response.Write " selected" end if%> value="2011">2011 only</option>
                  <option <% if session("admin_loanstock_year") = "2010" then Response.Write " selected" end if%> value="2010">2010 only</option>
                  <option <% if session("admin_loanstock_year") = "2009" then Response.Write " selected" end if%> value="2009">2009 only</option>
                  <option <% if session("admin_loanstock_year") = "2008" then Response.Write " selected" end if%> value="2008">2008 only</option>
                </select>
                <select name="cboMonth" onchange="searchStock()">
                  <option <% if session("admin_loanstock_month") = "" then Response.Write " selected" end if%> value="">All months</option>
                  <option <% if session("admin_loanstock_month") = "1" then Response.Write " selected" end if%> value="1">January</option>
                  <option <% if session("admin_loanstock_month") = "2" then Response.Write " selected" end if%> value="2">February</option>
                  <option <% if session("admin_loanstock_month") = "3" then Response.Write " selected" end if%> value="3">March</option>
                  <option <% if session("admin_loanstock_month") = "4" then Response.Write " selected" end if%> value="4">April</option>
                  <option <% if session("admin_loanstock_month") = "5" then Response.Write " selected" end if%> value="5">May</option>
                  <option <% if session("admin_loanstock_month") = "6" then Response.Write " selected" end if%> value="6">June</option>
                  <option <% if session("admin_loanstock_month") = "7" then Response.Write " selected" end if%> value="7">July</option>
                  <option <% if session("admin_loanstock_month") = "8" then Response.Write " selected" end if%> value="8">August</option>
                  <option <% if session("admin_loanstock_month") = "9" then Response.Write " selected" end if%> value="9">September</option>
                  <option <% if session("admin_loanstock_month") = "10" then Response.Write " selected" end if%> value="10">October</option>
                  <option <% if session("admin_loanstock_month") = "11" then Response.Write " selected" end if%> value="11">November</option>
                  <option <% if session("admin_loanstock_month") = "12" then Response.Write " selected" end if%> value="12">December</option>
                </select>
                <select name="cboSort" onchange="searchStock()">
                  <option <% if session("admin_loanstock_sort") = "oldest" then Response.Write " selected" end if%> value="oldest">Sort by: Oldest transfer date</option>
                  <option <% if session("admin_loanstock_sort") = "latest" then Response.Write " selected" end if%> value="latest">Sort by: Latest transfer date</option>
                  <option <% if session("admin_loanstock_sort") = "product" then Response.Write " selected" end if%> value="product">Sort by: Product name (A-Z)</option>
                  <option <% if session("admin_loanstock_sort") = "expensive" then Response.Write " selected" end if%> value="expensive">Sort by: Most expensive</option>
                  <option <% if session("admin_loanstock_sort") = "cheapest" then Response.Write " selected" end if%> value="cheapest">Sort by: Cheapest</option>
                  <option <% if session("admin_loanstock_sort") = "serial" then Response.Write " selected" end if%> value="serial">Sort by: Serial no</option>
                  <option <% if session("admin_loanstock_sort") = "order" then Response.Write " selected" end if%> value="order">Sort by: Order no</option>
                  <option <% if session("admin_loanstock_sort") = "shipment" then Response.Write " selected" end if%> value="shipment">Sort by: Shipment no</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchStock()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
          <td valign="top" align="right"><img src="../logistics/images/icon_excel.jpg" width="15" height="15" /> <a href="export_loanstock.asp?search=<%= request("txtSearch") %>&user=<%= session("admin_loanstock_user") %>&year=<%= session("admin_loanstock_year") %>&month=<%= session("admin_loanstock_month") %>&sort=<%= session("admin_loanstock_sort") %>">Export</a></td>
        </tr>
      </table>
      <table cellspacing="1" cellpadding="3" width="1200" bgcolor="#CCCCCC">
        <tr class="innerdoctitle" align="center">
          <td align="left" width="5%"></td>
          <td align="left" width="15%">Loan Account</td>
          <td align="left" width="20%">Product</td>
          <td align="left" width="15%">Serial #</td>
          <td align="center" width="5%">Qty</td>
          <td align="right" width="5%">LIC ($)</td>
          <td align="right" width="15%">Transfer date</td>
          <td align="right" width="10%">Order #</td>
          <td align="right" width="10%">Ship #</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>  
</table>
</body>
</html>