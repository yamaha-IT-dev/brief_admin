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
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Loan Stock Admin</title>
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/header.css">
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"></script>
<script>
$(function() {

});

function searchStock(){
    var strSearch       = document.forms[0].txtSearch.value;
    var strDepartment   = document.forms[0].cboDepartment.value;
    var strUser         = document.forms[0].cboUser.value;
    var strYear         = document.forms[0].cboYear.value;
    var strMonth        = document.forms[0].cboMonth.value;
    var strSort         = document.forms[0].cboSort.value;
    document.location.href = 'loan_admin.asp?type=search&txtSearch=' + strSearch + '&department=' + strDepartment + '&user=' + strUser + '&year=' + strYear + '&month=' + strMonth + '&sort=' + strSort;
}
    
function resetSearch(){
    document.location.href = 'loan_admin.asp?type=reset';
}
</script>
</head>
<body>
<%
session.lcid = 2057

sub setSearch
    Select case trim(request("type"))
        case "reset"
            session("loan_admin_search")        = ""
            session("loan_admin_department")    = ""
            session("loan_admin_user")          = ""
            session("loan_admin_year")          = ""
            session("loan_admin_month")         = ""
            session("loan_admin_sort")          = ""
            session("loan_admin_initial_page")  = 1
        case "search"
            session("loan_admin_search")        = trim(request("txtSearch"))
            session("loan_admin_department")    = trim(request("department"))
            session("loan_admin_user")          = trim(request("user"))
            session("loan_admin_year")          = trim(request("year"))
            session("loan_admin_month")         = trim(request("month"))
            session("loan_admin_sort")          = trim(request("sort"))
            session("loan_admin_initial_page")  = 1
    end Select
end sub

sub displayLoanStock
    dim iRecordCount
    iRecordCount = 0
    dim strDays
    dim strSQL

    dim intTotalLIC
    dim intTotalQty

    intTotalLIC = 0
    intTotalQty = 0

    dim intRecordCount

    dim strTodayDate
    strTodayDate = FormatDateTime(Date())

    'strSearchTxt = trim(Request("txtSearch"))

    call OpenBaseDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 5000

    if session("loan_admin_qty") = "" then
        session("loan_admin_qty") = "1"
    end if

    if session("loan_admin_sort") = "" then
        session("loan_admin_sort") = "oldest"
    end if

    strSQL = strSQL & "SELECT B9JUNO AS order_no, B9JUGY AS order_line, "
    strSQL = strSQL & " B9SKNO AS ship_number, B9SKGY AS ship_line, B9URKC AS account_code, B9SCSS AS warehouse, "
    strSQL = strSQL & " B9GREG AS item_group, B9SOSC AS item_code, B9SKSU - B9AHEN AS loan_qty, B9HSRC, Y1KOM1 as account_name, "
    strSQL = strSQL & "	B9SKJY, "
    strSQL = strSQL & "	B9SKJM, "
    strSQL = strSQL & "	B9SKJD, "
    strSQL = strSQL & " RIGHT('0' || B9SKJD,2)|| '/' || RIGHT ('0' || B9SKJM,2) || '/' || B9SKJY as loan_date, "
    strSQL = strSQL & " case "
    strSQL = strSQL & " 	when B9STJN <> '00' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) "
    strSQL = strSQL & " 		else 0 "
    strSQL = strSQL & " 	end AS lic, B9SIBN AS serial_number, "
    strSQL = strSQL & " B9ASFN AS comment, Y1REGN as department "
    strSQL = strSQL & " FROM BF9EP "
    strSQL = strSQL & " INNER JOIN EF2SP ON B9SOSC = E2SOSC "
    strSQL = strSQL & " INNER JOIN YF1MP ON CONCAT(B9URKC,B9JURC) = Y1KOKC "
    strSQL = strSQL & " WHERE Y1SKKI <> 'D' AND B9AHEN < B9SKSU "
    strSQL = strSQL & "				AND Y1REGN LIKE '%" & trim(session("loan_admin_department")) & "%' "
    strSQL = strSQL & "				AND (B9SOSC LIKE '%" & UCASE(trim(session("loan_admin_search"))) & "%' "
    strSQL = strSQL & "					OR B9JUNO LIKE '%" & UCASE(trim(session("loan_admin_search"))) & "%' "
    strSQL = strSQL & "					OR B9SIBN LIKE '%" & UCASE(trim(session("loan_admin_search"))) & "%' "
    strSQL = strSQL & "					OR B9SKNO LIKE '%" & UCASE(trim(session("loan_admin_search"))) & "%') "
    strSQL = strSQL & "				AND B9URKC LIKE '%" & UCASE(trim(session("loan_admin_user"))) & "%' "
    strSQL = strSQL & "				AND B9SKJY LIKE '%" & trim(session("loan_admin_year")) & "%' "
    strSQL = strSQL & "				AND B9SKJM LIKE '%" & trim(session("loan_admin_month")) & "%' "
    strSQL = strSQL & " AND (E2NGTY = "
    strSQL = strSQL & "	 (SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only) "
    strSQL = strSQL & "	AND E2NGTM = "
    strSQL = strSQL & "	 (SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only))"
    strSQL = strSQL & "		ORDER BY "

    select case session("loan_admin_sort")
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
    end select

    'response.write strSQL

    rs.Open strSQL, conn

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        rs.AbsolutePage = session("loan_admin_initial_page")

        For intRecord = 1 To rs.PageSize 
            strDays = DateDiff("d",rs("loan_date"), strTodayDate)

            intTotalLIC = intTotalLIC + Cdbl(rs("lic"))
            'intTotalLIC = intTotalLIC + rs("lic")
            intTotalQty = intTotalQty + Cint(rs("loan_qty"))
            'intTotalQty = intTotalQty + CINT(rs("loan_qty"))

            dim strFirstExpiryDate
            strFirstExpiryDate = DateAdd("m", 3, rs("loan_date"))

            dim strFinalExpiryDate
            strFinalExpiryDate = DateAdd("m", 6, rs("loan_date"))

            if DateDiff("d",strFirstExpiryDate, strTodayDate) > 0 then
                strDisplayList = strDisplayList & "<tr class=""info"">"
            else
                strDisplayList = strDisplayList & "<tr>"
            end if

            strDisplayList = strDisplayList & "<td class=""text-center"" nowrap><a href=""view_loan.asp?order=" & Trim(rs("order_no")) & "&line=" & Trim(rs("order_line")) & """><img src=""images/icon_view.png"" border=""0""></a></td>"
            strDisplayList = strDisplayList & "<td>" & rs("account_code") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("account_name") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("department") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("item_code") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("serial_number") & "</td>"
            strDisplayList = strDisplayList & "<td class=""text-center"">" & rs("loan_qty") & "</td>"
            strDisplayList = strDisplayList & "<td class=""text-right"">" & FormatNumber(rs("lic")) & "</td>"
            strDisplayList = strDisplayList & "<td class=""text-right"">" & FormatDateTime(rs("loan_date"),1) & "</td>"
            strDisplayList = strDisplayList & "<td class=""text-center"" nowrap>"
            if DateDiff("d",strFirstExpiryDate, strTodayDate) > 0 then
                strDisplayList = strDisplayList & " <span style=""color:red"">"
            end if
            strDisplayList = strDisplayList & FormatNumber(strDays,0) & "</span></td>"
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next

    else
        strDisplayList = "<tr><td colspan=""10"">No stocks found.</td></tr>"
    end if
    strDisplayList = strDisplayList & "<tr>"
    strDisplayList = strDisplayList & "<td colspan=""10"">"
    strDisplayList = strDisplayList & "<h3>Total Value: <u>$" & FormatNumber(intTotalLIC) & "</u></h3>"
    strDisplayList = strDisplayList & "<h3>Total Units: <u>" & intTotalQty & "</u></h3>"
    strDisplayList = strDisplayList & "<h3>Search results: <u>" & intRecordCount & "</u></h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseBaseDataBase()
end sub

sub main
    call getEmployeeDetails(session("logged_username"))
    call setSearch 

    if trim(session("loan_admin_initial_page"))  = "" then
        session("loan_admin_initial_page") = 1
    end if

    call displayLoanStock
end sub

call main

dim strDisplayList
%>

    <div class="blog-masthead">
        <div class="container">
            <nav class="blog-nav">
                <a class="blog-nav-item" href="loan_summary.asp"><i class="fa fa-home fa-lg"></i></a>
                <a class="blog-nav-item" href="loan-transfer.asp">Transfer</a>
                <a class="blog-nav-item" href="loan-sale.asp">Sale</a>
                <a class="blog-nag-item active">Admin</a>
            </nav>
        </div>
    </div>

    <div class="container">
        <h1 class="page-header"><i class="fa fa-cogs"></i> Loan Stock Admin</h1>

        <p>
            <a class="btn btn-success" href="export_loan_admin.asp?search=<%= request("txtSearch") %>&department=<%= session("loan_admin_department") %>&user=<%= session("loan_admin_user") %>&year=<%= session("loan_admin_year") %>&month=<%= session("loan_admin_month") %>&sort=<%= session("loan_admin_sort") %>"><i class="fa fa-download"></i> Export Excel</a>
        </p>

        <form name="frmSearch" id="frmSearch" method="post" action="loan_admin.asp?type=search" onsubmit="searchStock()" class="form-inline">
            <div class="form-group">
                <input type="text" name="txtSearch" maxlength="15" size="20" class="form-control" value="<%= request("txtSearch") %>" />
            </div>
            <div class="form-group">
                <select name="cboDepartment" class="form-control" onchange="searchStock()">
                  <option value="">All Dept</option>
                  <option <% if session("loan_admin_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("loan_admin_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                  <option <% if session("loan_admin_department") = "F" then Response.Write " selected" end if%> value="F">O&amp;F</option>
                  <option <% if session("loan_admin_department") = "YMEC" then Response.Write " selected" end if%> value="YMEC">YMEC</option>
                </select>
            </div>
            <div class="form-group">
                <select name="cboUser" class="form-control" onchange="searchStock()">
                  <option <% if session("loan_admin_user") = "" then Response.Write " selected" end if%> value="">All Users</option>
                  <option <% if session("loan_admin_user") = "7CAM01" then Response.Write " selected" end if%> value="7CAM01">Cameron Tait</option>
                  <option <% if session("loan_admin_user") = "7CH001" then Response.Write " selected" end if%> value="7CH001">Chris Herring</option>
                  <option <% if session("loan_admin_user") = "7DM001" then Response.Write " selected" end if%> value="7DM001">Dale Moore</option>
                  <option <% if session("loan_admin_user") = "7DMRS0" then Response.Write " selected" end if%> value="7DMRS0">Dale Moore (RS)</option>
                  <option <% if session("loan_admin_user") = "7DH000" then Response.Write " selected" end if%> value="7DH000">Damien Henderson</option>
                  <option <% if session("loan_admin_user") = "7DH002" then Response.Write " selected" end if%> value="7DH002">Damien Henderson (RS)</option>
                  <option <% if session("loan_admin_user") = "7DT001" then Response.Write " selected" end if%> value="7DT001">Dave Thwaites</option>
                  <option <% if session("loan_admin_user") = "7FED00" then Response.Write " selected" end if%> value="7FED00">Felix Elliot-Dedman</option>
                  <option <% if session("loan_admin_user") = "7EM001" then Response.Write " selected" end if%> value="7EM001">Euan McInnes</option>
                  <option <% if session("loan_admin_user") = "7EM002" then Response.Write " selected" end if%> value="7EM002">Euan McInnes (VOX Pedal)</option>
                  <option <% if session("loan_admin_user") = "7GN001" then Response.Write " selected" end if%> value="7GN001">George Nasr</option>
                  <option <% if session("loan_admin_user") = "7GL000" then Response.Write " selected" end if%> value="7GL000">Grant Lane</option>
                  <option <% if session("loan_admin_user") = "7JW001" then Response.Write " selected" end if%> value="7JW001">Jaclyn Williams</option>
                  <option <% if session("loan_admin_user") = "7JG000" then Response.Write " selected" end if%> value="7JG000">Jamie Goff</option>
                  <option <% if session("loan_admin_user") = "7AUDW0" then Response.Write " selected" end if%> value="7AUDW0">Jamie Goff (AUDW)</option>
                  <option <% if session("loan_admin_user") = "7JS001" then Response.Write " selected" end if%> value="7JS001">John Saccaro</option>
                  <option <% if session("loan_admin_user") = "7JP001" then Response.Write " selected" end if%> value="7JP001">Joseph Pantalleresco</option>
                  <option <% if session("loan_admin_user") = "7JD001" then Response.Write " selected" end if%> value="7JD001">Justin D'offay</option>
                  <option <% if session("loan_admin_user") = "7JD002" then Response.Write " selected" end if%> value="7JD002">Justin D'offay (RS)</option>
                  <option <% if session("loan_admin_user") = "7KJ001" then Response.Write " selected" end if%> value="7KJ001">Kevin Johnson</option>
                  <option <% if session("loan_admin_user") = "7LB000" then Response.Write " selected" end if%> value="7LB000">Leon Blaher</option>
                  <option <% if session("loan_admin_user") = "7MC001" then Response.Write " selected" end if%> value="7MC001">Mark Condon</option>
                  <option <% if session("loan_admin_user") = "7ML001" then Response.Write " selected" end if%> value="7ML001">Mark Loey</option>
                  <option <% if session("loan_admin_user") = "7ML002" then Response.Write " selected" end if%> value="7ML002">Mark Lapthorne</option>
                  <option <% if session("loan_admin_user") = "7MT000" then Response.Write " selected" end if%> value="7MT000">Mathew Taylor</option>
                  <option <% if session("loan_admin_user") = "7MH000" then Response.Write " selected" end if%> value="7MH000">Mick Hughes</option>
                  <option <% if session("loan_admin_user") = "7NB001" then Response.Write " selected" end if%> value="7NB001">Nathan Biggin</option>
                  <option <% if session("loan_admin_user") = "7PW000" then Response.Write " selected" end if%> value="7PW000">Paul Wheeler</option>
                  <option <% if session("loan_admin_user") = "7BP000" then Response.Write " selected" end if%> value="7BP000">Peter Beveridge</option>
                  <option <% if session("loan_admin_user") = "7RW001" then Response.Write " selected" end if%> value="7RW001">Russell Wykes</option>
                  <option <% if session("loan_admin_user") = "7SG001" then Response.Write " selected" end if%> value="7SG001">Simon Goldsworthy</option>
                  <option <% if session("loan_admin_user") = "7SL001" then Response.Write " selected" end if%> value="7SL001">Steve Legg</option>
                  <option <% if session("loan_admin_user") = "7SVR01" then Response.Write " selected" end if%> value="7SVR01">Steven Vranch</option>
                  <option <% if session("loan_admin_user") = "7TM000" then Response.Write " selected" end if%> value="7TM000">Terry McMahon</option>
                  <option <% if session("loan_admin_user") = "7SMC01" then Response.Write " selected" end if%> value="7SMC01">Shaun McMahon</option>
                  <option <% if session("loan_admin_user") = "7WF001" then Response.Write " selected" end if%> value="7WF001">Wesley Fischer</option>
                  <option <% if session("loan_admin_user") = "7YME01" then Response.Write " selected" end if%> value="7YME01">YMEC Altona</option>
                  <option <% if session("loan_admin_user") = "7YME02" then Response.Write " selected" end if%> value="7YME02">YMEC Balwyn</option>
                  <option <% if session("loan_admin_user") = "7YME09" then Response.Write " selected" end if%> value="7YME09">YMEC Baulkham Hills</option>
                  <option <% if session("loan_admin_user") = "7YME05" then Response.Write " selected" end if%> value="7YME05">YMEC Carnegie</option>
                  <option <% if session("loan_admin_user") = "7YME12" then Response.Write " selected" end if%> value="7YME12">YMEC Morley</option>
                  <option <% if session("loan_admin_user") = "7YME08" then Response.Write " selected" end if%> value="7YME08">YMEC Strathmore</option>
                </select>
            </div>
            <div class="form-group">
                <select name="cboYear" class="form-control" onchange="searchStock()">
                  <option <% if session("loan_admin_year") = "" then Response.Write " selected" end if%> value="">All years</option>
                  <option <% if session("loan_admin_year") = "2016" then Response.Write " selected" end if%> value="2016">2016 only</option>
                  <option <% if session("loan_admin_year") = "2015" then Response.Write " selected" end if%> value="2015">2015 only</option>
                  <option <% if session("loan_admin_year") = "2014" then Response.Write " selected" end if%> value="2014">2014 only</option>
                  <option <% if session("loan_admin_year") = "2013" then Response.Write " selected" end if%> value="2013">2013 only</option>
                  <option <% if session("loan_admin_year") = "2012" then Response.Write " selected" end if%> value="2012">2012 only</option>
                  <option <% if session("loan_admin_year") = "2011" then Response.Write " selected" end if%> value="2011">2011 only</option>
                  <option <% if session("loan_admin_year") = "2010" then Response.Write " selected" end if%> value="2010">2010 only</option>
                  <option <% if session("loan_admin_year") = "2009" then Response.Write " selected" end if%> value="2009">2009 only</option>
                  <option <% if session("loan_admin_year") = "2008" then Response.Write " selected" end if%> value="2008">2008 only</option>
                </select>
            </div>
            <div class="form-group">
                <select name="cboMonth" class="form-control" onchange="searchStock()">
                  <option <% if session("loan_admin_month") = "" then Response.Write " selected" end if%> value="">All months</option>
                  <option <% if session("loan_admin_month") = "1" then Response.Write " selected" end if%> value="1">January</option>
                  <option <% if session("loan_admin_month") = "2" then Response.Write " selected" end if%> value="2">February</option>
                  <option <% if session("loan_admin_month") = "3" then Response.Write " selected" end if%> value="3">March</option>
                  <option <% if session("loan_admin_month") = "4" then Response.Write " selected" end if%> value="4">April</option>
                  <option <% if session("loan_admin_month") = "5" then Response.Write " selected" end if%> value="5">May</option>
                  <option <% if session("loan_admin_month") = "6" then Response.Write " selected" end if%> value="6">June</option>
                  <option <% if session("loan_admin_month") = "7" then Response.Write " selected" end if%> value="7">July</option>
                  <option <% if session("loan_admin_month") = "8" then Response.Write " selected" end if%> value="8">August</option>
                  <option <% if session("loan_admin_month") = "9" then Response.Write " selected" end if%> value="9">September</option>
                  <option <% if session("loan_admin_month") = "10" then Response.Write " selected" end if%> value="10">October</option>
                  <option <% if session("loan_admin_month") = "11" then Response.Write " selected" end if%> value="11">November</option>
                  <option <% if session("loan_admin_month") = "12" then Response.Write " selected" end if%> value="12">December</option>
                </select>
            </div>
            <div class="form-group">
                <select name="cboSort" class="form-control" onchange="searchStock()">
                  <option <% if session("loan_admin_sort") = "oldest" then Response.Write " selected" end if%> value="oldest">Sort: Old - New</option>
                  <option <% if session("loan_admin_sort") = "latest" then Response.Write " selected" end if%> value="latest">Sort: New - Old</option>
                  <option <% if session("loan_admin_sort") = "product" then Response.Write " selected" end if%> value="product">Sort: Item Code (A-Z)</option>
                  <option <% if session("loan_admin_sort") = "expensive" then Response.Write " selected" end if%> value="expensive">Sort: Value (High-Low)</option>
                  <option <% if session("loan_admin_sort") = "cheapest" then Response.Write " selected" end if%> value="cheapest">Sort: Value (Low-High)</option>
                  <option <% if session("loan_admin_sort") = "serial" then Response.Write " selected" end if%> value="serial">Sort: Serial</option>
                </select>
            </div>
            <input type="button" name="btnSearch" value="Search" class="btn btn-primary" onclick="searchStock()" />
            <input type="button" name="btnReset" value="Reset" class="btn btn-default" onclick="resetSearch()" />
        </form>
    </div>
    <br>
    <div class="table-responsive">
        <table class="table table-striped table-hover loan_table">
            <thead>
                <tr>
                    <th width="5%"></th>
                    <th width="10%">Account</th>
                    <th width="15%">Account Name</th>
                    <th width="5%">Dept</th>
                    <th width="10%">Item Code</th>
                    <th width="10%">Serial</th>
                    <th class="text-center" width="5%">Qty</th>
                    <th class="text-right" width="10%">LIC $</th>
                    <th class="text-right" width="15%">Loan Date</th>
                    <th class="text-center" width="15%">Day Count</th>
                </tr>
            </thead>
            <tbody>
                <%= strDisplayList %>
            </tbody>
        </table>
    </div>
    <p><small>You are logged in as: <%= session("logged_username") %> (<%= UCASE(trim(session("emp_initial"))) %>)</small></p>
</body>
</html>