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
    <title>Loan Stock Summary</title>
    <link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
    <link rel="stylesheet" href="css/header.css">
    <link rel="stylesheet" href="https://use.fontawesome.com/603ac387f6.css">
    <script>
        function searchStock() {
            var strSort = document.forms[0].cboSort.value;
            document.location.href = 'loan_summary.asp?sort=' + strSort;
        }
    </script>
</head>
<body>
<%
sub setSearch
    session("loan_summary_sort") = trim(request("sort"))
    session("loan_summary_initial_page") = 1
end sub

sub displayLoanStock
    dim iRecordCount
    iRecordCount = 0
    dim strDays
    dim strSQL

    dim intRecordCount

    call OpenBaseDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 5000

    if session("loan_summary_sort") = "" then
        session("loan_summary_sort") = "name"
    end if

    strSQL = strSQL & "SELECT TRIM(B9URKC) as account_code, TRIM(Y1KOM1) as account_name, SUM(B9SKSU - B9AHEN) AS total_qty,"
    strSQL = strSQL & "	 SUM(case when B9STJN <> '00' then (B9SKSU - B9AHEN) * (E2IHTN + (E2IHTN * E2KZRT / 100) + (E2IHTN * E2SKKR / 100)) else 0 end) AS total_lic"
    strSQL = strSQL & "	FROM BF9EP "
    strSQL = strSQL & "		INNER JOIN EF2SP ON B9SOSC = E2SOSC "
    strSQL = strSQL & "		INNER JOIN YF1MP ON CONCAT(B9URKC,B9JURC) = Y1KOKC "
    strSQL = strSQL & "	WHERE "
    strSQL = strSQL & "		Y1SKKI <> 'D' "
    strSQL = strSQL & "		AND B9AHEN < B9SKSU "
    strSQL = strSQL & "	 	AND (E2NGTY = "
    strSQL = strSQL & "	(SELECT E2NGTY FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only)"
    strSQL = strSQL & "		AND E2NGTM = "
    strSQL = strSQL & "	(SELECT E2NGTM FROM EF2SP WHERE E2SOSC = B9SOSC AND E2IHTN + (E2IHTN * E2KZRT/100) + (E2IHTN * E2SKKR/100) <> 0 ORDER BY E2NGTY * 100 + E2NGTM DESC Fetch First 1 Row Only))"
    strSQL = strSQL & "		GROUP BY B9URKC, Y1KOM1 "
    strSQL = strSQL & " 		ORDER BY "

    select case session("loan_summary_sort")
        case "account"
            strSQL = strSQL & "1"
        case "name"
            strSQL = strSQL & "2"
        case "qty"
            strSQL = strSQL & "3 DESC"
        case "lic"
            strSQL = strSQL & "4 DESC"
    end select

    'response.write strSQL

    rs.Open strSQL, conn

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        rs.AbsolutePage = session("loan_summary_initial_page")

        For intRecord = 1 To rs.PageSize
            strDisplayList = strDisplayList & "<tr>"
            strDisplayList = strDisplayList & "<td>" & rs("account_code") & "</td>"
            strDisplayList = strDisplayList & "<td><a href=""loan-user.asp?account=" & Trim(rs("account_code")) & "&name=" & rs("account_name") & """>" & rs("account_name") & "</a></td>"
            strDisplayList = strDisplayList & "<td>" & rs("total_qty") & "</td>"
            strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("total_lic")) & "</td>"
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next
    else
        strDisplayList = "<tr><td colspan=""4"">No items found.</td></tr>"
    end if
    strDisplayList = strDisplayList & "<tr>"
    strDisplayList = strDisplayList & "<td colspan=""4"">"
    strDisplayList = strDisplayList & "<h3>Total: <u>" & intRecordCount & "</u> loan accounts</h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseBaseDataBase()
end sub

sub main
    call getEmployeeDetails(session("logged_username"))
    call setSearch

    if trim(session("loan_summary_initial_page")) = "" then
        session("loan_summary_initial_page") = 1
    end if

    call displayLoanStock
end sub

call main

dim strDisplayList
%>
    <div class="blog-masthead">
        <div class="container">
            <nav class="blog-nav">
                <a class="blog-nav-item active"><i class="fa fa-home fa-lg"></i></a>
                <a class="blog-nav-item" href="loan-transfer.asp">Transfer</a>
                <a class="blog-nav-item" href="loan-sale.asp">Sale</a>
            </nav>
        </div>
    </div>
    <div class="container">
        <h1 class="page-header"><i class="fa fa-list"></i> Loan Stock Summary</h1>
        <p align="right">
            <!--<a href="loan_admin.asp"><button type="button" class="btn btn-info">Admin Only</button></a>-->
            <a href="export_loan_summary.asp?sort=<%= session("loan_summary_sort") %>">
                <button type="button" class="btn btn-success"><i class="fa fa-download"></i> Export Summary</button>
            </a>
            <a href="export_loan_all.asp">
                <button type="button" class="btn btn-success"><i class="fa fa-download"></i> Export</button>
            </a>
            <a href="export_loan_all_location.asp">
                <button type="button" class="btn btn-success"><i class="fa fa-download"></i> Export with Locations</button>
            </a>
        </p>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead>
                    <tr>
                        <td>Account Code</td>
                        <td><a href="?sort=name">Account Name</a></td>
                        <td><a href="?sort=qty">Total Qty</a></td>
                        <td><a href="?sort=lic">Total Value $</a></td>
                    </tr>
                </thead>
                <tbody>
                    <%= strDisplayList %>
                </tbody>
            </table>
        </div>
        <p>You are logged in as: <%= session("logged_username") %> (<%= UCASE(trim(session("emp_initial"))) %>)</p>
    </div>
    <script src="//code.jquery.com/jquery.js"></script>
    <script src="bootstrap/js/bootstrap.min.js"></script>
</body>
</html>