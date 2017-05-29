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
<!--#include file="include/loan_functions.asp" -->
<!--#include file="../include/connection_base.asp" -->
<!--#include file="class/clsEmployee.asp" -->
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
<title>Outstanding Invoices</title>
<link rel="stylesheet" href="css/style.css">
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/header.css">
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
<script>
function searchOutstanding() {
    var strSearch       = document.forms[0].txtSearch.value;
    var intStartDate    = document.forms[0].cboStartDate.value;
    var intStartMonth   = document.forms[0].cboStartMonth.value;
    var intStartYear    = document.forms[0].cboStartYear.value;
    var intEndDate      = document.forms[0].cboEndDate.value;
    var intEndMonth     = document.forms[0].cboEndMonth.value;
    var intEndYear      = document.forms[0].cboEndYear.value;

    document.location.href = 'outstanding.asp?type=search&txtSearch=' + strSearch + '&cboStartDate=' + intStartDate + '&cboStartMonth=' + intStartMonth + '&cboStartYear=' + intStartYear + '&cboEndDate=' + intEndDate + '&cboEndMonth=' + intEndMonth + '&cboEndYear=' + intEndYear;
}

function resetSearch() {
    document.location.href = 'outstanding.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
    select case Trim(Request("type"))
        case "reset"
            session("outstanding_search")       = ""
            session("outstanding_start_date")   = ""
            session("outstanding_start_month")  = ""
            session("outstanding_start_year")   = ""
            session("outstanding_end_date")     = ""
            session("outstanding_end_month")    = ""
            session("outstanding_end_year")     = ""
            session("outstanding_initial_page") = 1
        case "search"
            session("outstanding_search")       = Trim(Request("txtSearch"))
            session("outstanding_start_date")   = Trim(Request("cboStartDate"))
            session("outstanding_start_month")  = Trim(Request("cboStartMonth"))
            session("outstanding_start_year")   = Trim(Request("cboStartYear"))
            session("outstanding_end_date")     = Trim(Request("cboEndDate"))
            session("outstanding_end_month")    = Trim(Request("cboEndMonth"))
            session("outstanding_end_year")     = Trim(Request("cboEndYear"))
            session("outstanding_initial_page") = 1
    end select
end sub

sub displayOutstandingInvoices
    dim iRecordCount
    iRecordCount = 0
    dim strDays
    dim strSQL
    dim intRecordCount

    dim intTotalCredit
    intTotalCredit = 0
    dim intTotalDebit
    intTotalDebit = 0

    call OpenBaseDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 5000

    if session("outstanding_start_year") = "" then
        session("outstanding_start_year") = "2015"
    end if

    if session("outstanding_end_year") = "" then
        session("outstanding_end_year") = "2016"
    end if

    strSQL = strSQL & "SELECT G1SKYC AS DealerCode, Y1KOM1 AS DealerName, B.G1UKNO AS InvoiceNo, G1UKGN AS Line, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD, B.G1TSYK AS CreditFlag,"
    strSQL = strSQL & " CONCAT(CONCAT(CONCAT(CONCAT(G1SHKD,'/'),G1SHKM),'/'),G1SHKY) AS DueDate,"
    strSQL = strSQL & " SUM(G1MKZB) AS Amount, B6AHNO "
    strSQL = strSQL & " FROM GF1EP B INNER JOIN YF1MP ON Y1KOKC = G1SKYC "
	strSQL = strSQL & " INNER JOIN (SELECT DISTINCT G1UKNO, CASE WHEN G1TSYK = 1 THEN '' ELSE B6AHNO END AS B6AHNO, G1TSYK FROM GF1EP LEFT JOIN BF6EP ON G1UKNO = B6INNO where b6ingy <> 999 GROUP BY G1UKNO, G1UKGN, B6AHNO, G1TSYK) A on B.G1UKNO = A.G1UKNO AND B.G1TSYK = A.G1TSYK "
    strSQL = strSQL & " WHERE G1SKKI <> 'D' AND Y1SKKI <> 'D' "
    strSQL = strSQL & " AND G1SKYC = '" & Ucase(session("outstanding_search")) & "' "
    strSQL = strSQL & " AND G1KSNO = 0 AND G1SHKY <> 0 "
    'strSQL = strSQL & " AND G1SHKY * 10000 + G1SHKM * 100 + G1SHKD BETWEEN 20150101 AND 20150630 "
    strSQL = strSQL & " AND G1SHKY * 10000 + G1SHKM * 100 + G1SHKD BETWEEN " & session("outstanding_start_year") & session("outstanding_start_month") & session("outstanding_start_date") & " AND " & session("outstanding_end_year") & session("outstanding_end_month") & session("outstanding_end_date") & " "
    strSQL = strSQL & " GROUP BY G1SKYC, Y1KOM1, B.G1UKNO, G1UKGN, B.G1TSYK, Y1YGKG, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD, B6AHNO, B.G1TSYK "
    strSQL = strSQL & " ORDER BY G1SHKY, G1SHKM, G1SHKD"

    'response.write strSQL

    rs.Open strSQL, conn

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        rs.AbsolutePage = session("outstanding_initial_page")

        For intRecord = 1 To rs.PageSize
            strDisplayList = strDisplayList & "<tr>"
            strDisplayList = strDisplayList & "<td>" & rs("DealerCode") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("DealerName") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("B6AHNO") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("InvoiceNo") & ""
            if len(trim(rs("Line"))) > 1 then
                strDisplayList = strDisplayList & "-" & rs("Line") & ""
            end if
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("G1SEKD") & " "
            Select Case trim(rs("G1SEKM"))
                case "1"
                    strDisplayList = strDisplayList & "Jan"
                case "2"
                    strDisplayList = strDisplayList & "Feb"
                case "3"
                    strDisplayList = strDisplayList & "Mar"
                case "4"
                    strDisplayList = strDisplayList & "Apr"
                case "5"
                    strDisplayList = strDisplayList & "May"
                case "6"
                    strDisplayList = strDisplayList & "Jun"
                case "7"
                    strDisplayList = strDisplayList & "Jul"
                case "8"
                    strDisplayList = strDisplayList & "Aug"
                case "9"
                    strDisplayList = strDisplayList & "Sep"
                case "10"
                    strDisplayList = strDisplayList & "Oct"
                case "11"
                    strDisplayList = strDisplayList & "Nov"
                case "12"
                    strDisplayList = strDisplayList & "Dec"
            end select
            strDisplayList = strDisplayList & " " & rs("G1SEKY") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("G1SHKD") & " "
            Select Case trim(rs("G1SHKM"))
                case "1"
                    strDisplayList = strDisplayList & "Jan"
                case "2"
                    strDisplayList = strDisplayList & "Feb"
                case "3"
                    strDisplayList = strDisplayList & "Mar"
                case "4"
                    strDisplayList = strDisplayList & "Apr"
                case "5"
                    strDisplayList = strDisplayList & "May"
                case "6"
                    strDisplayList = strDisplayList & "Jun"
                case "7"
                    strDisplayList = strDisplayList & "Jul"
                case "8"
                    strDisplayList = strDisplayList & "Aug"
                case "9"
                    strDisplayList = strDisplayList & "Sep"
                case "10"
                    strDisplayList = strDisplayList & "Oct"
                case "11"
                    strDisplayList = strDisplayList & "Nov"
                case "12"
                    strDisplayList = strDisplayList & "Dec"
            end select
            strDisplayList = strDisplayList & " " & rs("G1SHKY") & "</td>"
            'strDisplayList = strDisplayList & " " & rs("G1SHKY") & " (" & rs("DueDate") & ")</td>"
            if Trim(rs("CreditFlag")) = "1" then
                'strDisplayList = strDisplayList & "<td>" & rs("Amount") & "</td>"
                strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("Amount"),2) & "</td>"
                intTotalCredit = intTotalCredit + CCur(rs("Amount"))
            else
                strDisplayList = strDisplayList & "<td></td>"
            end if

            if Trim(rs("CreditFlag")) = "0" then
                'strDisplayList = strDisplayList & "<td>" & rs("Amount") & "</td>"
                strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("Amount"),2) & "</td>"
                intTotalDebit = intTotalDebit + CCur(rs("Amount"))
            else
                strDisplayList = strDisplayList & "<td></td>"
            end if
            'strDisplayList = strDisplayList & "<td>" & rs("CreditFlag") & "</td>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next
    else
        strDisplayList = "<tr><td colspan=""10"">No invoices found.</td></tr>"
    end if
    strDisplayList = strDisplayList & "<tr>"
    strDisplayList = strDisplayList & "<td colspan=""6""></td>"
    strDisplayList = strDisplayList & "<td><h4>" & FormatNumber(intTotalCredit) & "</h4></td>"
    strDisplayList = strDisplayList & "<td><h4>" & FormatNumber(intTotalDebit) & "</h4></td>"
    strDisplayList = strDisplayList & "</tr>"
    strDisplayList = strDisplayList & "<tr>"
    strDisplayList = strDisplayList & "<td colspan=""10"" align=""center"">"
    strDisplayList = strDisplayList & "<h3>Total: " & intRecordCount & " invoices</h3>"	
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseBaseDataBase()
end sub

sub main
    call setSearch

    if trim(session("outstanding_initial_page"))  = "" then
        session("outstanding_initial_page") = 1
    end if

    call displayOutstandingInvoices
end sub

call main

dim strDisplayList
%>
<div class="container">
  <h1 class="page-header"><i class="fa fa-usd"></i> Outstanding Invoices</h1>
  <form name="frmSearch" id="frmSearch" action="outstanding.asp?type=search" method="post" onsubmit="searchOutstanding()">
    <div class="float_left">
      <input type="text" class="form-control" name="txtSearch" value="<%= request("txtSearch") %>" maxlength="9" placeholder="9-character Dealer Code" />
      <strong>Dealer Code</strong> E.g. 1ME001000</div>
    <div class="float_left">
      <select name="cboStartDate" class="form-control" onchange="searchOutstanding()">
        <option <% if session("outstanding_start_date") = "01" then Response.Write " selected" end if%> value="01">1</option>
        <option <% if session("outstanding_start_date") = "02" then Response.Write " selected" end if%> value="02">2</option>
        <option <% if session("outstanding_start_date") = "03" then Response.Write " selected" end if%> value="03">3</option>
        <option <% if session("outstanding_start_date") = "04" then Response.Write " selected" end if%> value="04">4</option>
        <option <% if session("outstanding_start_date") = "05" then Response.Write " selected" end if%> value="05">5</option>
        <option <% if session("outstanding_start_date") = "06" then Response.Write " selected" end if%> value="06">6</option>
        <option <% if session("outstanding_start_date") = "07" then Response.Write " selected" end if%> value="07">7</option>
        <option <% if session("outstanding_start_date") = "08" then Response.Write " selected" end if%> value="08">8</option>
        <option <% if session("outstanding_start_date") = "09" then Response.Write " selected" end if%> value="09">9</option>
        <option <% if session("outstanding_start_date") = "10" then Response.Write " selected" end if%> value="10">10</option>
        <option <% if session("outstanding_start_date") = "11" then Response.Write " selected" end if%> value="11">11</option>
        <option <% if session("outstanding_start_date") = "12" then Response.Write " selected" end if%> value="12">12</option>
        <option <% if session("outstanding_start_date") = "13" then Response.Write " selected" end if%> value="13">13</option>
        <option <% if session("outstanding_start_date") = "14" then Response.Write " selected" end if%> value="14">14</option>
        <option <% if session("outstanding_start_date") = "15" then Response.Write " selected" end if%> value="15">15</option>
        <option <% if session("outstanding_start_date") = "16" then Response.Write " selected" end if%> value="16">16</option>
        <option <% if session("outstanding_start_date") = "17" then Response.Write " selected" end if%> value="17">17</option>
        <option <% if session("outstanding_start_date") = "18" then Response.Write " selected" end if%> value="18">18</option>
        <option <% if session("outstanding_start_date") = "19" then Response.Write " selected" end if%> value="19">19</option>
        <option <% if session("outstanding_start_date") = "20" then Response.Write " selected" end if%> value="20">20</option>
        <option <% if session("outstanding_start_date") = "21" then Response.Write " selected" end if%> value="21">21</option>
        <option <% if session("outstanding_start_date") = "22" then Response.Write " selected" end if%> value="22">22</option>
        <option <% if session("outstanding_start_date") = "23" then Response.Write " selected" end if%> value="23">23</option>
        <option <% if session("outstanding_start_date") = "24" then Response.Write " selected" end if%> value="24">24</option>
        <option <% if session("outstanding_start_date") = "25" then Response.Write " selected" end if%> value="25">25</option>
        <option <% if session("outstanding_start_date") = "26" then Response.Write " selected" end if%> value="26">26</option>
        <option <% if session("outstanding_start_date") = "27" then Response.Write " selected" end if%> value="27">27</option>
        <option <% if session("outstanding_start_date") = "28" then Response.Write " selected" end if%> value="28">28</option>
        <option <% if session("outstanding_start_date") = "29" then Response.Write " selected" end if%> value="29">29</option>
        <option <% if session("outstanding_start_date") = "30" then Response.Write " selected" end if%> value="30">30</option>
        <option <% if session("outstanding_start_date") = "31" then Response.Write " selected" end if%> value="31">31</option>
      </select><strong>Start Date</strong>
    </div>
    <div class="float_left">
      <select name="cboStartMonth" class="form-control" onchange="searchOutstanding()">
        <option <% if session("outstanding_start_month") = "01" then Response.Write " selected" end if%> value="01">Jan</option>
        <option <% if session("outstanding_start_month") = "02" then Response.Write " selected" end if%> value="02">Feb</option>
        <option <% if session("outstanding_start_month") = "03" then Response.Write " selected" end if%> value="03">Mar</option>
        <option <% if session("outstanding_start_month") = "04" then Response.Write " selected" end if%> value="04">Apr</option>
        <option <% if session("outstanding_start_month") = "05" then Response.Write " selected" end if%> value="05">May</option>
        <option <% if session("outstanding_start_month") = "06" then Response.Write " selected" end if%> value="06">Jun</option>
        <option <% if session("outstanding_start_month") = "07" then Response.Write " selected" end if%> value="07">Jul</option>
        <option <% if session("outstanding_start_month") = "08" then Response.Write " selected" end if%> value="08">Aug</option>
        <option <% if session("outstanding_start_month") = "09" then Response.Write " selected" end if%> value="09">Sep</option>
        <option <% if session("outstanding_start_month") = "10" then Response.Write " selected" end if%> value="10">Oct</option>
        <option <% if session("outstanding_start_month") = "11" then Response.Write " selected" end if%> value="11">Nov</option>
        <option <% if session("outstanding_start_month") = "12" then Response.Write " selected" end if%> value="12">Dec</option>
      </select>
    </div>
    <div class="float_left">
      <select name="cboStartYear" class="form-control" onchange="searchOutstanding()">
        <option <% if session("outstanding_start_year") = "2015" then Response.Write " selected" end if%> value="2015">2015</option>
        <option <% if session("outstanding_start_year") = "2014" then Response.Write " selected" end if%> value="2014">2014</option>
      </select>
    </div>
    <div class="float_left">
      <select name="cboEndDate" class="form-control" onchange="searchOutstanding()">
        <option <% if session("outstanding_end_date") = "01" then Response.Write " selected" end if%> value="01">1</option>
        <option <% if session("outstanding_end_date") = "02" then Response.Write " selected" end if%> value="02">2</option>
        <option <% if session("outstanding_end_date") = "03" then Response.Write " selected" end if%> value="03">3</option>
        <option <% if session("outstanding_end_date") = "04" then Response.Write " selected" end if%> value="04">4</option>
        <option <% if session("outstanding_end_date") = "05" then Response.Write " selected" end if%> value="05">5</option>
        <option <% if session("outstanding_end_date") = "06" then Response.Write " selected" end if%> value="06">6</option>
        <option <% if session("outstanding_end_date") = "07" then Response.Write " selected" end if%> value="07">7</option>
        <option <% if session("outstanding_end_date") = "08" then Response.Write " selected" end if%> value="08">8</option>
        <option <% if session("outstanding_end_date") = "09" then Response.Write " selected" end if%> value="09">9</option>
        <option <% if session("outstanding_end_date") = "10" then Response.Write " selected" end if%> value="10">10</option>
        <option <% if session("outstanding_end_date") = "11" then Response.Write " selected" end if%> value="11">11</option>
        <option <% if session("outstanding_end_date") = "12" then Response.Write " selected" end if%> value="12">12</option>
        <option <% if session("outstanding_end_date") = "13" then Response.Write " selected" end if%> value="13">13</option>
        <option <% if session("outstanding_end_date") = "14" then Response.Write " selected" end if%> value="14">14</option>
        <option <% if session("outstanding_end_date") = "15" then Response.Write " selected" end if%> value="15">15</option>
        <option <% if session("outstanding_end_date") = "16" then Response.Write " selected" end if%> value="16">16</option>
        <option <% if session("outstanding_end_date") = "17" then Response.Write " selected" end if%> value="17">17</option>
        <option <% if session("outstanding_end_date") = "18" then Response.Write " selected" end if%> value="18">18</option>
        <option <% if session("outstanding_end_date") = "19" then Response.Write " selected" end if%> value="19">19</option>
        <option <% if session("outstanding_end_date") = "20" then Response.Write " selected" end if%> value="20">20</option>
        <option <% if session("outstanding_end_date") = "21" then Response.Write " selected" end if%> value="21">21</option>
        <option <% if session("outstanding_end_date") = "22" then Response.Write " selected" end if%> value="22">22</option>
        <option <% if session("outstanding_end_date") = "23" then Response.Write " selected" end if%> value="23">23</option>
        <option <% if session("outstanding_end_date") = "24" then Response.Write " selected" end if%> value="24">24</option>
        <option <% if session("outstanding_end_date") = "25" then Response.Write " selected" end if%> value="25">25</option>
        <option <% if session("outstanding_end_date") = "26" then Response.Write " selected" end if%> value="26">26</option>
        <option <% if session("outstanding_end_date") = "27" then Response.Write " selected" end if%> value="27">27</option>
        <option <% if session("outstanding_end_date") = "28" then Response.Write " selected" end if%> value="28">28</option>
        <option <% if session("outstanding_end_date") = "29" then Response.Write " selected" end if%> value="29">29</option>
        <option <% if session("outstanding_end_date") = "30" then Response.Write " selected" end if%> value="30">30</option>
        <option <% if session("outstanding_end_date") = "31" then Response.Write " selected" end if%> value="31">31</option>
      </select><strong>End Date</strong>
    </div>
    <div class="float_left">
      <select name="cboEndMonth" class="form-control" onchange="searchOutstanding()">
        <option <% if session("outstanding_end_month") = "01" then Response.Write " selected" end if%> value="01">Jan</option>
        <option <% if session("outstanding_end_month") = "02" then Response.Write " selected" end if%> value="02">Feb</option>
        <option <% if session("outstanding_end_month") = "03" then Response.Write " selected" end if%> value="03">Mar</option>
        <option <% if session("outstanding_end_month") = "04" then Response.Write " selected" end if%> value="04">Apr</option>
        <option <% if session("outstanding_end_month") = "05" then Response.Write " selected" end if%> value="05">May</option>
        <option <% if session("outstanding_end_month") = "06" then Response.Write " selected" end if%> value="06">Jun</option>
        <option <% if session("outstanding_end_month") = "07" then Response.Write " selected" end if%> value="07">Jul</option>
        <option <% if session("outstanding_end_month") = "08" then Response.Write " selected" end if%> value="08">Aug</option>
        <option <% if session("outstanding_end_month") = "09" then Response.Write " selected" end if%> value="09">Sep</option>
        <option <% if session("outstanding_end_month") = "10" then Response.Write " selected" end if%> value="10">Oct</option>
        <option <% if session("outstanding_end_month") = "11" then Response.Write " selected" end if%> value="11">Nov</option>
        <option <% if session("outstanding_end_month") = "12" then Response.Write " selected" end if%> value="12">Dec</option>
      </select>
    </div>
    <div class="float_left">
      <select name="cboEndYear" class="form-control" onchange="searchOutstanding()">
        <option <% if session("outstanding_end_year") = "2017" then Response.Write " selected" end if%> value="2017">2017</option>
        <option <% if session("outstanding_end_year") = "2016" then Response.Write " selected" end if%> value="2016">2016</option>
        <option <% if session("outstanding_end_year") = "2015" then Response.Write " selected" end if%> value="2015">2015</option>
        <option <% if session("outstanding_end_year") = "2014" then Response.Write " selected" end if%> value="2014">2014</option>
      </select>
    </div>
    <div class="float_left">
      <input type="button" name="btnSearch" value="Search" onclick="searchOutstanding()" class="btn btn-primary" />
      <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" class="btn btn-primary" />
    </div>
  </form>
  <div class="new_line"></div>
  <h2 align="right"><a href="export_outstanding-invoices.asp">Export</a></h2>
  <div class="table-responsive">
    <table class="table table-striped">
      <thead>
        <tr>
          <td>Dealer Code</td>
          <td>Dealer Name</td>
		  <td>Order No</td>
          <td>Invoice</td>
          <td>Invoice Date</td>
          <td>Due Date</td>
          <td>Credit</td>
          <td>Debit</td>
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