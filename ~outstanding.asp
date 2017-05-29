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
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/header.css">
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
<script>
function searchOutstanding(){
    var strSearch 		= document.forms[0].txtSearch.value;	
    document.location.href = 'outstanding.asp?type=search&txtSearch=' + strSearch;
}

function resetSearch(){
	document.location.href = 'outstanding.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch	
	select case Trim(Request("type"))
		case "reset"
			session("outstanding_search") 		= ""			
			session("outstanding_initial_page") 	= 1
		case "search"
			session("outstanding_search") 		= Trim(Request("txtSearch"))			
			session("outstanding_initial_page") 	= 1
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
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 5000
	
	strSQL = strSQL & "SELECT G1SKYC AS DealerCode, Y1KOM1 AS DealerName, G1UKNO AS InvoiceNo, G1UKGN AS Line, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD, G1TSYK AS CreditFlag,"
	'strSQL = strSQL & " SUM(G1MKZB) AS unallocatedAmount, SUM(G1KJKG) AS Amount"	
	strSQL = strSQL & " SUM(G1MKZB) AS Amount"	
	strSQL = strSQL & " FROM GF1EP INNER JOIN YF1MP ON Y1KOKC = G1SKYC WHERE"	
	strSQL = strSQL & " G1SKKI <> 'D' AND Y1SKKI <> 'D' "
	strSQL = strSQL & " AND G1SKYC = '" & Ucase(session("outstanding_search")) & "' "
	strSQL = strSQL & " AND G1KSNO = 0 AND G1SHKY <> 0"
	strSQL = strSQL & " GROUP BY G1SKYC, Y1KOM1, G1UKNO, G1UKGN, G1TSYK, Y1YGKG, G1SEKY, G1SEKM, G1SEKD, G1SHKY, G1SHKM, G1SHKD "
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
			if Trim(rs("CreditFlag")) = "1" then
				strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("Amount")) & "</td>"
				intTotalCredit = intTotalCredit + CCur(rs("Amount"))
			else
				strDisplayList = strDisplayList & "<td></td>"
			end if
			
			if Trim(rs("CreditFlag")) = "0" then
				strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("Amount")) & "</td>"
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
	strDisplayList = strDisplayList & "<td colspan=""5""></td>"
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
  <h1 class="page-header"><i class="fa fa-usd"></i> Outstanding Invoices (BETA)</h1>
  <form name="frmSearch" id="frmSearch" action="outstanding.asp?type=search" method="post" onsubmit="searchOutstanding()">
    <div class="row">
      <div class="form-group col-md-6">
        <input type="text" class="form-control" name="txtSearch" value="<%= request("txtSearch") %>" maxlength="9" placeholder="9-character Dealer Code" /> E.g. 1ME001000
        <br>
        <input type="button" name="btnSearch" value="Search" onclick="searchOutstanding()" class="btn btn-primary" />
        <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" class="btn btn-primary" />
      </div>
      <div class="form-group col-md-6">
        <h2 align="right"><a href="export_outstanding.asp">Export</a></h2>
      </div>
    </div>
  </form>
  <div class="table-responsive">
    <table class="table table-striped">
      <thead>
        <tr>
          <td>Dealer Code</td>
          <td>Dealer Name</td>
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