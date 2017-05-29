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
<!--#include file="class/clsLoanApproval.asp " -->
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
<title>Loan Stock Renewal</title>
<!-- <link rel="stylesheet" href="include/stylesheet.css" type="text/css" /> -->
<!-- <link rel="stylesheet" href="css/sticky-navigation.css" /> -->
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="css/header.css">
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"></script>
<script>
$(function() {
    // var sticky_navigation_offset_top = $('#sticky_navigation').offset().top;

    // var sticky_navigation = function(){
    //     var scroll_top = $(window).scrollTop();

    //     if (scroll_top > sticky_navigation_offset_top) {
    //         $('#sticky_navigation').css({ 'position': 'fixed', 'top':0, 'left':0 });
    //     } else {
    //         $('#sticky_navigation').css({ 'position': 'relative' });
    //     }
    // };

    // sticky_navigation();

    // $(window).scroll(function() {
    //     sticky_navigation();
    // });

    // $('a[href="#"]').click(function(event){ 
    //     event.preventDefault(); 
    // });
});

function searchStock() {
    var strSearch       = document.forms[0].txtSearch.value;
    var strDepartment   = document.forms[0].cboDepartment.value;
    var strUser         = document.forms[0].cboUser.value;
    var strStatus       = document.forms[0].cboStatus.value;
    var strSort         = document.forms[0].cboSort.value;

    document.location.href = 'loan_renewal.asp?type=search&txtSearch=' + strSearch + '&department=' + strDepartment + '&user=' + strUser + '&status=' + strStatus + '&sort=' + strSort;
}

function resetSearch() {
    document.location.href = 'loan_renewal.asp?type=reset';
}

function submitApproval(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true) {
        theForm.Action.value = 'Approve';
        return true;
    }
}

function submitRejection(theForm) {
    var blnSubmit = true;

    if (blnSubmit == true) {
        theForm.Action.value = 'Reject';
        return true;
    }
}

function validateUpdateComment(theForm) {
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
session.lcid = 2057
sub setSearch
    Select case trim(request("type"))
        case "reset"
            session("loan_renewal_search")          = ""
            session("loan_renewal_department")      = ""
            session("loan_renewal_user")            = ""
            session("loan_renewal_status")          = ""
            session("loan_renewal_sort")            = ""
            session("loan_renewal_initial_page")    = 1
        case "search"
            session("loan_renewal_search")          = trim(request("txtSearch"))
            session("loan_renewal_department")      = trim(request("department"))
            session("loan_renewal_user")            = trim(request("user"))
            session("loan_renewal_status")          = trim(request("status"))
            session("loan_renewal_sort")            = trim(request("sort"))
            session("loan_renewal_initial_page")    = 1
    end Select
end sub

sub displayLoanStock
    dim iRecordCount
    iRecordCount = 0
    dim strSortBy
    dim strSortItem
    dim strDays
    dim strSQL

    dim strPageResultNumber
    dim strRecordPerPage
    dim intRecordCount

    dim strTodayDate
    strTodayDate = FormatDateTime(Date())

    'strSearchTxt = trim(Request("txtSearch"))

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 5000

    if session("loan_renewal_sort") = "" then
        session("loan_renewal_sort") = "latest"
    end if

    if session("loan_renewal_status") = "" then
        session("loan_renewal_status") = "1"
    end if

    strSQL = strSQL & " SELECT * FROM tbl_loan_renewal "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "     (renAccountCode LIKE '%" & UCASE(trim(session("loan_renewal_search"))) & "%' "
    strSQL = strSQL & "         OR renAccountName LIKE '%" & UCASE(trim(session("loan_renewal_search"))) & "%' "
    strSQL = strSQL & "         OR renItemCode LIKE '%" & UCASE(trim(session("loan_renewal_search"))) & "%' "
    strSQL = strSQL & "         OR renSerialNo LIKE '%" & UCASE(trim(session("loan_renewal_search"))) & "%') "
    strSQL = strSQL & "             AND renDepartment LIKE '%" & trim(session("loan_renewal_department")) & "%' "
    strSQL = strSQL & "             AND renAccountCode LIKE '%" & UCASE(session("loan_renewal_user")) & "%' "
    if session("loan_renewal_status") = "1" then
        strSQL = strSQL & "     AND renStatus = 1 "
    else
        strSQL = strSQL & "     AND renStatus <> 1 "
    end if
    strSQL = strSQL & " ORDER BY "

    select case session("loan_renewal_sort")
        case "latest"
            strSQL = strSQL & " renDateCreated DESC"
        case "oldest"
            strSQL = strSQL & " renDateCreated"
        case "product"
            strSQL = strSQL & " renItemCode"
        case "expensive"
            strSQL = strSQL & " renLIC"
        case "cheapest"
            strSQL = strSQL & " renLIC DESC"
    end select

    'response.write strSQL

    rs.Open strSQL, conn

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then
        rs.AbsolutePage = session("loan_renewal_initial_page")
        For intRecord = 1 To rs.PageSize
            strDisplayList = strDisplayList & "<tr>"
            strDisplayList = strDisplayList & "<td>" & rs("renID") & "</td>"
            strDisplayList = strDisplayList & "<td>" & FormatDateTime(rs("renDateCreated"),1) & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("renDepartment") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("renAccountCode") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("renAccountName") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("renItemCode") & "</td>"
            strDisplayList = strDisplayList & "<td>" & rs("renSerialNo") & "</td>"
            strDisplayList = strDisplayList & "<td nowrap>" & rs("renLocation") & "</td>"
            strDisplayList = strDisplayList & "<td align=""right"">" & FormatNumber(rs("renLIC")) & "</td>"
            strDisplayList = strDisplayList & "<td align=""right"" nowrap>" & FormatDateTime(rs("renLoanDate"),1) & "</td>"
            strDisplayList = strDisplayList & "<td align=""right"" nowrap>" & FormatDateTime(rs("renExpiryDate"),1) & "</td>"
			strDisplayList = strDisplayList & "<td>"
			'if IsNull(rs("renComments")) then
			'	strDisplayList = strDisplayList & "	<form method=""post"" name=""form_add_comment"" id=""form_add_comment"" onsubmit=""return validateUpdateComment(this)"">"
			'	strDisplayList = strDisplayList & "	<table>"
			'	strDisplayList = strDisplayList & "		<tr>"
			'	strDisplayList = strDisplayList & "			<td align=""center"">"
			'	strDisplayList = strDisplayList & "				<input type=""hidden"" name=""action"" value=""Add"">"
			'	strDisplayList = strDisplayList & "				<input type=""hidden"" name=""renID"" value=""" & Trim(rs("renID")) & """>"				
			'	strDisplayList = strDisplayList & "				<input type=""text"" id=""txtComment"" name=""txtComment"" class=""form-control"" maxlength=""255"" size=""30"" required>"
			'	strDisplayList = strDisplayList & "			</td>"
			'	strDisplayList = strDisplayList & "			<td class=""save-column""><input type=""submit"" value=""Save"" class=""btn btn-primary"" /></td>"
			'	strDisplayList = strDisplayList & "		</tr>"
			'	strDisplayList = strDisplayList & "	</table>"
			'	strDisplayList = strDisplayList & "	</form>"
			'else
				strDisplayList = strDisplayList & "	<form method=""post"" name=""form_add_comment"" id=""form_add_comment"" onsubmit=""return validateUpdateComment(this)"">"
				strDisplayList = strDisplayList & "	<table>"
				strDisplayList = strDisplayList & "	<tr>"
				strDisplayList = strDisplayList & "		<td align=""center"">"
				strDisplayList = strDisplayList & "			<input type=""hidden"" name=""action"" value=""Update"">"
				strDisplayList = strDisplayList & "			<input type=""hidden"" name=""renID"" value=""" & Trim(rs("renID")) & """>"
				strDisplayList = strDisplayList & "<input type=""text"" id=""txtComments"" name=""txtComments"" class=""form-control"" maxlength=""255"" size=""30"" value=""" & rs("renComments") & """ required>"
				strDisplayList = strDisplayList & "		</td>"
				strDisplayList = strDisplayList & "		<td class=""save-column""><input type=""submit"" value=""Save"" class=""btn btn-primary"" /></td>"
				strDisplayList = strDisplayList & "	</tr>"
				strDisplayList = strDisplayList & "	</table>"
				strDisplayList = strDisplayList & "	</form>"
			'end if
			strDisplayList = strDisplayList & "</td>"	
			
            strDisplayList = strDisplayList & "<td align=""left"" nowrap>"

            Select Case	rs("renStatus")
                case 0
                    strDisplayList = strDisplayList & "<span class=""text-success""><i class=""fa fa-check""></i> Approved</span>"
                case 1
                    strDisplayList = strDisplayList & "<table><tr><td align=""center"">"
                    strDisplayList = strDisplayList & "     <form method=""post"" name=""form_approve"" id=""form_approve"" onsubmit=""return submitApproval(this)"">"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""action"" value=""Approve"">"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renID"" value=""" & rs("renID") & """>"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renOrderNo"" value=""" & rs("renOrderNo") & """>"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renOrderLine"" value=""" & rs("renOrderLine") & """>"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renItemCode"" value=""" & rs("renItemCode") & """>"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renCreatedByEmail"" value=""" & rs("renCreatedByEmail") & """>"
                    strDisplayList = strDisplayList & "         <input type=""submit"" "
					if session("logged_username") <> "gandig" and session("logged_username") <> "tasih" and session("logged_username") <> "simong" and session("logged_username") <> "marka"then
						strDisplayList = strDisplayList & "	disabled"
					end if
					strDisplayList = strDisplayList & " value=""Approve"" class=""btn btn-success btn-sm"" />"
                    strDisplayList = strDisplayList & "     </form>"
                    strDisplayList = strDisplayList & " </td>"
                    strDisplayList = strDisplayList & " <td align=""center"">"
                    strDisplayList = strDisplayList & "     <form method=""post"" name=""form_reject"" id=""form_reject"" onsubmit=""return submitRejection(this)"">"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""action"" value=""Reject"">"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renID"" value=""" & rs("renID") & """>"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renItemCode"" value=""" & rs("renItemCode") & """>"
                    strDisplayList = strDisplayList & "         <input type=""hidden"" name=""renCreatedByEmail"" value=""" & rs("renCreatedByEmail") & """>"
                    strDisplayList = strDisplayList & "         <input type=""submit"" "
					if session("logged_username") <> "gandig" and session("logged_username") <> "tasih" and session("logged_username") <> "simong" and session("logged_username") <> "marka"then
						strDisplayList = strDisplayList & "	disabled"
					end if
					strDisplayList = strDisplayList & " value=""Reject"" class=""btn btn-danger btn-sm"" />"
                    strDisplayList = strDisplayList & "	    </form>"
                    strDisplayList = strDisplayList & " </td></tr></table>"
                case 2
                    strDisplayList = strDisplayList & "<span class=""text-warning""><i class=""fa fa-times""></i> Rejected</span>"
            end select

            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next

    else
        strDisplayList = "<tr><td colspan=""12"">No renewals found.</td></tr>"
    end if
    strDisplayList = strDisplayList & "<tr>"
    strDisplayList = strDisplayList & "<td colspan=""12"">"
    strDisplayList = strDisplayList & "<h3>Total: <u>" & intRecordCount & "</u> renewals</h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
    call getEmployeeDetails(session("logged_username"))
    call setSearch 
    
    if trim(session("loan_renewal_initial_page"))  = "" then
        session("loan_renewal_initial_page") = 1
    end if

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        intRenewalID        = Trim(Request("renID"))
        strOrderNo          = Trim(Request("renOrderNo"))
        strOrderLine        = Trim(Request("renOrderLine"))
        strItemCode         = Trim(Request("renItemCode"))
        strCreatedByEmail   = Trim(Request("renCreatedByEmail"))
		strRenComments		= Replace(Trim(Request.Form("txtComments")),"'","''")
		
        select case Trim(Request("Action"))
            case "Approve"
                call approveLoanRenewal(intRenewalID, strItemCode, strCreatedByEmail, strRenComments, session("logged_username"))
                call incrementRenewalCounter(strOrderNo, strOrderLine)
            case "Reject"
                call rejectLoanRenewal(intRenewalID, strItemCode, strCreatedByEmail, strRenComments, session("logged_username"))
			case "Update"				
                call updateComments(intRenewalID, strItemCode, strRenComments, session("logged_username"))			
        end select
    end if

    call displayLoanStock
end sub

call main

dim strDisplayList

dim intRenewalID
dim strOrderNo
dim strOrderLine
dim strItemCode
dim strCreatedByEmail
%>

    <div class="blog-masthead">
        <div class="container">
            <nav class="blog-nav">
              <a class="blog-nav-item" href="loan_summary.asp"><i class="fa fa-home fa-lg"></i></a>
              <a class="blog-nav-item" href="loan-transfer.asp">Transfer</a>
              <a class="blog-nav-item" href="loan-sale.asp">Sale</a>
              <a class="blog-nav-item active">Renewal
              </a>
            </nav>
        </div>
    </div>

    <div class="container">
        <h1 class="page-header"><i class="fa fa-repeat"></i> Loan Stock Renewal</h1>
        <form name="frmSearch" id="frmSearch" method="post" action="loan_renewal.asp?type=search" onsubmit="searchStock()" class="form-inline">
            <div class="form-group">
                <input type="text" name="txtSearch" maxlength="15" size="35" class="form-control" placeholder="Account / Account Name / Item Code / Serial No" value="<%= request("txtSearch") %>" />
            </div>
            <div class="form-group">
                <select name="cboDepartment" class="form-control" onchange="searchStock()">
                    <option value="">All Dept</option>
                    <option <% if session("loan_renewal_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                    <option <% if session("loan_renewal_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                    <option <% if session("loan_renewal_department") = "F" then Response.Write " selected" end if%> value="F">O&amp;F</option>
                    <option <% if session("loan_renewal_department") = "YMEC" then Response.Write " selected" end if%> value="YMEC">YMEC</option>
                </select>
            </div>
            <div class="form-group">
                <select name="cboUser" class="form-control" onchange="searchStock()">
                    <option <% if session("loan_renewal_user") = "" then Response.Write " selected" end if%> value="">All Users</option>
                    <option <% if session("loan_renewal_user") = "7CAM01" then Response.Write " selected" end if%> value="7CAM01">Cameron Tait</option>
                    <option <% if session("loan_renewal_user") = "7CH001" then Response.Write " selected" end if%> value="7CH001">Chris Herring</option>
                    <option <% if session("loan_renewal_user") = "7DM001" then Response.Write " selected" end if%> value="7DM001">Dale Moore</option>
                    <option <% if session("loan_renewal_user") = "7DMRS0" then Response.Write " selected" end if%> value="7DMRS0">Dale Moore (RS)</option>
                    <option <% if session("loan_renewal_user") = "7DH000" then Response.Write " selected" end if%> value="7DH000">Damien Henderson</option>
                    <option <% if session("loan_renewal_user") = "7DH002" then Response.Write " selected" end if%> value="7DH002">Damien Henderson (RS)</option>
                    <option <% if session("loan_renewal_user") = "7DT001" then Response.Write " selected" end if%> value="7DT001">Dave Thwaites</option>
                    <option <% if session("loan_renewal_user") = "7FED00" then Response.Write " selected" end if%> value="7FED00">Felix Elliot-Dedman</option>
                    <option <% if session("loan_renewal_user") = "7EM001" then Response.Write " selected" end if%> value="7EM001">Euan McInnes</option>
                    <option <% if session("loan_renewal_user") = "7EM002" then Response.Write " selected" end if%> value="7EM002">Euan McInnes (VOX Pedal)</option>
                    <option <% if session("loan_renewal_user") = "7GN001" then Response.Write " selected" end if%> value="7GN001">George Nasr</option>
                    <option <% if session("loan_renewal_user") = "7GL000" then Response.Write " selected" end if%> value="7GL000">Grant Lane</option>
                    <option <% if session("loan_renewal_user") = "7JW001" then Response.Write " selected" end if%> value="7JW001">Jaclyn Williams</option>
                    <option <% if session("loan_renewal_user") = "7JG000" then Response.Write " selected" end if%> value="7JG000">Jamie Goff</option>
                    <option <% if session("loan_renewal_user") = "7AUDW0" then Response.Write " selected" end if%> value="7AUDW0">Jamie Goff (AUDW)</option>
                    <option <% if session("loan_renewal_user") = "7JS001" then Response.Write " selected" end if%> value="7JS001">John Saccaro</option>
                    <option <% if session("loan_renewal_user") = "7JP001" then Response.Write " selected" end if%> value="7JP001">Joseph Pantalleresco</option>
                    <option <% if session("loan_renewal_user") = "7JD001" then Response.Write " selected" end if%> value="7JD001">Justin D'offay</option>
                    <option <% if session("loan_renewal_user") = "7JD002" then Response.Write " selected" end if%> value="7JD002">Justin D'offay (RS)</option>
                    <option <% if session("loan_renewal_user") = "7KJ001" then Response.Write " selected" end if%> value="7KJ001">Kevin Johnson</option>
                    <option <% if session("loan_renewal_user") = "7LB000" then Response.Write " selected" end if%> value="7LB000">Leon Blaher</option>
                    <option <% if session("loan_renewal_user") = "7MC001" then Response.Write " selected" end if%> value="7MC001">Mark Condon</option>
                    <option <% if session("loan_renewal_user") = "7ML001" then Response.Write " selected" end if%> value="7ML001">Mark Loey</option>
                    <option <% if session("loan_renewal_user") = "7ML002" then Response.Write " selected" end if%> value="7ML002">Mark Lapthorne</option>
                    <option <% if session("loan_renewal_user") = "7MT000" then Response.Write " selected" end if%> value="7MT000">Mathew Taylor</option>
                    <option <% if session("loan_renewal_user") = "7MH000" then Response.Write " selected" end if%> value="7MH000">Mick Hughes</option>
                    <option <% if session("loan_renewal_user") = "7NB001" then Response.Write " selected" end if%> value="7NB001">Nathan Biggin</option>
                    <option <% if session("loan_renewal_user") = "7PW000" then Response.Write " selected" end if%> value="7PW000">Paul Wheeler</option>
                    <option <% if session("loan_renewal_user") = "7BP000" then Response.Write " selected" end if%> value="7BP000">Peter Beveridge</option>
                    <option <% if session("loan_renewal_user") = "7RW001" then Response.Write " selected" end if%> value="7RW001">Russell Wykes</option>
                    <option <% if session("loan_renewal_user") = "7SG001" then Response.Write " selected" end if%> value="7SG001">Simon Goldsworthy</option>
                    <option <% if session("loan_renewal_user") = "7SL001" then Response.Write " selected" end if%> value="7SL001">Steve Legg</option>
                    <option <% if session("loan_renewal_user") = "7SVR01" then Response.Write " selected" end if%> value="7SVR01">Steven Vranch</option>
                    <option <% if session("loan_renewal_user") = "7TM000" then Response.Write " selected" end if%> value="7TM000">Terry McMahon</option>
                    <option <% if session("loan_renewal_user") = "7SMC01" then Response.Write " selected" end if%> value="7SMC01">Shaun McMahon</option>
                    <option <% if session("loan_renewal_user") = "7WF001" then Response.Write " selected" end if%> value="7WF001">Wesley Fischer</option>
                    <option <% if session("loan_renewal_user") = "7YME01" then Response.Write " selected" end if%> value="7YME01">YMEC Altona</option>
                    <option <% if session("loan_renewal_user") = "7YME02" then Response.Write " selected" end if%> value="7YME02">YMEC Balwyn</option>
                    <option <% if session("loan_renewal_user") = "7YME09" then Response.Write " selected" end if%> value="7YME09">YMEC Baulkham Hills</option>
                    <option <% if session("loan_renewal_user") = "7YME05" then Response.Write " selected" end if%> value="7YME05">YMEC Carnegie</option>
                    <option <% if session("loan_renewal_user") = "7YME12" then Response.Write " selected" end if%> value="7YME12">YMEC Morley</option>
                    <option <% if session("loan_renewal_user") = "7YME08" then Response.Write " selected" end if%> value="7YME08">YMEC Strathmore</option>
                </select>
            </div>
            <div class="form-group">
                <select name="cboStatus" class="form-control" onchange="searchStock()">
                    <option <% if session("loan_renewal_status") = "1" then Response.Write " selected" end if%> value="1">Pending Approval</option>
                    <option <% if session("loan_renewal_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
            </div>
            <div class="form-group">
                <select name="cboSort" class="form-control" onchange="searchStock()">
                    <option <% if session("loan_renewal_sort") = "latest" then Response.Write " selected" end if%> value="latest">Sort: New - Old</option>
                    <option <% if session("loan_renewal_sort") = "oldest" then Response.Write " selected" end if%> value="oldest">Sort: Old - New</option>
                    <option <% if session("loan_renewal_sort") = "product" then Response.Write " selected" end if%> value="product">Sort: Item code (A-Z)</option>
                    <option <% if session("loan_renewal_sort") = "expensive" then Response.Write " selected" end if%> value="expensive">Sort: Value (High-Low)</option>
                    <option <% if session("loan_renewal_sort") = "cheapest" then Response.Write " selected" end if%> value="cheapest">Sort: Value (Low-High)</option>
                </select>
            </div>
            <input type="button" name="btnSearch" value="Search" class="btn btn-primary" onclick="searchStock()" />
            <input type="button" name="btnReset" value="Reset" class="btn btn-default" onclick="resetSearch()" />
        </form>
        <br>
		</div>	
        <div class="table-responsive">
            <table class="table table-striped loan_table">
                <thead>
                    <tr>
                        <td>ID</td>
                        <td>Request Date</td>
                        <td>Dept</td>
                        <td>Account</td>
                        <td>Account Name</td>
                        <td>Item Code</td>
                        <td>Serial</td>
                        <td>Location</td>
                        <td align="right">LIC $</td>
                        <td align="right">Loan Date</td>
                        <td align="right">Expiry Date</td>
						<td>Comments</td>
                        <td>Status</td>
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