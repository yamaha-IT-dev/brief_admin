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
<!--#include file="../include/connection.asp" -->
<!--#include file="class/clsEmployee.asp" -->
<!--#include file="class/clsLoanSale.asp" -->
<!--#include file="class/clsLoanTransfer.asp" -->
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
    <title>Loan Stock Transfer</title>
    <link rel="stylesheet" href="bootstrap/css/bootstrap.min.css">
    <link rel="stylesheet" href="css/header.css">
    <link rel="stylesheet" href="https://use.fontawesome.com/603ac387f6.css">
    <script>
        function searchTransfer() {
            var strSearch       = document.forms[0].txtSearch.value;
            var strStatus       = document.forms[0].cboStatus.value;
            //var strSort        = document.forms[0].cboSort.value;

            document.location.href = 'loan-sale.asp?type=search&txtSearch=' + strSearch + '&status=' + strStatus;
        }

        function resetSearch() {
            document.location.href = 'loan-sale.asp?type=reset';
        }
    </script>
</head>
<body>
<%
sub setSearch
    select case Trim(Request("type"))
        case "reset"
            session("loan_sale_search")         = ""
            session("loan_sale_status")         = 1
            'session("loan_sale_sort")           = ""
            session("loan_sale_initial_page")   = 1
        case "search"
            session("loan_sale_search")         = Trim(Request("txtSearch"))
            session("loan_sale_status")         = Trim(Request("status"))
            'session("loan_sale_sort")           = Trim(Request("sort"))
            session("loan_sale_initial_page")   = 1
    end select
end sub

sub displayLoanStock
    dim iRecordCount
    iRecordCount = 0
    dim strDays
    dim strSQL

    dim intRecordCount

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 5000

    if session("loan_sale_status") = "" then
        session("loan_sale_status") = 1
    end if

    strSQL = strSQL & "SELECT * FROM tbl_loan_sale"	
    strSQL = strSQL & "	WHERE (saleAccountCode LIKE '%" & session("loan_sale_search") & "%' "
    strSQL = strSQL & "			OR saleModelNo LIKE '%" & session("loan_sale_search") & "%' "
    strSQL = strSQL & "			OR saleOrderNo LIKE '%" & session("loan_sale_search") & "%' "
    strSQL = strSQL & "			OR saleDealerCode LIKE '%" & session("loan_sale_search") & "%') "
    strSQL = strSQL & "			AND saleStatus = '" & session("loan_sale_status") & "' "
    strSQL = strSQL & " ORDER BY saleID DESC"

    'select case session("loan_sale_sort")
    '    case "account"
    '        strSQL = strSQL & "1"
    '    case "name"
    '        strSQL = strSQL & "2"
    '    case "qty"
    '        strSQL = strSQL & "3 DESC"
    '    case "lic"
    '        strSQL = strSQL & "4 DESC"
    'end select

    'response.write strSQL

    rs.Open strSQL, conn

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        rs.AbsolutePage = session("loan_sale_initial_page")

        For intRecord = 1 To rs.PageSize
            strDisplayList = strDisplayList & "<tr>"
            strDisplayList = strDisplayList & "<td>" & rs("saleID") & "</td>"
            strDisplayList = strDisplayList & "<td><strong>" & rs("saleCreatedBy") & "</strong><br>" & rs("saleDateCreated") & "</td>"
            strDisplayList = strDisplayList & "<td nowrap><a href=""loan-user.asp?account=" & Trim(rs("saleAccountCode")) & """>" & rs("saleAccountCode") & "</a></td>"
            strDisplayList = strDisplayList & "<td>" & rs("saleDealerCode") & "</td>"
            strDisplayList = strDisplayList & "<td><strong>" & rs("saleModelNo") & "</strong><br>" & rs("saleSerialNo") & "</td>"
            strDisplayList = strDisplayList & "<td>"
            select case rs("saleLogisticsConfirmation")
                case 0
                    strDisplayList = strDisplayList & "<form method=""post"" name=""form_confirm"" id=""form_confirm"" action=""loan-sale.asp"">"
                    strDisplayList = strDisplayList & "		<input type=""hidden"" name=""action"" value=""confirm"">"
                    strDisplayList = strDisplayList & "		<input type=""hidden"" name=""saleID"" value=""" & rs("saleID") & """>"
                    strDisplayList = strDisplayList & "		<input type=""hidden"" name=""saleCreatedBy"" value=""" & rs("saleCreatedBy") & """>"
                    strDisplayList = strDisplayList & "		<input type=""hidden"" name=""saleCreatedByEmail"" value=""" & rs("saleCreatedByEmail") & """>"
                    strDisplayList = strDisplayList & "		<input type=""hidden"" name=""saleDealerCode"" value=""" & rs("saleDealerCode") & """>"
                    strDisplayList = strDisplayList & "		<input type=""submit"" "
                    if (session("logged_username") <> "craigd" and session("logged_username") <> "kurtt" and session("logged_username") <> "atsukot" and session("logged_username") <> "harsonos") then
                        strDisplayList = strDisplayList & "disabled"
                    end if
                    strDisplayList = strDisplayList & "	value=""Confirm"" class=""btn btn-success"" />"
                    strDisplayList = strDisplayList & "</form>"
                case 1
                    strDisplayList = strDisplayList & "<font color=""green""><i class=""fa fa-check-square-o""></i> Confirmed</font> <br><strong>" & rs("saleLogisticsConfirmationBy") & "</strong><br>" & rs("saleLogisticsConfirmationDate") & ""
            end select
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td>"
            strDisplayList = strDisplayList & "<form method=""post"" name=""form_save"" id=""form_save"" action=""loan-sale.asp"">"
            strDisplayList = strDisplayList & "	<table>"
            strDisplayList = strDisplayList & "		<tr>"
            strDisplayList = strDisplayList & "			<td>"
            strDisplayList = strDisplayList & "				<input type=""hidden"" name=""action"" value=""save"">"
            strDisplayList = strDisplayList & "				<input type=""hidden"" name=""saleID"" value=""" & Trim(rs("saleID")) & """>"
            strDisplayList = strDisplayList & "				<input type=""text"" id=""txtOrderNo"" name=""txtOrderNo"" class=""form-control"" maxlength=""10"" size=""7"" value=""" & rs("salePurchaseOrderNo") & """ required>"
            strDisplayList = strDisplayList & "			</td>"
            strDisplayList = strDisplayList & "			<td class=""save-column"">"
            strDisplayList = strDisplayList & "				<input type=""submit"" value=""Save"" class=""btn btn-primary"" />"
            strDisplayList = strDisplayList & "			</td>"
            strDisplayList = strDisplayList & "		</tr>"
            strDisplayList = strDisplayList & "	</table>"
            strDisplayList = strDisplayList & "	</form>"
            strDisplayList = strDisplayList & "</td>"
            strDisplayList = strDisplayList & "<td nowrap>"
            Select Case	rs("saleStatus")
                case 1
                    strDisplayList = strDisplayList & "<font color=""blue"">In-progress</font>"
                case 2
                    strDisplayList = strDisplayList & "<font color=""red"">Rejected</font>"	
                case 0
                    strDisplayList = strDisplayList & "<font color=""green"">Completed</font>"
            end select
            strDisplayList = strDisplayList & "</td>"

            strDisplayList = strDisplayList & "<td><strong>" & rs("saleModifiedBy") & "</strong><br>" & rs("saleDateModified") & "</td>"
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next
    else
        strDisplayList = "<tr><td colspan=""9"">No records found.</td></tr>"
    end if
    strDisplayList = strDisplayList & "<tr>"
    strDisplayList = strDisplayList & "<td colspan=""9"">"
    strDisplayList = strDisplayList & "<h3>Total: <u>" & intRecordCount & "</u></h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
    call getEmployeeDetails(session("logged_username"))
    call setSearch

    if trim(session("loan_sale_initial_page"))  = "" then
        session("loan_sale_initial_page") = 1
    end if

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        dim saleID, saleNewOrderNo, saleCreatedBy, saleCreatedByEmail, saleDealerCode
        saleID          = Trim(Request("saleID"))
        saleOrderNo     = Replace(Trim(Request.Form("txtOrderNo")),"'","''")
        saleCreatedBy   = Trim(Request("saleCreatedBy"))
        'saleCreatedEmail = Trim(Request("saleCreatedByEmail"))
        saleDealerCode  = Trim(Request("saleDealerCode"))

        call getRequesterDetails(saleCreatedBy)
        'call getRecipientDetails(saleDealerCode)

        Select Case Trim(Request("action"))
            case "confirm"
                call confirmSale(saleID,session("logged_username"),session("requester_email"))
            case "save"
                call updateSaleOrderNo(saleID,saleOrderNo,session("logged_username"))
        end select
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
                <a class="blog-nav-item active">Sale</a>
            </nav>
        </div>
    </div>
    <div class="container-fluid">
        <h1 class="page-header"><i class="fa fa-cart-arrow-down"></i> Loan Stock Sale</h1>
        <form name="frmSearch" id="frmSearch" action="loan-sale.asp?type=search" method="post" onsubmit="searchTransfer()">
            <div class="form-group">
                <input type="text" class="form-control" name="txtSearch" size="70" value="<%= request("txtSearch") %>" maxlength="20" placeholder="Account Code / Dealer / Model / Serial No / Sales Order No" />
            </div>
            <div class="form-group">
                <select name="cboStatus" class="form-control" onchange="searchTransfer()">
                    <option <% if session("loan_sale_status") = "1" then Response.Write " selected" end if%> value="1">In-progress</option>
                    <option <% if session("loan_sale_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
            </div>
            <div class="form-group">
                <input type="button" name="btnSearch" value="Search" onclick="searchTransfer()" class="btn btn-primary" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" class="btn btn-primary" />
            </div>
        </form>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead>
                    <tr>
                        <td>Sale ID</td>
                        <td>Created</td>
                        <td>Loan Account</td>
                        <td>Dealer Code</td>
                        <td>Model (Serial)</td>
                        <td>Logistics</td>
                        <td>Sales Order No</td>
                        <td>Status</td>
                        <td>Modified</td>
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