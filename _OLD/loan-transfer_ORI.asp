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
<link rel="stylesheet" href="//maxcdn.bootstrapcdn.com/font-awesome/4.3.0/css/font-awesome.min.css">
<script>
function searchStock(){
	var strSort	  			= document.forms[0].cboSort.value;
    document.location.href 	= 'loan_summary.asp?sort=' + strSort;
}
</script>
</head>
<body>
<%
sub setSearch		
	session("loan_transfer_sort") 		 = trim(request("sort"))
	session("loan_transfer_initial_page") = 1
end sub

sub displayLoanStock
	dim iRecordCount
	iRecordCount = 0
    dim strDays
    dim strSQL
	
	dim intRecordCount
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 5000
	
	if session("loan_transfer_sort") = "" then
		session("loan_transfer_sort") = "name"
	end if
	
	strSQL = strSQL & "SELECT * FROM tbl_loan_transfer"	
	'strSQL = strSQL & " ORDER BY "
	
	'select case session("loan_transfer_sort")
	'	case "account"
	'		strSQL = strSQL & "1"
	'	case "name"
	'		strSQL = strSQL & "2"
	'	case "qty"
	'		strSQL = strSQL & "3 DESC"
	'	case "lic"
	'		strSQL = strSQL & "4 DESC"
	'end select
	
	'response.write strSQL
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("loan_transfer_initial_page")
	
		For intRecord = 1 To rs.PageSize												
			strDisplayList = strDisplayList & "<tr>"
			'strDisplayList = strDisplayList & "<td><a href=""update-loan-transfer.asp?id=" & Trim(rs("traID")) & """><button type=""button"" class=""btn btn-primary""><i class=""fa fa-pencil-square-o""></i> " & rs("traID") & "</button></a></td>"			
			strDisplayList = strDisplayList & "<td>" & rs("traID") & "</td>"
			strDisplayList = strDisplayList & "<td><strong>" & rs("traCreatedBy") & "</strong><br>" & rs("traDateCreated") & "</td>"
			strDisplayList = strDisplayList & "<td nowrap><a href=""loan-user.asp?account=" & Trim(rs("traAccountCode")) & """>" & rs("traAccountCode") & "</a> <i class=""fa fa-arrow-right""></i> <a href=""loan-user.asp?account=" & Trim(rs("traRecipient")) & """>" & rs("traRecipient") & "</a></td>"
			strDisplayList = strDisplayList & "<td><strong>" & rs("traModelNo") & "</strong><br>(" & rs("traQty") & ")</td>"
			strDisplayList = strDisplayList & "<td>" & rs("traSerialNo") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			select case rs("traMarketingApproval") 
				case 0
					strDisplayList = strDisplayList & "<table>"
					strDisplayList = strDisplayList & "		<tr>"
					strDisplayList = strDisplayList & "			<td>"
					strDisplayList = strDisplayList & "				<form method=""post"" name=""form_approve"" id=""form_approve"" action=""loan-transfer.asp"">"
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""action"" value=""approve"">"			
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""traID"" value=""" & rs("traID") & """>"
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""traCreatedBy"" value=""" & rs("traCreatedBy") & """>"
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""traRecipient"" value=""" & rs("traRecipient") & """>"
					strDisplayList = strDisplayList & "					<input type=""submit"" "
					if session("logged_username") <> "harsonos" and session("logged_username") <> "russellw" then
						strDisplayList = strDisplayList & "	disabled"
					end if
					strDisplayList = strDisplayList & "	value=""Approve"" class=""btn btn-success"" />"
					strDisplayList = strDisplayList & "				</form>"
					strDisplayList = strDisplayList & "			</td>"
					strDisplayList = strDisplayList & "			<td class=""save-column"">"
					strDisplayList = strDisplayList & "				<form method=""post"" name=""form_reject"" id=""form_reject"" action=""loan-transfer.asp"">"
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""action"" value=""reject"">"			
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""traID"" value=""" & rs("traID") & """>"
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""traCreatedBy"" value=""" & rs("traCreatedBy") & """>"
					strDisplayList = strDisplayList & "					<input type=""hidden"" name=""traRecipient"" value=""" & rs("traRecipient") & """>"
					strDisplayList = strDisplayList & "					<input type=""submit"" "
					if session("logged_username") <> "harsonos" and session("logged_username") <> "russellw" then
						strDisplayList = strDisplayList & "	disabled"
					end if
					strDisplayList = strDisplayList & "	value=""Reject"" class=""btn btn-danger"" />"
					strDisplayList = strDisplayList & "				</form>"
					strDisplayList = strDisplayList & "			</td>"
					strDisplayList = strDisplayList & "		</tr>"
					strDisplayList = strDisplayList & "</table>"
				case 1
					strDisplayList = strDisplayList & "<font color=""green""><i class=""fa fa-check-square-o""></i> OK</font> <br><strong>" & rs("traMarketingApprovalBy") & "</strong><br>" & rs("traMarketingApprovalDate") & ""
				case 2
					strDisplayList = strDisplayList & "<font color=""red""><i class=""fa fa-ban""></i> Rejected</font> <br><strong>" & rs("traMarketingRejectionBy") & "</strong><br>" & rs("traMarketingRejectionDate") & ""
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			select case rs("traRecipientConfirmation") 
				case 0
					strDisplayList = strDisplayList & "<form method=""post"" name=""form_acknowledge"" id=""form_acknowledge"" action=""loan-transfer.asp"">"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""action"" value=""acknowledge"">"			
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""traID"" value=""" & rs("traID") & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""traCreatedBy"" value=""" & rs("traCreatedBy") & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""traRecipient"" value=""" & rs("traRecipient") & """>"
					strDisplayList = strDisplayList & "		<input type=""submit"" "
					'if rs("traMarketingApproval") = 0 or session("emp_initial") <> rs("traRecipient") or session("logged_username") <> "shaunm" then
					if rs("traMarketingApproval") = 0 or rs("traMarketingApproval") = 2 then
						strDisplayList = strDisplayList & "disabled"
					end if
					strDisplayList = strDisplayList & " value=""Confirm"" class=""btn btn-success"" />"
					
					strDisplayList = strDisplayList & "</form>"
				case 1
					strDisplayList = strDisplayList & "<font color=""green""><i class=""fa fa-check-square-o""></i> OK</font> <br><strong>" & rs("traRecipientConfirmationBy") & "</strong><br>" & rs("traRecipientConfirmationDate") & ""
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			select case rs("traLogisticsConfirmation") 
				case 0
					strDisplayList = strDisplayList & "<form method=""post"" name=""form_confirm"" id=""form_confirm"" action=""loan-transfer.asp"">"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""action"" value=""confirm"">"			
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""traID"" value=""" & rs("traID") & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""traCreatedBy"" value=""" & rs("traCreatedBy") & """>"
					strDisplayList = strDisplayList & "		<input type=""hidden"" name=""traRecipient"" value=""" & rs("traRecipient") & """>"
					strDisplayList = strDisplayList & "		<input type=""submit"" "
					'if rs("traRecipientConfirmation") = 0 and (session("logged_username") <> "craigd" or session("logged_username") <> "kurtt" or session("logged_username") <> "johannas") then
					if rs("traRecipientConfirmation") = 0 then
						strDisplayList = strDisplayList & "disabled"
					end if
					strDisplayList = strDisplayList & "	value=""Confirm"" class=""btn btn-success"" />"
					strDisplayList = strDisplayList & "</form>"
				case 1
					strDisplayList = strDisplayList & "<font color=""green""><i class=""fa fa-check-square-o""></i> OK</font> <br><strong>" & rs("traLogisticsConfirmationBy") & "</strong><br>" & rs("traLogisticsConfirmationDate") & ""	
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_save"" id=""form_save"" action=""loan-transfer.asp"">"
			strDisplayList = strDisplayList & "	<table>"
			strDisplayList = strDisplayList & "		<tr>"
			strDisplayList = strDisplayList & "			<td>"			
			strDisplayList = strDisplayList & "				<input type=""hidden"" name=""action"" value=""save"">"
			strDisplayList = strDisplayList & "				<input type=""hidden"" name=""traID"" value=""" & Trim(rs("traID")) & """>"
			strDisplayList = strDisplayList & "				<input type=""text"" id=""txtConnote"" name=""txtConnote"" class=""form-control"" maxlength=""10"" size=""7"" value=""" & rs("traConnote") & """ required>"
			strDisplayList = strDisplayList & "			</td>"
			strDisplayList = strDisplayList & "			<td class=""save-column"">"
			strDisplayList = strDisplayList & "				<input type=""submit"" value=""Save"" class=""btn btn-primary"" />"			
			strDisplayList = strDisplayList & "			</td>"
			strDisplayList = strDisplayList & "		</tr>"
			strDisplayList = strDisplayList & "	</table>"
			strDisplayList = strDisplayList & "	</form>"
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td nowrap>"
			Select Case	rs("traStatus")
				case 1
					strDisplayList = strDisplayList & "<font color=""blue"">In-progress</font>"
				case 0
					strDisplayList = strDisplayList & "<font color=""green""><i class=""fa fa-check-square-o""></i> Done</font>"
			end select
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td><strong>" & rs("traModifiedBy") & "</strong><br>" & rs("traDateModified") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr><td colspan=""11"">No transfers found.</td></tr>"
	end if
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""11"">"
	strDisplayList = strDisplayList & "<h3>Total transfer: <u>" & intRecordCount & "</u></h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	call getEmployeeDetails(session("logged_username"))
	call setSearch
	
	if trim(session("loan_transfer_initial_page"))  = "" then
    	session("loan_transfer_initial_page") = 1
	end if
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		dim traID, traConnote, traCreatedBy, traRecipient
		traID			= Trim(Request("traID"))
		traConnote  	= Replace(Trim(Request.Form("txtConnote")),"'","''")
		traCreatedBy 	= Trim(Request("traCreatedBy"))
		traRecipient 	= Trim(Request("traRecipient"))
				
		call getRequesterDetails(traCreatedBy)	
		call getRecipientDetails(traRecipient)
		
		Select Case Trim(Request("action"))
			case "approve"
				call approveTransfer(traID,session("logged_username"),session("requester_email"),session("recipient_email"))
			case "reject"
				call rejectTransfer(traID,session("logged_username"),session("requester_email"),session("recipient_email"))
			case "acknowledge"
				call acknowledgeTransfer(traID,session("logged_username"),session("requester_email"))
			case "confirm"
				call confirmTransfer(traID,session("logged_username"),session("requester_email"),session("recipient_email"))	
			case "save"
				call updateTransferConnote(traID,traConnote,session("logged_username"))
		end select
	end if
    
    call displayLoanStock
end sub

call main

dim strDisplayList
%>
<div class="blog-masthead">
  <div class="container">
    <nav class="blog-nav"> <a class="blog-nav-item" href="loan_summary.asp"><i class="fa fa-home fa-lg"></i></a> <a class="blog-nav-item active">Transfer</a> </nav>
  </div>
</div>
<div class="container">
  <h1 class="page-header"><i class="fa fa-share"></i> Loan Stock Transfer (BETA)</h1>
  <div class="table-responsive">
    <table class="table table-striped">
      <thead>
        <tr>
          <td>ID</td>
          <td>Created</td>
          <td>From <i class="fa fa-arrow-right"></i> To</td>
          <td>Model (Qty)</td>
          <td>Serial</td>
          <td>Marketing</td>
          <td>Recipient</td>
          <td>Logistics</td>
          <td>Connote / Sales Order</td>
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