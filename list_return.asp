<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsEmployee.asp " -->
<!--#include file="class/clsWarehouseReturn.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Warehouse Return</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strType  		= document.forms[0].cboType.value;
	var strDepartment  	= document.forms[0].cboDepartment.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_return.asp?type=search&txtSearch=' + strSearch + '&cboType=' + strType + '&cboDepartment=' + strDepartment + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_return.asp?type=reset';
}

function validateUpdateReturnForm(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtGRA,"GRA");	
	reason += validateSpecialCharacters(theForm.txtGRA,"GRA");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'update';
  		theForm.submit();

		return true;
    }
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("return_search") 		= ""
			session("return_type") 			= ""
			session("return_department") 	= ""
			session("return_status") 		= ""
			session("return_initial_page") 	= 1
		case "search"
			session("return_search") 		= trim(Request("txtSearch"))
			session("return_type") 			= request("cboType")
			session("return_department") 	= request("cboDepartment")
			session("return_status") 		= Trim(Request("cboStatus"))
			session("return_initial_page") 	= 1
	end select
end sub

sub displayReturn
	dim iRecordCount
	iRecordCount = 0
    dim strSQL
	dim intRecordCount
	dim strTodayDate
	dim strDays

	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	if session("return_status") = "" then
		session("return_status") = "1"
	end if
	
	strSQL = "SELECT * FROM yma_quarantines "
	strSQL = strSQL & "	WHERE department LIKE '%" & session("return_department") & "%' "
	if session("return_type") <> "" then
		strSQL = strSQL & "		AND return_type = '" & session("return_type") & "' "
	end if
	strSQL = strSQL & "	AND (item_code LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "		OR shipment_no LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "		OR description LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "		OR return_carrier LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "		OR return_connote LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "		OR serial_no LIKE '%" & session("return_search") & "%' ) "
	strSQL = strSQL & "	AND status LIKE '%" & session("return_status") & "%' "
	strSQL = strSQL & "	ORDER BY date_created DESC"
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("return_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("return_initial_page") = intpage

			if session("return_initial_page") < 1 then session("return_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("return_initial_page") = intpage

			if session("return_initial_page") > intPageCount then session("return_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("return_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("return_initial_page")

		For intRecord = 1 To rs.PageSize
			strDays = DateDiff("d",rs("date_created"), strTodayDate)
			
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update_return"" id=""form_update_return"" onsubmit=""return validateUpdateReturnForm(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""quarantine_id"" value=""" & trim(rs("quarantine_id")) & """>"

			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) then
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""updated_today"">"
				else
					strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
				end if
			else
				'strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				else
					strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
				end if
			end if

			'strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""update_return.asp?id=" & rs("quarantine_id") & """>Edit</a></td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("return_type")
				case 1
					strDisplayList = strDisplayList & "Managed"
				case 2
					strDisplayList = strDisplayList & "Un-addressed"
				case 0
					strDisplayList = strDisplayList & "Un-managed"				
				case else
			 		strDisplayList = strDisplayList & "-"
			end select
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("department") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("item_code") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			Select Case rs("stock_type")
				case "1"
					strDisplayList = strDisplayList & "<font color=red>Damaged</font>"
				case "2"
					strDisplayList = strDisplayList & "<font color=red>Partial</font>"
				case else
			 		strDisplayList = strDisplayList & "-"
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("return_connote") & "</td>"	
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("dealer") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("shipment_no") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("qty") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if rs("photos") = 1 then				
				strDisplayList = strDisplayList & "<a href=""file:\\YAMMAS22\quarantine\" & rs("quarantine_id") & """ target=""_blank"" class=""screenshot"" rel=""file:\\YAMMAS22\quarantine\" & rs("quarantine_id") & "\1.jpg""><img src=""images/camera_icon.gif"" border=""0""></a>"
			else
				strDisplayList = strDisplayList & "-"
			end if
			strDisplayList = strDisplayList & "</td>"

			strDisplayList = strDisplayList & "<td align=""center"">" & rs("return_carrier") & "</td>"		
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("original_connote") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("serial_no") & "</td>"
			
			'strDisplayList = strDisplayList & "<td align=""center"">"
			'Select Case rs("instruction")
			'	case "1"
			'		strDisplayList = strDisplayList & "Return to good stock 3T"
			'	case "2"
			'		strDisplayList = strDisplayList & "Transfer to Excel 3XL"
			'	case "3"
			'		strDisplayList = strDisplayList & "Resend to customer"
			'	case "4"
			'		strDisplayList = strDisplayList & "Damaged item to Excel - good stock to 3T"		
			'	case else
			' 		strDisplayList = strDisplayList & "-"
			'end select
			'strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center"">"
			strDisplayList = strDisplayList & "	<select name=""cboInstruction"">" 
			select case rs("instruction")
				case 1
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"" selected>Return to good stock 3T</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Transfer to Excel 3XL</option>"
					strDisplayList = strDisplayList & "<option value=""3"">Resend to customer</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Damaged item to Excel - good stock to 3T</option>"
				case 2
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Return to good stock 3T</option>"
					strDisplayList = strDisplayList & "<option value=""2"" selected>Transfer to Excel 3XL</option>"
					strDisplayList = strDisplayList & "<option value=""3"">Resend to customer</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Damaged item to Excel - good stock to 3T</option>"
				case 3
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Return to good stock 3T</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Transfer to Excel 3XL</option>"
					strDisplayList = strDisplayList & "<option value=""3"" selected>Resend to customer</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Damaged item to Excel - good stock to 3T</option>"
				case 4
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Return to good stock 3T</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Transfer to Excel 3XL</option>"
					strDisplayList = strDisplayList & "<option value=""3"">Resend to customer</option>"
					strDisplayList = strDisplayList & "<option value=""4"" selected>Damaged item to Excel - good stock to 3T</option>"			
				case else
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Return to good stock 3T</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Transfer to Excel 3XL</option>"
					strDisplayList = strDisplayList & "<option value=""3"">Resend to customer</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Damaged item to Excel - good stock to 3T</option>"
			end select
			strDisplayList = strDisplayList & "	</select>"
			strDisplayList = strDisplayList & "</td>"
			
			if rs("status") = 1 then
				strDisplayList = strDisplayList & "<td align=""center"">Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
									
			strDisplayList = strDisplayList & "<td align=""center"">"
			strDisplayList = strDisplayList & "	<select name=""cboReason"">" 
			select case rs("reason_code")
				case 1
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"" selected>Damaged in Transit</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Order Cancelled</option>"
					strDisplayList = strDisplayList & "<option value=""3"">No longer required</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Order not in system</option>"
					strDisplayList = strDisplayList & "<option value=""5"">Other</option>"
				case 2
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Damaged in Transit</option>"
					strDisplayList = strDisplayList & "<option value=""2"" selected>Order Cancelled</option>"
					strDisplayList = strDisplayList & "<option value=""3"">No longer required</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Order not in system</option>"
					strDisplayList = strDisplayList & "<option value=""5"">Other</option>"
				case 3
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Damaged in Transit</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Order Cancelled</option>"
					strDisplayList = strDisplayList & "<option value=""3"" selected>No longer required</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Order not in system</option>"
					strDisplayList = strDisplayList & "<option value=""5"">Other</option>"
				case 4
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Damaged in Transit</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Order Cancelled</option>"
					strDisplayList = strDisplayList & "<option value=""3"">No longer required</option>"
					strDisplayList = strDisplayList & "<option value=""4"" selected>Order not in system</option>"
					strDisplayList = strDisplayList & "<option value=""5"">Other</option>"
				case 5
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Damaged in Transit</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Order Cancelled</option>"
					strDisplayList = strDisplayList & "<option value=""3"">No longer required</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Order not in system</option>"
					strDisplayList = strDisplayList & "<option value=""5"" selected>Other</option>"
				case else
					strDisplayList = strDisplayList & "<option value="""">...</option>"	
					strDisplayList = strDisplayList & "<option value=""1"">Damaged in Transit</option>"
					strDisplayList = strDisplayList & "<option value=""2"">Order Cancelled</option>"
					strDisplayList = strDisplayList & "<option value=""3"">No longer required</option>"
					strDisplayList = strDisplayList & "<option value=""4"">Order not in system</option>"
					strDisplayList = strDisplayList & "<option value=""5"">Other</option>"	
			end select
			strDisplayList = strDisplayList & "	</select>"
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtGRA"" name=""txtGRA"" maxlength=""8"" size=""10"" value=""" & rs("gra") & """ ></td>"
			
			if rs("status") = 0 then
				strDisplayList = strDisplayList & "<td align=""center"">-</td>"		
			else
				strDisplayList = strDisplayList & "<td align=""center""><input type=""submit"" value=""Update"" /></td>"	
			end if	
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""17"" align=""center"">No records found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""17"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_return.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("return_initial_page") & ">"

	if session("return_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("return_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strItemDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboStatus"" value=" & strStatus & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("return_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		intID		= Request("quarantine_id")
		intInstructionID = Request("cboInstruction")
		intReasonID	= Request.Form("cboReason")
		strGRA 		= Trim(Request.Form("txtGRA"))
		
		Select Case Trim(Request("action"))
			case "update"
				call updateReturn(intID, intInstructionID, intReasonID, strGRA, session("logged_username"))
				'call displayReturn
		end select
	else
		session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
	
		call getEmployeeDetails(session("logged_username"))
		call setSearch
	
		if trim(session("return_initial_page")) = "" then
			session("return_initial_page") = 1
		end if
		
		call displayReturn	
	end if
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList

dim intID, intInstructionID, intReasonID, strGRA
%>
<div style="padding:20px 20px 20px 20px;">
  <div align="right"><a href="http://intranet/"><img src="images/yamaha_logo.jpg" border="0" /></a></div>
  <h2>Warehouse Return</h2>
  <div class="alert alert-search">
    <form name="frmSearch" id="frmSearch" action="list_return.asp?type=search" method="post" onsubmit="searchItem()">
      <h3>Search Parameters:</h3>
      Item code / Shipment no / Description / Return connote / Serial no :
      <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
      <select name="cboType" onchange="searchItem()">
        <option value="">All Types</option>
        <option <% if session("return_type") = "1" then Response.Write " selected" end if%> value="1">Managed</option>
        <option <% if session("return_type") = "0" then Response.Write " selected" end if%> value="0">Un-managed</option>
        <option <% if session("return_type") = "2" then Response.Write " selected" end if%> value="2">Un-addressed</option>
      </select>
      <select name="cboDepartment" onchange="searchItem()">
        <option value="">All Depts</option>
        <option <% if session("return_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
        <option <% if session("return_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
      </select>
      <select name="cboStatus" onchange="searchItem()">
        <option <% if session("return_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
        <option <% if session("return_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
      </select>
      <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
      <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
    </form>
  </div>
  <div align="right"><img src="images/legend-blue.gif" border="1" /> = updated today</div>
  <table cellspacing="0" cellpadding="5" class="loan_table" border="0">
    <tr class="loan_header_row">
      <td>Type</td>
      <td>Dept</td>
      <td>Item</td>
      <td>Stock</td>
      <td>Return connote</td>
      <td>Dealer</td>
      <td>Shipment</td>
      <td>Qty</td>
      <td>Photo</td>
      <td>Carrier</td>
      <td>Original connote</td>
      <td>Serial no</td>
      <td>Instruction</td>
      <td>Status</td>
      <td>Reason</td>
      <td>GRA</td>
      <td>&nbsp;</td>
    </tr>
    <%= strDisplayList %>
  </table>
</div>
</body>
</html>