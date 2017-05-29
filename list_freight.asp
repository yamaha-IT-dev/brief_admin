<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsFreight.asp " -->
<!--#include file="class/clsFreightItem.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>MPD Freight</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchFreight(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_freight.asp?type=search&keyword=' + strSearch + '&status=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_freight.asp?type=reset';
}

function validateUpdateFreightForm(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtConnote,"Connote");	
	reason += validateSpecialCharacters(theForm.txtConnote,"Connote");
	
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
			session("freight_search") 		= ""
			session("freight_status") 		= ""
			session("freight_initial_page") = 1
		case "search"
			session("freight_search") 		= Trim(Request("keyword"))
			session("freight_status") 		= Trim(Request("status"))
			session("freight_initial_page") = 1
	end select
end sub

sub displayFreight
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

	if session("freight_status") = "" then
		session("freight_status") = "2"
	end if
	
	strSQL = "SELECT * FROM mpd_freight WHERE "
	strSQL = strSQL & "		(pickup_name LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR pickup_contact LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR pickup_phone LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR pickup_address LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR pickup_city LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR pickup_state LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR pickup_postcode LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR pickup_comments LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_name LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_contact LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_phone LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_address LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_city LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_state LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_postcode LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR receiver_comments LIKE '%" & session("freight_search") & "%'"
	strSQL = strSQL & "		OR connote LIKE '%" & session("freight_search") & "%')"
	strSQL = strSQL & "	AND status LIKE '%" & session("freight_status") & "%' "
	strSQL = strSQL & "	ORDER BY date_created DESC"
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("freight_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("freight_initial_page") = intpage

			if session("freight_initial_page") < 1 then session("freight_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("freight_initial_page") = intpage

			if session("freight_initial_page") > intPageCount then session("freight_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("freight_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("freight_initial_page")

		For intRecord = 1 To rs.PageSize
			strDays = DateDiff("d",rs("date_created"), strTodayDate)
			
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update_freight"" id=""form_update_freight"" onsubmit=""return validateUpdateFreightForm(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""freight_id"" value=""" & trim(rs("freight_id")) & """>"

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

			strDisplayList = strDisplayList & "<td align=""center"">" & rs("created_by") & ", " & rs("date_created") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pickup_name") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pickup_contact") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pickup_phone") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pickup_address") & ", " & rs("pickup_city") & " " & rs("pickup_state") & " " & rs("pickup_postcode") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("receiver_name") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("receiver_contact") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("receiver_phone") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("receiver_address") & ", " & rs("receiver_city") & " " & rs("receiver_state") & " " & rs("receiver_postcode") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">"
			strDisplayList = strDisplayList & "<select name=""cboStatus"">" 
			select case rs("status")
				case 1
					strDisplayList = strDisplayList & "<option value=""1"" selected>Draft</option>"
					strDisplayList = strDisplayList & "<option value=""2"">New</option>"
					strDisplayList = strDisplayList & "<option value=""0"">Completed</option>"
				case 2
					strDisplayList = strDisplayList & "<option value=""1"">Draft</option>"
					strDisplayList = strDisplayList & "<option value=""2"" selected>New</option>"
					strDisplayList = strDisplayList & "<option value=""0"">Completed</option>"			
				case 0
					strDisplayList = strDisplayList & "<option value=""1"">Draft</option>"
					strDisplayList = strDisplayList & "<option value=""2"">New</option>"
					strDisplayList = strDisplayList & "<option value=""0"" selected>Completed</option>"
			end select
			strDisplayList = strDisplayList & "</select>"
			strDisplayList = strDisplayList & "</td>"						
			
			strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtConnote"" name=""txtConnote"" maxlength=""8"" size=""10"" value=""" & rs("connote") & """ ></td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if rs("status") <> 0 then
				strDisplayList = strDisplayList & "<input type=""submit"" value=""Update"" />"
			end if	
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""12"" align=""center"">No freights found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""12"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_freight.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("freight_initial_page") & ">"

	if session("freight_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("freight_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboStatus"" value=" & strStatus & ">"
    strDisplayList = strDisplayList & "<p>Page: " & session("freight_initial_page") & " to " & intpagecount & "</p>"
	strDisplayList = strDisplayList & "<h3>Total: " & intRecordCount & " freights.</h3>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		intFreightID	= Request("freight_id")
		intStatus 		= Request("cboStatus")
		strConnote		= Trim(Request.Form("txtConnote"))
		
		Select Case Trim(Request("action"))
			case "update"
				call updateFreightConnote(intFreightID, intStatus, strConnote, session("logged_username"))
		end select
	else
		call setSearch
	
		if trim(session("freight_initial_page")) = "" then
			session("freight_initial_page") = 1
		end if
		
		call displayFreight	
	end if
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList

dim intID, intInstructionID, intReasonID, strGRA
%>
<div class="main">
  <!-- #include file="include/header.asp" -->
  <h2>MPD Freight</h2>
  <div class="alert alert-search">
    <form name="frmSearch" id="frmSearch" action="list_freight.asp?type=search" method="post" onsubmit="searchFreight()">
      <h3>Search Parameters:</h3>
      Name / Contact / Phone / Address / Comments / Connote:
       <input type="text" name="txtSearch" size="25" value="<%= request("keyword") %>" maxlength="20" />    
      <select name="cboStatus" onchange="searchFreight()">
      	<option <% if session("freight_status") = "2" then Response.Write " selected" end if%> value="2">New</option>
        <option <% if session("freight_status") = "1" then Response.Write " selected" end if%> value="1">Draft</option>        
        <option <% if session("freight_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
      </select>
      <input type="button" name="btnSearch" value="Search" onclick="searchFreight()" />
      <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
    </form>
  </div>
  <div align="right"><img src="images/legend-blue.gif" border="1" /> = updated today</div>
  <table cellspacing="0" cellpadding="5" class="loan_table" border="0">
    <tr class="loan_header_row">
      <td>&nbsp;</td>	
      <td colspan="4" class="list_subheader">PICKUP</td>
      <td colspan="4" class="list_subheader">RECEIVER</td>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr class="loan_header_row">
      <td>Created</td>
      <td>Name</td>
      <td>Contact</td>
      <td>Phone</td>
      <td>Address</td>         
      <td>Name</td>
      <td>Contact</td>
      <td>Phone</td>
      <td>Address</td>      
      <td>Status</td>
      <td>Connote</td>
      <td>&nbsp;</td>
    </tr>
    <%= strDisplayList %>
  </table>
</div>
</body>
</html>