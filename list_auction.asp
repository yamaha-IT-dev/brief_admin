<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsAuction.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Auction Temp Table</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchAuction(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strSort 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_auction.asp?type=search&keyword=' + strSearch + '&sort=' + strSort;
}

function resetSearch(){
	document.location.href = 'list_auction.asp?type=reset';
}

function validateUpdateAuctionForm(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtTitle,"Title");	
	reason += validateSpecialCharacters(theForm.txtTitle,"Title");
	
	reason += validateEmptyField(theForm.txtDescription,"Description");	
	reason += validateSpecialCharacters(theForm.txtDescription,"Description");
	
	reason += validateNumeric(theForm.txtReserve,"Reserve");
	
	reason += validateEmptyField(theForm.txtLocation,"Location");	
	reason += validateSpecialCharacters(theForm.txtLocation,"Location");
	
	
	
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
		'case "reset"
		'	session("auction_search") 		= ""
		'	session("auction_sort") 		= ""
		case "search"
			session("auction_search") 		= Trim(Request("keyword"))
			session("auction_sort") 		= Trim(Request("sort"))			
	end select
end sub

sub displayAuction
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
	rs.PageSize = 500

	if session("auction_sort") = "" then
		session("auction_sort") = "aucItemCode"
	end if
	
	strSQL = "SELECT * FROM tbl_auction WHERE "
	strSQL = strSQL & "		aucItemCode LIKE '%" & session("auction_search") & "%'"
	strSQL = strSQL & "		OR aucSerialNo LIKE '%" & session("auction_search") & "%'"
	strSQL = strSQL & "		OR aucAccountCode LIKE '%" & session("auction_search") & "%'"
	strSQL = strSQL & "		OR aucAccountName LIKE '%" & session("auction_search") & "%'"	
	strSQL = strSQL & "	ORDER BY " & session("auction_sort")
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount	

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

		For intRecord = 1 To rs.PageSize
			strDays = DateDiff("d",rs("aucDateCreated"), strTodayDate)
			
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update"" id=""form_update"" onsubmit=""return validateUpdateAuctionForm(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""aucID"" value=""" & trim(rs("aucID")) & """>"

			if (DateDiff("d",rs("aucDateModified"), strTodayDate) = 0) then
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""updated_today"">"
				else
					strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
				end if
			else
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				else
					strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
				end if
			end if
			strDisplayList = strDisplayList & "<td align=""center""><a onclick=""return confirm('Are you sure you want to delete " & rs("aucItemCode") & " ?');"" href='delete_auction.asp?id=" & rs("aucID") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>" & rs("aucCreatedBy") & ", " & FormatDateTime(rs("aucDateCreated"),1) & ""
			if DateDiff("d",rs("aucDateCreated"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucItemCode") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucSerialNo") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucLIC") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucAccountCode") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("aucAccountName") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtTitle"" name=""txtTitle"" maxlength=""30"" size=""35"" value=""" & rs("aucItemTitle") & """ ></td>"
			strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtDescription"" name=""txtDescription"" maxlength=""40"" size=""45"" value=""" & rs("aucDescription") & """ ></td>"
			strDisplayList = strDisplayList & "<td align=""center"">$ <input type=""text"" id=""txtReserve"" name=""txtReserve"" maxlength=""3"" size=""3"" value=""" & rs("aucReservePrice") & """ ></td>"
			strDisplayList = strDisplayList & "<td align=""center""><input type=""text"" id=""txtLocation"" name=""txtLocation"" maxlength=""30"" size=""30"" value=""" & rs("aucLocation") & """ ></td>"
			strDisplayList = strDisplayList & "<td align=""center""><input type=""submit"" value=""Update"" /></td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""12"" align=""center"">No auction items found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""12"" class=""recordspaging"">"	
	strDisplayList = strDisplayList & "<h3>Total: " & intRecordCount & " items.</h3>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		aucID			= Request("aucID")		
		aucItemTitle	= Replace(Request.Form("txtTitle"),"'","''")
		aucDescription	= Replace(Request.Form("txtDescription"),"'","''")
		aucReservePrice = Trim(Request.Form("txtReserve"))
		aucLocation 	= Replace(Request.Form("txtLocation"),"'","''")		
		
		Select Case Trim(Request("action"))
			case "update"
				call updateAuction(aucID, aucItemTitle, aucDescription, aucReservePrice, aucLocation, session("logged_username"))
		end select		
	end if
	
	call setSearch
	call displayAuction	
end sub

call main

dim rs, intPageCount, intpage, intRecord, strDisplayList
dim aucID, aucItemTitle, aucDescription, aucReservePrice, aucLocation
%>
<div class="main">
  <!-- #include file="include/header_auction.asp" -->
  <h2>Auction Temporary Table
    </h2><div class="alert alert-search">
      <form name="frmSearch" id="frmSearch" action="list_auction.asp?type=search" method="post" onsubmit="searchAuction()">
      <strong>Search: </strong>
      Item Code / Serial No / Account Code / Account Name:
       <input type="text" name="txtSearch" size="25" value="<%= request("keyword") %>" maxlength="20" />    
      <select name="cboStatus" onchange="searchAuction()">
      	<option <% if session("auction_sort") = "aucItemCode" then Response.Write " selected" end if%> value="aucItemCode">Sort by: Item Code</option>
        <option <% if session("auction_sort") = "aucAccountName" then Response.Write " selected" end if%> value="aucAccountName">Sort by: Account Name</option>
        <option <% if session("auction_sort") = "aucDateCreated" then Response.Write " selected" end if%> value="aucDateCreated">Sort by: Date Created</option>
      </select>
      <input type="button" name="btnSearch" value="Search" onclick="searchAuction()" />
      <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
    </form>
  </div>
  
  <p align="right"><img src="images/icon_excel.jpg" /> <a href="export_auction.asp?search=<%= session("auction_search") %>&sort=<%= session("auction_sort") %>">Export</a></p>
  <table cellspacing="0" cellpadding="5" class="loan_table" border="0">
    <tr class="loan_header_row">
      <td>&nbsp;</td>
      <td>Created</td>
      <td>Item Code</td>
      <td>Serial #</td>
      <td>LIC</td>
      <td>Acc. Code</td>
      <td>Name</td>
      <td>Title</td>         
      <td>Description</td>
      <td>Reserve</td>
      <td>Location</td>
      <td>&nbsp;</td>
    </tr>
    <%= strDisplayList %>
  </table>
  <p align="right"><img src="images/legend-blue.gif" border="1" /> = updated today</p>
</div>
</body>
</html>