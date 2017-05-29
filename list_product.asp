<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>MPD Products</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="JavaScript" type="text/javascript">
function searchAuction(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strSort 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_product.asp?type=search&keyword=' + strSearch + '&sort=' + strSort;
}

function resetSearch(){
	document.location.href = 'list_product.asp?type=reset';
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
			session("base_product_search") 		= ""
			session("base_product_sort") 		= ""
			session("base_product_initial_page")= 1
		case "search"
			session("base_product_search") 		= Trim(Request("keyword"))
			session("base_product_sort") 		= Trim(Request("sort"))	
			session("base_product_initial_page")= 1
	end select
end sub

sub displayProduct
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
	rs.PageSize = 20

	if session("base_product_status") = "" then
		session("base_product_status") = "1"
	end if		
	
	strSQL = "SELECT DISTINCT "
	strSQL = strSQL & "	LTRIM(RTRIM(Y3SOSC)) AS prod_code,"
	strSQL = strSQL & "	LTRIM(RTRIM(Y3GREG)) AS prod_group,"
	strSQL = strSQL & "	LTRIM(RTRIM(YDSGMB)) AS prod_category,"
	strSQL = strSQL & "	LTRIM(RTRIM(Y3SYMB)) AS prod_description," 
	strSQL = strSQL & "	YINWPR AS prod_rrp,"
	strSQL = strSQL & "	Y3JRGR AS prod_weight,"
	strSQL = strSQL & "	Y3SIZW AS prod_width,"
	strSQL = strSQL & "	Y3SIZH AS prod_height,"
	strSQL = strSQL & "	Y3SiZD AS prod_depth,"	
	strSQL = strSQL & "	Y3YOSV AS prod_volume,"
	strSQL = strSQL & "	LTRIM(RTRIM(Y3EANC)) AS ean_code"
	strSQL = strSQL & " 	FROM openquery(as400, 'SELECT * FROM YF3MP"
	strSQL = strSQL & "			INNER JOIN YFDMP ON YDSGCD = Y3GREG"
	strSQL = strSQL & "			INNER JOIN YFIMP ON Y3SOSC = YISOSC" 
	strSQL = strSQL & "		INNER JOIN YF6MP ON Y3SOSC = Y6SOSC"
	strSQL = strSQL & "	WHERE Y3SKKI <> ''D''"
	strSQL = strSQL & "		AND (Y3SOSC LIKE ''%" & session("base_product_search") & "%'')"
	strSQL = strSQL & "		AND YDSKKI <> ''D''"
	strSQL = strSQL & "		AND YISKKI <> ''D''"
	strSQL = strSQL & "		AND YIUSPT = ''S50''"
	strSQL = strSQL & "		AND YDSGID = ''1''"
	strSQL = strSQL & "		AND LEFT(Y3GREG,1) IN (''3'',''4'')"
	strSQL = strSQL & "		AND LEFT(Y3SOSC,1) NOT IN (''#'',''*'')"
	strSQL = strSQL & "		AND Y3LFCY <> ''D''')"
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("base_product_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("base_product_initial_page") = intpage

			if session("base_product_initial_page") < 1 then session("base_product_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("base_product_initial_page") = intpage

			if session("base_product_initial_page") > intPageCount then session("base_product_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("base_product_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("base_product_initial_page")

		For intRecord = 1 To rs.PageSize
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prod_group") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prod_code") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prod_description") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prod_weight") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prod_width") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prod_height") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prod_depth") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""8"" align=""center"">No products found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""8"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_product.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("base_product_initial_page") & ">"

	if session("base_product_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("base_product_initial_page") = intpagecount or intRecordCount = 0 then
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
    strDisplayList = strDisplayList & "Page: " & session("base_product_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<h3>Search results: " & intRecordCount & " products.</h3>"
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
		end select
	end if
	'call getEmployeeDetails(session("logged_username"))
	call setSearch
	
	if trim(session("base_product_initial_page")) = "" then
		session("base_product_initial_page") = 1
	end if
	
	call displayProduct
end sub

call main

dim rs
dim intPageCount, intpage, intRecord, strDisplayList
dim intID, intInstructionID, intReasonID, strGRA
%>
<div style="padding:20px 20px 20px 20px;">
<!-- #include file="include/header_auction.asp" -->
  <h2>MPD Products (BASE)</h2>
  <div class="alert alert-search">
    <form name="frmSearch" id="frmSearch" action="list_product.asp?type=search&keyword=<%= Request.Form("txtSearch") %>" method="post" onsubmit="searchAuction()">
      <strong>Search: </strong> Product Name / Description:
      <input type="text" name="txtSearch" size="25" value="<%= session("base_product_search") %>" maxlength="20" />
      <select name="cboStatus" onchange="searchAuction()">
        <option <% if session("base_product_sort") = "aucItemCode" then Response.Write " selected" end if%> value="aucItemCode">Sort by: Product Name</option>
        <option <% if session("base_product_sort") = "aucAccountName" then Response.Write " selected" end if%> value="aucAccountName">Sort by: Account Name</option>
        <option <% if session("base_product_sort") = "aucDateCreated" then Response.Write " selected" end if%> value="aucDateCreated">Sort by: Date Created</option>
      </select>
      <input type="button" name="btnSearch" value="Search" onclick="searchAuction()" />
      <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
    </form>
  </div>  
  <div align="right"><img src="images/legend-blue.gif" border="1" /> = updated today</div>
  <table cellspacing="0" cellpadding="5" class="loan_table" border="0">
    <tr class="loan_header_row">
      <td class="form_header">Category</td>
      <td class="form_header">Name</td>
      <td class="form_header">Description</td>
      <td class="form_header">Weight (kg)</td>
      <td class="form_header">Width (cm)</td>
      <td class="form_header">Height (cm)</td>
      <td class="form_header">Depth (cm)</td>
      <!--<td class="form_header">Qty</td>
      <td class="form_header">Pallet</td>-->
      <td class="form_header"></td>
    </tr>
    <%= strDisplayList %>
  </table>
</div>
</body>
</html>