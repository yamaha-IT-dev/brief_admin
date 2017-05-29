<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsFreight.asp" -->
<!--#include file="class/clsFreightItem.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Confirm MPD Freight Request</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script language="javascript" type="text/javascript">
function searchProduct(){    
    var strSearch 	= document.forms[0].txtSearch.value;	
	document.location.href = 'confirm_freight.asp?type=search&keyword=' + strSearch;
}
    
function resetSearch(){
	document.location.href = 'confirm_freight.asp?type=reset';    
}  

function validateAddFreightItem(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateNumeric(theForm.txtQty,"Qty");	

  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.Action.value = 'Add';
		return true;
    }
}
</script>
<%
sub setSearch
	select case trim(request("type"))
		case "reset"
			session("mpd_product_search") 		= ""
			session("mpd_product_initial_page") = 1
		case "search"
			session("mpd_product_search") 		= trim(Request("keyword"))
			session("mpd_product_initial_page") = 1
	end select
end sub

sub displayProduct
	dim iRecordCount
	iRecordCount = 0
    dim strSQL
	dim intRecordCount
	dim strTodayDate
	
	strTodayDate = FormatDateTime(Date())
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 20
	
	strSQL = "SELECT prod_id, prod_category, prod_code, prod_description, prod_depth, prod_width, prod_height, prod_weight "
	strSQL = strSQL & "	FROM mpd_product "	
	strSQL = strSQL & "	WHERE "
	strSQL = strSQL & "		(prod_code LIKE '%" & session("mpd_product_search") & "%' "
	strSQL = strSQL & "			OR prod_description LIKE '%" & session("mpd_product_search") & "%' "
	strSQL = strSQL & "			OR prod_category LIKE '%" & session("mpd_product_search") & "%')"	
	strSQL = strSQL & "	ORDER BY prod_code "
	
	'response.Write strSQL
	
	rs.Open strSQL, conn
	
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("mpd_product_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("mpd_product_initial_page") = intpage
			
			if session("mpd_product_initial_page") < 1 then session("mpd_product_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("mpd_product_initial_page") = intpage
			
			if session("mpd_product_initial_page") > intPageCount then session("mpd_product_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("mpd_product_initial_page") = intpage
    end select

    strProductList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("mpd_product_initial_page")
	
		For intRecord = 1 To rs.PageSize
						
			if iRecordCount Mod 2 = 0 then
				strProductList = strProductList & "<tr class=""innerdoc"">"
			else
				strProductList = strProductList & "<tr class=""innerdoc_2"">"
			end if
			
			strProductList = strProductList & "<form method=""post"" name=""form_add_freight_item"" id=""form_add_freight_item"" onsubmit=""return validateAddFreightItem(this)"">"
			strProductList = strProductList & "<input type=""hidden"" name=""action"" value=""add"">"
			strProductList = strProductList & "<input type=""hidden"" name=""prod_id"" value=""" & trim(rs("prod_id")) & """>"
			strProductList = strProductList & "<td align=""center"">" & rs("prod_category") & "</td>"
			strProductList = strProductList & "<td align=""center"">" & rs("prod_code") & "</td>"
			strProductList = strProductList & "<td align=""center"">" & rs("prod_description") & "</td>"
			strProductList = strProductList & "<td align=""center"">" & rs("prod_weight") & "</td>"
			strProductList = strProductList & "<td align=""center"">" & rs("prod_width") & "</td>"
			strProductList = strProductList & "<td align=""center"">" & rs("prod_height") & "</td>"
			strProductList = strProductList & "<td align=""center"">" & rs("prod_depth") & "</td>"
			strProductList = strProductList & "<td align=""center""><input type=""text"" id=""txtQty"" name=""txtQty"" maxlength=""3"" size=""3"" value=""1""></td>"
			strProductList = strProductList & "<td align=""center"">"
			strProductList = strProductList & "<select name=""cboPallet"">"
			strProductList = strProductList & "<option value=""Carton"">Carton</option>"
			strProductList = strProductList & "<option value=""Roadcase"">Roadcase</option>"
			strProductList = strProductList & "<option value=""Other"">Other</option>"		
			strProductList = strProductList & "</select>"
			strProductList = strProductList & "</td>"
			strProductList = strProductList & "<td align=""center""><input type=""submit"" value=""Add"" /></td>"
			strProductList = strProductList & "</tr>"
			strProductList = strProductList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strProductList = "<tr class=""innerdoc""><td colspan=""10"" align=""center"">No products found.</td></tr>"
	end if
	
	strProductList = strProductList & "<tr>"
	strProductList = strProductList & "<td colspan=""10"" align=""center"">"
	strProductList = strProductList & "<form name=""MovePage"" action=""confirm_freight.asp"" method=""post"">"
    strProductList = strProductList & "<input type=""hidden"" name=""intpage"" value=" & session("mpd_product_initial_page") & ">"
	
	if session("mpd_product_initial_page") = 1 then
   		strProductList = strProductList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strProductList = strProductList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strProductList = strProductList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strProductList = strProductList & "<input type=""submit"" name=""action"" value=""<"">"
	end if	
	if session("mpd_product_initial_page") = intpagecount or intRecordCount = 0 then
    	strProductList = strProductList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strProductList = strProductList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strProductList = strProductList & "<input type=""submit"" name=""action"" value="">"">"
    	strProductList = strProductList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strProductList = strProductList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strProductList = strProductList & "<input type=""hidden"" name=""cboStatus"" value=" & strStatus & ">"
    strProductList = strProductList & "<br />"
    strProductList = strProductList & "<small>Page: " & session("mpd_product_initial_page") & " to " & intpagecount
	strProductList = strProductList & "</small><br />"
	strProductList = strProductList & "<strong>Search results: " & intRecordCount & " products.</strong>"
    strProductList = strProductList & "</form>"
    strProductList = strProductList & "</td>"
    strProductList = strProductList & "</tr>"
	
    call CloseDataBase()
end sub

sub displayFreightItems
	dim iRecordCount
	iRecordCount = 0
   
    dim strSQL
	dim intRecordCount

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	strSQL = "SELECT * FROM mpd_freight_item "
	strSQL = strSQL & "	WHERE freight_id = '" & session("new_freight_id") & "' "
	strSQL = strSQL & "	ORDER BY name"

	'Response.Write strSQL

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
		
    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then
		For intRecord = 1 To rs.PageSize
			strDisplayList = strDisplayList & "<tr>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("name") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("details") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("qty") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pallet") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("length") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("width") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("height") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("weight") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next	
	end if

	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""8"" align=""center"">"	
	strDisplayList = strDisplayList & "Total: " & intRecordCount & " items.</h2>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	if Session("new_freight_id") <> "" then
		if Request.ServerVariables("REQUEST_METHOD") = "POST" then
			select case Trim(Request.Form("Action"))
				case "Add"
					strName		= Replace(Trim(Request.Form("txtName")),"'","''")
					strDetails	= Replace(Trim(Request.Form("txtDetails")),"'","''")
					intQty		= Trim(Request.Form("txtQty"))
					strPallet	= Trim(Request.Form("cboPallet"))
					intLength	= Trim(Request.Form("txtLength"))
					intWidth	= Trim(Request.Form("txtWidth"))
					intHeight	= Trim(Request.Form("txtHeight"))
					intWeight	= Trim(Request.Form("txtWeight"))
					
					call addFreightItem(session("new_freight_id"), strName, strDetails, intQty, strPallet, intLength, intWidth, intHeight, intWeight, session("logged_username"))
			end select
		else
		 	call setSearch
			
			if trim(session("mpd_product_initial_page")) = "" then
				session("mpd_product_initial_page") = 1
			end if
		
			
			call displayProduct
			call displayFreightItems
		end if
	else
		call clearFreightSessionVariables
		Response.Redirect("add_freight.asp")
	end if
end sub

call main

dim strMessageText, strProductList, strDisplayList
dim strName, strDetails, intQty, strPallet, intLength, intWidth, intHeight, intWeight
%>
</head>
<body>
<div class="main"> 
  <!-- #include file="include/header.asp" --> 
  <img src="images/backward_arrow.gif" border="0" /> <a href="list_freight.asp">Back to List</a>
  <h2>Submit MPD Freight Request</h2>
  <p><font color="red"><%= strMessageText %></font></p>
  ID: <%= Session("new_freight_id") %>
  <table>
    <tr>
      <td><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="form_header">Pickup Details</td>
          </tr>
          <tr>
            <td width="30%"><strong>Name:</strong></td>
            <td width="70%"><%= Session("strPickupName") %></td>
          </tr>
          <tr>
            <td><strong>Contact Person:</strong></td>
            <td><%= Session("strPickupContact")	%></td>
          </tr>
          <tr>
            <td><strong>Phone:</strong></td>
            <td><%= Session("strPickupPhone") %></td>
          </tr>
          <tr>
            <td><strong>Address:</strong></td>
            <td><%= Session("strPickupAddress") %></td>
          </tr>
          <tr>
            <td><strong>City:</strong></td>
            <td><%= Session("strPickupCity") %></td>
          </tr>
          <tr>
            <td><strong>State:</strong></td>
            <td><%= Session("strPickupState") %></td>
          </tr>
          <tr>
            <td><strong>Postcode:</strong></td>
            <td><%= Session("intPickupPostcode") %></td>
          </tr>
          <tr>
            <td valign="top"><strong>Comments:</strong></td>
            <td><%= Session("strPickupComments") %></td>
          </tr>
        </table></td>
      <td valign="top"><table cellpadding="5" cellspacing="0" class="item_maintenance_box">
          <tr>
            <td colspan="2" class="form_header">Receiver Details</td>
          </tr>
          <tr>
            <td width="30%"><strong>Name:</strong></td>
            <td width="70%"><%= Session("strReceiverName") %></td>
          </tr>
          <tr>
            <td><strong>Contact Person:</strong></td>
            <td><%= Session("strReceiverContact") %></td>
          </tr>
          <tr>
            <td><strong>Phone:</strong></td>
            <td><%= Session("strReceiverPhone") %></td>
          </tr>
          <tr>
            <td><strong>Address:</strong></td>
            <td><%= Session("strReceiverAddress") %></td>
          </tr>
          <tr>
            <td><strong>City:</strong></td>
            <td><%= Session("strReceiverCity") %></td>
          </tr>
          <tr>
            <td><strong>State:</strong></td>
            <td><%= Session("strReceiverState") %></td>
          </tr>
          <tr>
            <td><strong>Postcode:</strong></td>
            <td><%= Session("intReceiverPostcode") %></td>
          </tr>
          <tr>
            <td valign="top"><strong>Comments:</strong></td>
            <td><%= Session("strReceiverComments") %></td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td colspan="2" style="padding-top:20px;"><div class="alert alert-search">
          <form name="frmSearch" id="frmSearch" action="confirm_freight.asp?type=search" method="post" onsubmit="searchProduct()">
            <h2>Search Product:</h2>
            Product code / Category
            <input type="text" name="txtSearch" size="25" value="<%= request("keyword") %>" maxlength="20" />
            <input type="button" name="btnSearch" value="Search" onclick="searchProduct()" />
            <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
          </form>
        </div>
        <table cellspacing="0" cellpadding="5" class="form_box_nowidth" border="0" width="100%">
          <tr>
            <td class="form_header">Category</td>
            <td class="form_header">Name</td>
            <td class="form_header">Description</td>
            <td class="form_header">Weight (kg)</td>
            <td class="form_header">Width (cm)</td>
            <td class="form_header">Height (cm)</td>
            <td class="form_header">Depth (cm)</td>
            <td class="form_header">Qty</td>
            <td class="form_header">Pallet</td>
            <td class="form_header"></td>
          </tr>
          <%= strProductList %>
        </table>
        <h2>Added Products</h2>
        <table cellspacing="0" cellpadding="5" class="form_box_nowidth" border="0" width="100%">
          <tr>
            <td class="form_header">Name</td>
            <td class="form_header">Description</td>
            <td class="form_header">Qty</td>
            <td class="form_header">Pallet</td>
            <td class="form_header">Weight (kg)</td>
            <td class="form_header">Width (cm)</td>
            <td class="form_header">Height (cm)</td>
            <td class="form_header">Depth (cm)</td>
          </tr>
          <%= strDisplayList %>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>