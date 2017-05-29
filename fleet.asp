<% strSection = "marketing" %>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsFleet.asp " -->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Fleet Management</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/javascript.js"></script>
<script>
function searchProduct(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strStatus 		= document.forms[0].cboStatus.value;
    document.location.href = 'fleet.asp?type=search&txtSearch=' + strSearch + '&status=' + strStatus;	
}
    
function resetSearch(){
	document.location.href = 'fleet.asp?type=reset';    
}

function validateAddForm(theForm) {
	var reason 		= "";
	var blnSubmit 	= true;
	
	reason += validateEmptyField(theForm.txtRego,"Rego");
	reason += validateSpecialCharacters(theForm.txtRego,"Rego");
	
	reason += validateEmptyField(theForm.txtStaff,"Staff");
	reason += validateSpecialCharacters(theForm.txtStaff,"Staff");
	
	reason += validateNumeric(theForm.txtOdoLimit,"Limit");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);
    	
		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.action.value = 'add';
		
		return true;
    }
}

function validateUpdateForm(theForm) {
	var reason = "";
	var blnSubmit = true;

	reason += validateEmptyField(theForm.txtNo,"Rego");	
	reason += validateSpecialCharacters(theForm.txtNo,"Rego");
	
	reason += validateEmptyField(theForm.txtDriver,"Driver");
	reason += validateSpecialCharacters(theForm.txtDriver,"Driver");
	
  	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}

	if (blnSubmit == true){
        theForm.action.value = 'update';
  		theForm.submit();

		return true;
    }
}

</script>
</head>
<body style="padding-top:20px; padding-left:20px;">
<%
sub setSearch	
	select case trim(request("type"))
		case "reset" 
			Session("fleet_search") 		= ""	
			Session("fleet_status") 		= ""
			Session("fleet_initial_page") = 1
		case "search"
			Session("fleet_search") 		= server.htmlencode(trim(Request("txtSearch")))
			Session("fleet_status") 		= trim(request("status"))
			Session("fleet_initial_page") = 1
	end select
end sub

sub displayType
	dim strSQL
	
	dim intRecordCount
	
	dim iRecordCount
	iRecordCount = 0		
	
	dim strTodayDate	
	strTodayDate = FormatDateTime(Date())
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 900
	
	strSQL = "SELECT * FROM tbl_fleet "
	strSQL = strSQL & " WHERE "
	strSQL = strSQL & "		regNo LIKE '%" & Session("fleet_search") & "%' "
	strSQL = strSQL & "	 	AND regStatus LIKE '%" & Session("fleet_status") & "%' "
	strSQL = strSQL & " ORDER BY regDriver"
	
	'response.write strSQL & "<br>"
	
	rs.Open strSQL, conn
	
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("action")
	    case "<<"
		    intpage = 1
			Session("fleet_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			Session("fleet_initial_page") = intpage
			
			if Session("fleet_initial_page") < 1 then Session("fleet_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			Session("fleet_initial_page") = intpage
			
			if Session("fleet_initial_page") > intPageCount then Session("fleet_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			Session("fleet_initial_page") = intpage
    end select

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	    rs.AbsolutePage = Session("fleet_initial_page")  
	
		For intRecord = 1 To rs.PageSize
			strDisplayList = strDisplayList & "<form method=""post"" name=""form_update"" id=""form_update"" onsubmit=""return validateUpdateForm(this)"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""action"" value=""update"">"		
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""id"" value=""" & rs("regID") & """>"
			
			
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				else
					strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
				end if
			'strDisplayList = strDisplayList & "<td>" & rs("regID") & "</td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtName"" name=""txtNo"" maxlength=""7"" size=""6"" value=""" & rs("regNo") & """ ></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtDriver"" name=""txtDriver"" maxlength=""30"" size=""30"" value=""" & rs("regDriver") & """ required></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtDepartment"" name=""txtDepartment"" maxlength=""4"" size=""4"" value=""" & rs("regDepartment") & """ required></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtState"" name=""txtState"" maxlength=""3"" size=""3"" value=""" & rs("regState") & """ required></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtStart"" name=""txtStart"" maxlength=""12"" size=""12"" value=""" & rs("regStart") & """ placeholder=""DD/MM/YYYY"" required></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtStart"" name=""txtEnd"" maxlength=""12"" size=""12"" value=""" & rs("regEnd") & """ placeholder=""DD/MM/YYYY"" required></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtTyres"" name=""txtLimit"" maxlength=""6"" size=""6"" value=""" & rs("regLimit") & """ required></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtTyres"" name=""txtTyres"" maxlength=""1"" size=""1"" value=""" & rs("regTyres") & """ required></td>"
			strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtUsed"" name=""txtUsed"" maxlength=""1"" size=""1"" value=""" & rs("regUsed") & """ required></td>"
			'strDisplayList = strDisplayList & "<td><input type=""text"" id=""txtMaintenance"" name=""txtMaintenance"" maxlength=""1"" size=""1"" value=""" & rs("regMaintenance") & """ required></td>"
			strDisplayList = strDisplayList & "<td><select name=""cboMaintenance"">"            
			if rs("regMaintenance") = 1 then
				strDisplayList = strDisplayList & "<option value=""1"" selected>Yes</option><option value=""0"" style=""color:red"">No</option>"
			else
				strDisplayList = strDisplayList & "<option value=""1"">Yes</option><option value=""0"" selected style=""color:red"">No</option>"
			end if
			strDisplayList = strDisplayList & " </select></td>"
			strDisplayList = strDisplayList & "<td><select name=""cboStatus"">"            
			if rs("regStatus") = 1 then
				strDisplayList = strDisplayList & "<option value=""1"" selected>Active</option><option value=""0"" style=""background-color:#FFFF00"">In-active</option>"
			else
				strDisplayList = strDisplayList & "<option value=""1"">Active</option><option value=""0"" selected style=""background-color:#FFFF00"">In-active</option>"
			end if
			strDisplayList = strDisplayList & " </select></td>"		
			strDisplayList = strDisplayList & "<td><input type=""hidden"" id=""myHidden"" /><input type=""submit"" value=""Update"" /></td>"
			strDisplayList = strDisplayList & "<td>" & rs("regCreatedBy") & " - " & rs("regDateCreated") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("regModifiedBy") & " - " & rs("regDateModified") & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""14"">No records found.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""14"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Total: " & intRecordCount & "</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		Select Case Trim(Request("action"))
			case "add"		
				call addFleet
			case "update"
				dim intID
				intID = Trim(Request("id"))
				
				call updateFleet(intID, session("logged_username"))				
		end select
	end if
	
	call setSearch
	
	if trim(Session("fleet_initial_page")) = "" then
    	Session("fleet_initial_page") = 1
	end if
    
    call displayType
end sub

call main

dim rs, intPageCount, intpage, intRecord, strDisplayList, strMessageText
%>
<table>
  <tr>
    <td><h1>Fleet Management</h1>
      <div class="alert alert-search">
        <form name="frmSearch" id="frmSearch" action="fleet.asp?type=search" method="post" onsubmit="searchProduct()">
          <input type="text" name="txtSearch" size="25" value="<%= Session("fleet_search") %>" maxlength="20" placeholder="Search Name" />
          <select name="cboStatus" onchange="searchProduct()">
            <option <% if Session("fleet_status") = ""  then Response.Write " selected" end if%> value="">All status</option>
            <option <% if Session("fleet_status") = "1" then Response.Write " selected" end if%> value="1" style="color:green">Active</option>
            <option <% if Session("fleet_status") = "0" then Response.Write " selected" end if%> value="0" style="color:red">In-active</option>
          </select>
          <input type="button" name="btnSearch" value="Search" onclick="searchProduct()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </form>
      </div>
      <table cellspacing="0" cellpadding="5">
        <tr class="innerdoctitle">          
          <td>Rego</td>
          <td>Driver</td>
          <td>Dept</td>
          <td>State</td>
          <td>Start</td>
          <td>End</td>
          <td>Limit</td>
          <td>Tyres</td>
          <td>Used</td>
          <td>Maintenance</td>
          <td>Status</td>
          <td></td>
          <td>Created</td>
          <td>Modified</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
    <td valign="top"><!--<h2>Add New Fleet</h2>
      <form action="" method="post" name="form_add" id="form_add" onsubmit="return validateAddForm(this)">
        <p><strong>Rego:</strong><br />
          <input type="text" id="txtRego" name="txtRego" maxlength="30" size="30" required />
        </p>
        <p><strong>Driver:</strong><br />
          <input type="text" id="txtStaff" name="txtStaff" maxlength="30" size="30" required />
        </p>
        <p>Department:<br>
          <select name="cboDepartment">
            <option value="AV">AV</option>
            <option value="MPD">MPD</option>
            <option value="Other">Other</option>
          </select>
        </p>
        <p>State:<br>
          <select name="cboState">
            <option value="VIC">VIC</option>
            <option value="NSW">NSW</option>
            <option value="ACT">ACT</option>
            <option value="QLD">QLD</option>
            <option value="NT">NT</option>
            <option value="WA">WA</option>
            <option value="SA">SA</option>
            <option value="TAS">TAS</option>
          </select>
        </p>
        <p><strong>Start:</strong><br />
          <input type="text" id="txtStartDate" name="txtStartDate" maxlength="12" size="20" placeholder="DD/MM/YYYY" required />
        </p>
        <p><strong>End:</strong><br />
          <input type="text" id="txtEndDate" name="txtEndDate" maxlength="12" size="20" placeholder="DD/MM/YYYY" required />
        </p>
        <p><strong>Limit:</strong><br />
          <input type="text" id="txtOdoLimit" name="txtOdoLimit" maxlength="6" size="8" required />
        </p>
        <p><strong>Tyres:</strong><br />
          <select name="cboState">
            <option value="0">0</option>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>            
          </select>
        </p>
        <p><strong>Used:</strong><br />
          <select name="cboUsed">
            <option value="0">0</option>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>            
          </select>
        </p>
        <p><strong>Maintenance:</strong><br />
          <select name="cboMaintenance">
            <option value="0">No</option>
            <option value="1">Yes</option>                  
          </select>
        </p>
        <p>
          <input type="hidden" name="action" />
          <input type="submit" value="Submit" />
        </p>
      </form>--></td>
  </tr>
</table>
</body>
</html>