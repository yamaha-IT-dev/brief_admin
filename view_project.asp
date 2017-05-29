<%
session.lcid = 2057

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsComment.asp" -->
<!--#include file="class/clsEmployee.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>View Project</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/generic_form_validations.js"></script>
<script type="text/javascript" src="include/usableforms.js"></script>
<script language="JavaScript" type="text/javascript">
function completeProgress(form) {
	if (form.cboProgress.value == "100") {
		alert("100%!");
		form.cboStatus.value == "0";
	} else {	
		form.cboStatus.value == "";
	}
}

function validateFormOnSubmit(theForm) {
	var reason = "";
	var blnSubmit = true;	
	
	reason += validateEmptyField(theForm.txtTitle,"Title");		
	reason += validateEmptyField(theForm.txtGLcode,"GL Code");

	if (blnSubmit == true){
        theForm.Action.value = 'Update';		
		return true;
    }
}

function submitComment(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	reason += validateEmptyField(theForm.txtComment,"Comment");
	
	if (reason != "") {
    	alert("Some fields need correction:\n" + reason);

		blnSubmit = false;
		return false;
  	}
	
	if (blnSubmit == true){
		theForm.Action.value = 'Comment';
		
		return true;		
    }
}
</script>
<%
'-----------------------------------------------
' GET PROJECT
'-----------------------------------------------
Sub getProject
	dim intID
	intID = request("id")

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSqlQuery = "SELECT * FROM yma_project WHERE project_id = " & intID

	rs.Open strSqlQuery, conn

	'response.write strSqlQuery

    if not DB_RecSetIsEmpty(rs) Then
		session("project_title") 			= rs("project_title")
		session("project_gl_code") 			= rs("project_gl_code")
		session("project_contact") 			= rs("project_contact")
		session("project_output_printed") 	= rs("project_output_printed")
		session("project_output_ad") 		= rs("project_output_ad")
		session("project_output_web") 		= rs("project_output_web")
		session("project_output_interactive")= rs("project_output_interactive")
		session("project_output_details") 	= rs("project_output_details")
		session("project_deadline") 		= rs("project_deadline")
		session("project_image_location") 	= rs("project_image_location")
		session("project_copy_location") 	= rs("project_copy_location")		
		session("project_description") 		= rs("project_description")
		session("project_quote") 			= rs("project_quote")
		session("project_progress") 		= rs("project_progress")
		session("project_comments") 		= rs("project_comments")
		session("project_date_created") 	= rs("project_date_created")
		session("project_created_by") 		= rs("project_created_by")
		session("project_date_modified") 	= rs("project_date_modified")
		session("project_modified_by") 		= rs("project_modified_by")
		session("project_status") 			= rs("project_status")		
    end if

    call CloseDataBase()

end sub
'-----------------------------------------------
' UPDATE PROJECT
'-----------------------------------------------
sub updateProject	
	dim strSQL
	dim intID
	intID = request("id")
	
	dim strTitle
	dim strEmail
	dim intOutputPrinted
	dim intOutputAd
	dim intOutputWeb
	dim intOutputInteractive
	dim strOutputDetails
	dim strDeadline
	dim strImagesLocation
	dim strCopyLocation
	dim strDescription
	dim strGLcode
	
	dim intStatus
	dim strUsername
			
	strTitle 				= Replace(Request.Form("txtTitle"),"'","''")
	strEmail  				= Trim(Request.Form("txtEmail"))
	intOutputPrinted 		= request("chkOutputPrinted")
	intOutputAd 			= request("chkOutputAd")
	intOutputWeb 			= request("chkOutputWeb")
	intOutputInteractive 	= request("chkOutputInteractive")
	strOutputDetails 		= Replace(Request.Form("txtOutputDetails"),"'","''")
	strDeadline				= Trim(Request.Form("txtDeadline"))	
	strImagesLocation 		= Replace(Request.Form("txtImagesLocation"),"'","''")
	strCopyLocation 		= Replace(Request.Form("txtCopyLocation"),"'","''")
	strDescription 			= Replace(Request.Form("txtDescription"),"'","''")
	strQuote  				= Trim(Request.Form("txtQuote"))
	intProgress  			= Trim(Request.Form("cboProgress"))
	strGLcode  				= Trim(Request.Form("txtGLcode"))
	intStatus				= Trim(Request.form("cboStatus"))
	strUsername 			= LCASE(Request.ServerVariables("REMOTE_USER"))
	
	Call OpenDataBase()
	
	strSQL = "UPDATE yma_project SET "
	strSQL = strSQL & "project_title = '" & Server.HTMLEncode(strTitle) & "',"
	strSQL = strSQL & "project_gl_code = '" & Server.HTMLEncode(strGLcode) & "',"
	strSQL = strSQL & "project_output_printed = '" & intOutputPrinted & "',"
	strSQL = strSQL & "project_output_ad = '" & intOutputAd & "',"
	strSQL = strSQL & "project_output_web = '" & intOutputWeb & "',"
	strSQL = strSQL & "project_output_interactive = '" & intOutputInteractive & "',"
	strSQL = strSQL & "project_output_details = '" & Server.HTMLEncode(strOutputDetails) & "',"
	strSQL = strSQL & "project_deadline = CONVERT(datetime,'" & strDeadline & "',103),"
	strSQL = strSQL & "project_image_location = '" & strImagesLocation & "',"
	strSQL = strSQL & "project_copy_location = '" & strCopyLocation & "',"
	strSQL = strSQL & "project_description = '" & Server.HTMLEncode(strDescription) & "',"
	strSQL = strSQL & "project_date_modified = getdate(),"
	strSQL = strSQL & "project_modified_by = '" & strUsername & "' WHERE project_id = " & intID
	
	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if
	
	Call CloseDataBase()
end sub

sub main
	dim intID
	intID 	= request("id")
	
	call getProject
	call listComments(intID,projectModuleID)
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		select case Trim(Request.Form("Action"))
			case "Update"
				call updateProject
				call getProject
			case "Comment"
				call addComment(intID,projectModuleID)
				call listComments(intID,projectModuleID)
		end select
	end if
end sub

call main

dim strMessageText
dim strCommentsList
%>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
<!-- #include file="include/header.asp" -->
  <tr>
    <td valign="top" class="maincontent">
      <h2>View Project</h2>
      <table cellpadding="4" cellspacing="0" class="created_table">
        <tr>
          <td class="created_column_1"><strong>Created:</strong></td>
          <td class="created_column_2"><%= session("project_created_by") %></td>
          <td class="created_column_3"><%= displayDateFormatted(session("project_date_created")) %></td>
        </tr>
        <tr>
          <td><strong>Last modified:</strong></td>
          <td><%= session("project_modified_by") %></td>
          <td><%= displayDateFormatted(session("project_date_modified")) %></td>
        </tr>
      </table>
      <p><font color="red"><%= strMessageText %></font></p>
      <table cellspacing="0" cellpadding="0" width="1200">
        <tr>
          <td width="45%" style="padding-right:15px;"><form action="" method="post" name="form_update_project" id="form_update_project" onsubmit="return validateFormOnSubmit(this)">
              <table cellpadding="5" cellspacing="0" class="item_maintenance_box">
                <tr>
                  <td colspan="5" class="item_maintenance_header">Brief Details</td>
                </tr>
                <tr>
                  <td width="70%" colspan="5">Title<span class="mandatory">*</span>:<br />
                    <input name="txtTitle" type="text" id="txtTitle" size="70" maxlength="60" value="<%= session("project_title") %>" /></td>
                </tr>
                <tr>
                  <td colspan="5">Requested by: <a href="mailto:<%= session("project_created_by") %>"><%= session("project_created_by") %></a></td>
                </tr>
                <tr>
                  <td width="20%">Output:</td>
                  <td width="20%"><input type="checkbox" name="chkOutputPrinted" id="chkOutputPrinted" value="1" <% if session("project_output_printed") = "1" then Response.Write " checked" end if%> />
                    Printed</td>
                  <td width="20%"><input type="checkbox" name="chkOutputAd" id="chkOutputAd" value="1" <% if session("project_output_ad") = "1" then Response.Write " checked" end if%> />
                    Advertising</td>
                  <td width="20%"><input type="checkbox" name="chkOutputWeb" id="chkOutputWeb" value="1" <% if session("project_output_web") = "1" then Response.Write " checked" end if%> />
                    Web</td>
                  <td width="20%"><input type="checkbox" name="chkOutputInteractive" id="chkOutputInteractive" value="1" <% if session("project_output_interactive") = "1" then Response.Write " checked" end if%> />
                    Interactive</td>
                </tr>
                <tr>
                  <td colspan="5">Output Details:<br />
                    <input type="text" id="txtOutputDetails" name="txtOutputDetails" maxlength="70" size="80" value="<%= session("project_output_details") %>" /></td>
                </tr>
                <tr>
                  <td colspan="2">Deadline:<br />
                    <input type="text" id="txtDeadline" name="txtDeadline" maxlength="10" size="10" value="<%= session("project_deadline") %>" /></td>
                  <td colspan="3">GL code<span class="mandatory">*</span>:<br />
                    <input type="text" id="txtGLcode" name="txtGLcode" maxlength="20" size="30" value="<%= session("project_gl_code") %>" /></td>
                </tr>
                <tr>
                  <td colspan="5">Image(s) supplied location:<br />
                    <input type="text" id="txtImagesLocation" name="txtImagesLocation" maxlength="80" size="90" value="<%= session("project_image_location") %>" /></td>
                </tr>
                <tr>
                  <td colspan="5">Copy supplied location:<br />
                    <input type="text" id="txtCopyLocation" name="txtCopyLocation" maxlength="80" size="90" value="<%= session("project_copy_location") %>" /></td>
                </tr>
                <tr>
                  <td colspan="5">Description:<br />
                    <textarea name="txtDescription" id="txtDescription" cols="70" rows="4"><%= session("project_description") %></textarea></td>
                </tr>
                <tr>
                  <td colspan="2">&nbsp;</td>
                  <td colspan="3">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2">Quote: <%= session("project_quote") %> hour(s)</td>
                  <td colspan="3">Progress: <%= session("project_progress") %> %</td>
                </tr>
                <tr class="status_row">
                  <td colspan="5">Status: 
				  <% select case session("project_status")
						case 1
							Response.Write("Open")
						case 2
							Response.Write("On-hold")
						case 0
							Response.Write("Completed")
					end select	
				  %>
				  </td>
                </tr>
              </table>
              <p>
                <input type="hidden" name="Action" />
                <input type="submit" value="Update Project" />
              </p>
            </form></td>
          <td valign="top" width="55%"><h3>Comments<br />
              <img src="images/comment_bar.jpg" border="0" /></h3>
            <table cellpadding="5" cellspacing="0" border="0" class="comments_box">
              <%= strCommentsList %>
              <tr>
                <td><form action="" name="form_add_comment" id="form_add_comment" method="post" onsubmit="return submitComment(this)">
                    <p>
                      <input type="text" name="txtComment" id="txtComment" maxlength="60" size="65" />
                      <input type="hidden" name="Action" />
                      <input type="submit" value="Add Comment" />
                    </p>
                  </form></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>