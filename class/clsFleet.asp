<%
'----------------------------------------------------------------------------------------
' ADD RESOURCE TYPE
'----------------------------------------------------------------------------------------
Function addFleet
	dim strSQL
	
	call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_fleet ("
	strSQL = strSQL & "regNo,"
	strSQL = strSQL & "regDriver,"
	strSQL = strSQL & "regDepartment,"
	strSQL = strSQL & "regState,"
	strSQL = strSQL & "regStart,"
	strSQL = strSQL & "regEnd,"
	strSQL = strSQL & "regLimit,"
	strSQL = strSQL & "regTyres,"
	strSQL = strSQL & "regUsed,"
	strSQL = strSQL & "regMaintenance,"
	strSQL = strSQL & "regCreatedBy"
	strSQL = strSQL & ")VALUES( "	
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("txtRego"))) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("txtStaff"))) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("cboDepartment"))) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("cboState"))) & "',"
	strSQL = strSQL & "CONVERT(DateTime,'" & Trim(Request.Form("txtStartDate")) & "',103),"
	strSQL = strSQL & "CONVERT(DateTime,'" & Trim(Request.Form("txtEndDate")) & "',103),"
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("txtOdoLimit"))) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("cboTyres"))) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("cboUsed"))) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(Trim(Request.Form("cboMaintenance"))) & "',"
	strSQL = strSQL & "'" & session("logged_username") & "')"
	
	response.Write strSQL	
	  
	on error resume next
	conn.Execute strSQL
	
	On error Goto 0
	
	if err <> 0 then
		strMessageText = err.description
	else	
		strMessageText = "ADDED"
		'Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	end if 
	
	call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE FLEET
'----------------------------------------------------------------------------------------
Function updateFleet(intID, strModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_fleet SET "
	strSQL = strSQL & "regNo = '" & Server.HTMLEncode(Trim(Request.Form("txtNo"))) & "',"
	strSQL = strSQL & "regDriver = '" & Server.HTMLEncode(Trim(Request.Form("txtDriver"))) & "',"
	strSQL = strSQL & "regDepartment = '" & Server.HTMLEncode(Trim(Request.Form("txtDepartment"))) & "',"
	strSQL = strSQL & "regState = '" & Server.HTMLEncode(Trim(Request.Form("txtState"))) & "',"
	strSQL = strSQL & "regStart = CONVERT(DateTime,'" & Trim(Request.Form("txtStart")) & "',103),"
	strSQL = strSQL & "regEnd = CONVERT(DateTime,'" & Trim(Request.Form("txtEnd")) & "',103),"
	strSQL = strSQL & "regLimit = '" & Server.HTMLEncode(Trim(Request.Form("txtLimit"))) & "',"
	strSQL = strSQL & "regTyres = '" & Server.HTMLEncode(Trim(Request.Form("txtTyres"))) & "',"
	strSQL = strSQL & "regUsed = '" & Server.HTMLEncode(Trim(Request.Form("txtUsed"))) & "',"
	strSQL = strSQL & "regMaintenance = '" & Server.HTMLEncode(Trim(Request.Form("cboMaintenance"))) & "',"
	strSQL = strSQL & "regStatus = '" & Trim(Request.Form("cboStatus")) & "',"
	strSQL = strSQL & "regDateModified = GetDate(),"
	strSQL = strSQL & "regModifiedBy = '" & strModifiedBy & "' WHERE regID = " & intID

	response.Write strSQL
	
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		'strMessageText = "The record has been updated."
	end if

	Call CloseDataBase()
end function

%>