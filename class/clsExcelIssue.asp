<%
'-----------------------------------------------
' ADD ISSUE
'-----------------------------------------------
function addIssue(issASC, issContactName, issProduct, issReportedFault, issDiagnosedFault, issReason, issReturnDate, issReturnConnote, issComments, issSpareParts, issDispatchDate, issDispatchConnote, issStatus, issCreatedBy)
	dim strSQL

	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_excel_issue (issASC, issContactName, issProduct, issReportedFault, issDiagnosedFault, issReason, issReturnDate, issReturnConnote, issComments, issSpareParts, issDispatchDate, issDispatchConnote, issCreatedBy) VALUES ("
	strSQL = strSQL & " '" & Server.HTMLEncode(issASC) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(issContactName) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(issProduct) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(issReportedFault) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(issDiagnosedFault) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(issReason) & "',"
	strSQL = strSQL & " CONVERT(DateTime,'" & issReturnDate & "',103),"
	strSQL = strSQL & " '" & Server.HTMLEncode(issReturnConnote) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(issComments) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(issSpareParts) & "',"
	strSQL = strSQL & " CONVERT(DateTime,'" & issDispatchDate & "',103),"
	strSQL = strSQL & " '" & Server.HTMLEncode(issDispatchConnote) & "',"
	strSQL = strSQL & " '" & issCreatedBy & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else				
		strMessageText = "The record has been added."
	end if 
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' GET ISSUE
'-----------------------------------------------
Function getIssue(intID)
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_excel_issue WHERE issID = " & intID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("issASC") 				= rs("issASC")
		session("issContactName") 		= rs("issContactName")
		session("issProduct") 			= rs("issProduct")
		session("issReportedFault") 	= rs("issReportedFault")
		session("issDiagnosedFault") 	= rs("issDiagnosedFault")		
		session("issReason") 			= rs("issReason")
		session("issReturnDate") 		= rs("issReturnDate")
		session("issReturnConnote") 	= rs("issReturnConnote")
		session("issComments") 			= rs("issComments")
		session("issSpareParts") 		= rs("issSpareParts")
		session("issDispatchDate") 		= rs("issDispatchDate")
		session("issDispatchConnote") 	= rs("issDispatchConnote")
		session("issStatus") 			= rs("issStatus")
		session("issDateCreated") 		= rs("issDateCreated")
		session("issCreatedBy") 		= rs("issCreatedBy")
		session("issDateModified") 		= rs("issDateModified")
		session("issModifiedBy") 		= rs("issModifiedBy")			
    end if

    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE ISSUE
'----------------------------------------------------------------------------------------
Function updateIssue(intID, issASC, issContactName, issProduct, issReportedFault, issDiagnosedFault, issReason, issReturnDate, issReturnConnote, issComments, issSpareParts, issDispatchDate, issDispatchConnote, issStatus, issModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_excel_issue SET "
	strSQL = strSQL & "issASC = '" & Server.HTMLEncode(issASC) & "',"
	strSQL = strSQL & "issContactName = '" & Server.HTMLEncode(issContactName) & "',"
	strSQL = strSQL & "issProduct = '" & Server.HTMLEncode(issProduct) & "',"
	strSQL = strSQL & "issReportedFault = '" & Server.HTMLEncode(issReportedFault) & "',"
	strSQL = strSQL & "issDiagnosedFault = '" & Server.HTMLEncode(issDiagnosedFault) & "',"
	strSQL = strSQL & "issReason = '" & Server.HTMLEncode(issReason) & "',"
	strSQL = strSQL & "issReturnDate = CONVERT(DateTime,'" & issReturnDate & "',103),"
	strSQL = strSQL & "issReturnConnote = '" & Server.HTMLEncode(issReturnConnote) & "',"
	strSQL = strSQL & "issComments = '" & Server.HTMLEncode(issComments) & "',"
	strSQL = strSQL & "issSpareParts = '" & Server.HTMLEncode(issSpareParts) & "',"
	strSQL = strSQL & "issDispatchDate = CONVERT(DateTime,'" & issDispatchDate & "',103),"
	strSQL = strSQL & "issDispatchConnote = '" & Server.HTMLEncode(issDispatchConnote) & "',"
	strSQL = strSQL & "issStatus = '" & Server.HTMLEncode(issStatus) & "',"
	strSQL = strSQL & "issDateModified = GetDate(),"
	strSQL = strSQL & "issModifiedBy = '" & issModifiedBy & "' WHERE issID = " & intID

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if

	Call CloseDataBase()
end function
%>