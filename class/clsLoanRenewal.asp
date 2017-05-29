<%
'-----------------------------------------------
' ADD LOAN STOCK RENEWAL - LIST
'-----------------------------------------------
function addRenewal(renOrderNo, renOrderLine, renAccountCode, renAccountName, renDepartment, renItemCode, renSerialNo, renLocation, renLIC, renLoanDate, renCreatedBy, renCreatedByEmail)
	dim strSQL
	dim strRecipient
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())
	
	dim strExpiryDate
	strExpiryDate 	= DateAdd("m", 3, strTodayDate)
	
	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_loan_renewal (renOrderNo, renOrderLine,  renAccountCode, renAccountName, renDepartment, renItemCode, renSerialNo, renLocation, renLIC, renLoanDate, renExpiryDate, renCreatedBy, renCreatedByEmail) VALUES ("
	strSQL = strSQL & "'" & Server.HTMLEncode(renOrderNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renOrderLine) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renAccountCode) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renAccountName) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renDepartment) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renItemCode) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renSerialNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renLocation) & "',"
	strSQL = strSQL & "'" & renLic & "',"
	strSQL = strSQL & " CONVERT(datetime,'" & renLoanDate & "',103),"
	strSQL = strSQL & " CONVERT(datetime,'" & strExpiryDate & "',103),"
	strSQL = strSQL & "'" & renCreatedBy & "',"
	strSQL = strSQL & "'" & renCreatedByEmail & "')"
	
	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		Set oMail = Server.CreateObject("CDO.Message")
		Set iConf = Server.CreateObject("CDO.Configuration")
		Set Flds = iConf.Fields
		
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update
		
		emailFrom 	= "automailer@music.yamaha.com"
		
		if renDepartment = "AV" then
			emailTo 	 = "simon.goldsworthy@music.yamaha.com"
			emailCc 	 = Trim(renCreatedByEmail)
			emailSubject = "New AV Loan Stock Renewal"
			strRecipient = "Simon"
		elseif renDepartment = "MPD" then
			emailTo 	 = "michael.shade@music.yamaha.com"
			emailCc 	 = Trim(renCreatedByEmail)
			emailSubject = "New MPED Loan Stock Renewal"
			strRecipient = "Michael"
		else
			emailTo 	 = "tasi.herbu@music.yamaha.com"
			emailCc 	 = Trim(renCreatedByEmail)
			emailSubject = "New O&F Loan Stock Renewal"
			strRecipient = "Tasi"
		end if
		
		emailBodyText =	"Hi " & strRecipient & "," & vbCrLf _
						& " " & vbCrLf _
						& "There is a new Loan Stock Renewal Request that needs your approval." & vbCrLf _
						& " " & vbCrLf _
						& "Account   : " & renAccountCode & vbCrLf _
						& "Name      : " & renAccountName & vbCrLf _
						& "Product   : " & renItemCode & vbCrLf _
						& "Serial    : " & renSerialNo & vbCrLf _
						& "Location  : " & renLocation & vbCrLf _
						& "LIC $     : " & FormatNumber(renLIC) & vbCrLf _
						& "Loan date : " & renLoanDate & vbCrLf _
						& "Expiry    : " & strExpiryDate & vbCrLf _
						& " " & vbCrLf _
						& "Please click on the below link to approve it:" & vbCrLf _
						& "http://intranet:96/loan_renewal.asp?type=search&txtSearch=&department=" & renDepartment & "&user=&status=1&sort=latest" & vbCrLf _	
						& " " & vbCrLf _
						& "Thank you. (This is an automated email - please do not reply to this email)"
						
		Set oMail.Configuration = iConf
		
		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send
				
		Set iConf = Nothing
		Set Flds = Nothing
		
		'response.Redirect("loan_user.asp?account=" & session("loan_user_account"))
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Renewal Request has been submitted.</div>"
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOAN STOCK RENEWAL - STORED PROC!!
'-----------------------------------------------
function addLoanRenewal(renOrderNo, renOrderLine, renAccountCode, renAccountName, renDepartment, renItemCode, renSerialNo, renLocation, renLIC, renLoanDate, renCreatedBy, renCreatedByEmail)
	Dim cmdObj, paraObj
	
    call OpenDataBase
	
    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spAddLoanLocation"
    cmdObj.CommandType = AdCmdStoredProc
		
	Set paraObj = cmdObj.CreateParameter("@stockOrderNo",AdInteger,AdParamInput,4, stockOrderNo)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@stockOrderLine",AdInteger,AdParamInput,4, stockOrderLine)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@stockLocation",AdVarChar,AdParamInput,120, stockLocation)
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter("@stockCreatedBy",AdVarChar,AdParamInput,50, stockCreatedBy)
	cmdObj.Parameters.Append paraObj	
	
    On Error Resume Next
        Dim rs
        Dim id
        set rs = cmdObj.Execute
        id = rs(0)
        set rs = nothing
    On error Goto 0
	
    if CheckForSQLError(conn,"Add",MessageText) = TRUE then
        addLoanRenewal = FALSE
        strMessageText = MessageText
		'strMessageText = err.description
    else
		addLoanRenewal = TRUE
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Location has been saved.</div>"
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)
	
    call CloseDataBase
end function

'-----------------------------------------------
' ADD LOAN STOCK RENEWAL
'-----------------------------------------------
function addLoanStockRenewal(renOrderNo, renOrderLine, renComments, renCreatedBy, renCreatedByEmail)
	dim strSQL
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())
	
	dim strExpiryDate
	strExpiryDate 	= DateAdd("m", 3, strTodayDate)
	
	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_loan_renewal (renOrderNo, renOrderLine, renComments, renExpiryDate, renCreatedBy, renCreatedByEmail) VALUES ("
	strSQL = strSQL & "'" & Server.HTMLEncode(renOrderNo) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renOrderLine) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(renComments) & "',"
	strSQL = strSQL & " CONVERT(datetime,'" & strExpiryDate & "',103),"
	strSQL = strSQL & "'" & renCreatedBy & "',"
	strSQL = strSQL & "'" & renCreatedByEmail & "')"
	
	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The Renewal Request has been submitted.</div>"
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' LIST LOAN RENEWAL
'-----------------------------------------------
function listLoanRenewal(loanOrderNo, loanOrderLine)
    dim strSQL
	dim intRecordCount
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_loan_renewal "
	strSQL = strSQL & "	WHERE renOrderNo = '" & loanOrderNo & "' "
	strSQL = strSQL & "		AND renOrderLine = '" & loanOrderLine & "' "
	strSQL = strSQL & "	ORDER BY renDateCreated"
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	
	session("loanstock_renewal_record_count") = rs.recordcount
	
    strLoanRenewalList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
		For intRecord = 1 To rs.PageSize
			strLoanRenewalList = strLoanRenewalList & "<tr>"			
			strLoanRenewalList = strLoanRenewalList & "	<td align=""center"">" & FormatDateTime(Trim(rs("renDateCreated")),1) & "</td>"
			strLoanRenewalList = strLoanRenewalList & "	<td align=""center"">" & FormatDateTime(Trim(rs("renExpiryDate")),1) & "</td>"
			strLoanRenewalList = strLoanRenewalList & "	<td align=""center"">" & Trim(rs("renComments")) & "</td>"
			strLoanRenewalList = strLoanRenewalList & "	<td align=""center""><table><tr><td>"
			strLoanRenewalList = strLoanRenewalList & "		<form method=""post"" name=""form_approve"" id=""form_approve"" onsubmit=""return submitApproval(this)"">"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""hidden"" name=""action"" value=""Approve"">"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""hidden"" name=""renID"" value=""" & rs("renID") & """>"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""hidden"" name=""renCreatedByEmail"" value=""" & rs("renCreatedByEmail") & """>"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""submit"" value=""Approve"" style=""color:green"" />"
			strLoanRenewalList = strLoanRenewalList & "		</form></td>"
			strLoanRenewalList = strLoanRenewalList & "		<td><form method=""post"" name=""form_reject"" id=""form_reject"" onsubmit=""return submitRejection(this)"">"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""hidden"" name=""action"" value=""Reject"">"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""hidden"" name=""renID"" value=""" & rs("renID") & """>"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""hidden"" name=""renCreatedByEmail"" value=""" & rs("renCreatedByEmail") & """>"
			strLoanRenewalList = strLoanRenewalList & "			<input type=""submit"" value=""Reject"" style=""color:red"" />"
			strLoanRenewalList = strLoanRenewalList & "		</form></td></tr></table>"
			strLoanRenewalList = strLoanRenewalList & "	</td>"
			strLoanRenewalList = strLoanRenewalList & "</tr>"
			
			rs.movenext
			
			If rs.EOF Then Exit For
		next
	else
        strLoanRenewalList = "<tr><td colspan=""4"">&nbsp;</td></tr>"
	end if
	
	strLoanRenewalList = strLoanRenewalList & "<tr>"
	
    call CloseDataBase()
end function

'-----------------------------------------------
' GET LOAN STOCK RENEWAL
'-----------------------------------------------
Function getLoanStockRenewal(renOrderNo, renOrderLine)
    dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_ren WHERE renOrderNo = " & renOrderNo & " AND renOrderLine = " & renOrderLine

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then		
		session("renLocation") 		= rs("renLocation")
		session("renSerialNo") 		= rs("renSerialNo")
		session("renComments") 		= rs("renComments")		
		session("renStatus") 	   	= rs("renStatus")		
		session("renDateCreated")  	= rs("renDateCreated")
		session("renCreatedBy") 	= rs("renCreatedBy")
		session("renDateModified") 	= rs("renDateModified")
		session("renModifiedBy") 	= rs("renModifiedBy")	
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET LOAN STOCK RENEWAL
'-----------------------------------------------
function getNewLoanDate(renOrderNo, renOrderLine)
	dim strSQL
	
	'Call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT top 1 dateadd(m, -3, renExpiryDate) as newLoanDate FROM tbl_loan_renewal WHERE renOrderNo = " & renOrderNo & " AND renOrderLine = " & renOrderLine & " order by renID desc "
	
	rs.Open strSQL, conn
	
	if not DB_RecSetIsEmpty(rs) Then
		session("newLoanDate") =  rs("newLoanDate")
	end if
	
	'call CloseDataBase()
end function


'----------------------------------------------------------------------------------------
' UPDATE ISSUE
'----------------------------------------------------------------------------------------
Function updateLoanStockRenewal(renOrderNo, renOrderLine, renLocation, renSerialNo, renComments, renModifiedBy)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE tbl_loan_renewal SET "
	strSQL = strSQL & "renLocation = '" & Server.HTMLEncode(renLocation) & "',"
	strSQL = strSQL & "renSerialNo = '" & Server.HTMLEncode(renSerialNo) & "',"
	strSQL = strSQL & "renComments = '" & Server.HTMLEncode(renComments) & "',"	
	strSQL = strSQL & "renDateModified = GetDate(),"
	strSQL = strSQL & "renModifiedBy = '" & renModifiedBy & "' WHERE renOrderNo = " & renOrderNo & " AND renOrderLine = " & renOrderLine

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The Renewal Request has been updated.</div>"
	end if

	Call CloseDataBase()
end function
%>