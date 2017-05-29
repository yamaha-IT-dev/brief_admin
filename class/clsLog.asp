<%
'-----------------------------------------------
' LIST ALL LOGS
'-----------------------------------------------
function listLogs(intID,intTypeID)
    dim strSQL
	dim intRecordCount
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_log WHERE log_associated_id = '" & intID & "' AND log_type = '" & intTypeID & "' ORDER BY log_date"
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	

    strLogsList = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
	
		For intRecord = 1 To rs.PageSize
		
			strLogsList = strLogsList & "<tr><td width=""20%"">"	& trim(rs("log_username")) & "</td>"
			strLogsList = strLogsList & "<td width=""25%"">" & trim(rs("log_activity")) & "</td>"
			strLogsList = strLogsList & "<td width=""55%"">" & WeekDayName(WeekDay(rs("log_date"))) & ", " & FormatDateTime(rs("log_date"),1) & " at " & FormatDateTime(rs("log_date"),3) & "</td></tr>"
			
			rs.movenext
			
			If rs.EOF Then Exit For
		next
	else
        strLogsList = "<tr><td colspan=""3"">&nbsp;</td></tr>"
	end if
	
	strLogsList = strLogsList & "<tr>"
	
    call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOG
'-----------------------------------------------
function addLog(intID,intTypeID,strActivity)
	dim strSQL

	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_log (log_type, log_associated_id, log_username, log_activity) VALUES ("
	strSQL = strSQL & " '" & intTypeID & "',"
	strSQL = strSQL & " '" & intID & "',"
	strSQL = strSQL & " '" & session("logged_username") & "',"
	strSQL = strSQL & " '" & strActivity & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else				
		strMessageText = "This activity has been logged."
	end if 
	
	Call CloseDataBase()
end function
%>