<%
'-----------------------------------------------
' GET LOAN ACCOUNT DEPARTMENT (using account code)
'-----------------------------------------------
Function getLoanAccountDepartment(strAccountCode)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT emp_department FROM yma_employee "
	strSQL = strSQL & " WHERE emp_initial = '" & strAccountCode & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("acc_department")	= rs("emp_department")
    end if

    call CloseDataBase()
end function
%>