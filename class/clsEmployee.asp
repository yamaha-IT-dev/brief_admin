<%
'-----------------------------------------------
' ADD NEW EMPLOYEE
'-----------------------------------------------
function addEmployee(strUsername,strFirstname,strLastname,strEmail,strDepartment,intAdmin,intManagerID,intStatus)
	dim strSQL
	
	Call OpenDataBase()
		
	strSQL = "INSERT INTO yma_employee ("
	strSQL = strSQL & "emp_username, "
	strSQL = strSQL & "emp_firstname, "
	strSQL = strSQL & "emp_lastname, "
	strSQL = strSQL & "emp_email, "
	strSQL = strSQL & "emp_department, "
	strSQL = strSQL & "emp_admin, "
	strSQL = strSQL & "emp_manager_id, "
	strSQL = strSQL & "emp_created_by "
	strSQL = strSQL & ") VALUES ("
	strSQL = strSQL & " '" & strUsername & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strFirstname) & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strLastname) & "',"
	strSQL = strSQL & " '" & strEmail & "',"
	strSQL = strSQL & " '" & strDepartment & "',"
	strSQL = strSQL & " '" & intAdmin & "',"
	strSQL = strSQL & " '" & intManagerID & "',"
	strSQL = strSQL & " '" & session("logged_username") & "' "
	strSQL = strSQL & ")"
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been added."
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' GET EMPLOYEE DETAILS (using username)
'-----------------------------------------------
Function getEmployeeDetails(strUsername)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_employee "
	strSQL = strSQL & " WHERE emp_username = '" & strUsername & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("employee_id") 		= rs("employee_id")
		session("emp_username") 	= rs("emp_username")
		session("emp_firstname") 	= rs("emp_firstname")
		session("emp_lastname") 	= rs("emp_lastname")
		session("emp_initial") 		= rs("emp_initial")
		session("emp_email") 		= rs("emp_email")
		session("emp_department")	= rs("emp_department")
		session("emp_status") 		= rs("emp_status")
		session("emp_admin") 		= rs("emp_admin")
		session("emp_manager_id") 	= rs("emp_manager_id")
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET EMPLOYEE DETAILS (using ID)
'-----------------------------------------------
Function getEmployee(intEmployeeID)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_employee "
	strSQL = strSQL & " WHERE employee_id = '" & intEmployeeID & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("employee_id") 			= rs("employee_id")
		session("employee_username") 	= rs("emp_username")
		session("employee_firstname") 	= rs("emp_firstname")
		session("employee_lastname") 	= rs("emp_lastname")
		session("employee_initial") 	= rs("emp_initial")
		session("employee_email") 		= rs("emp_email")
		session("employee_department")	= rs("emp_department")
		session("employee_status") 		= rs("emp_status")
		session("employee_admin") 		= rs("emp_admin")
		session("employee_manager_id") 	= rs("emp_manager_id")
		session("employee_date_created")= rs("emp_date_created")
		session("employee_created_by") 	= rs("emp_created_by")
		session("employee_date_modified")= rs("emp_date_modified")
		session("employee_modified_by") = rs("emp_modified_by")
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET EMPLOYEE LIST
'-----------------------------------------------
function getYourEmployeeList(intID)
    dim strSQL
    dim rs
	dim intEmployeeID
	dim strEmployeeUsername
	dim strEmployeeFirstname
	dim strEmployeeLastname
	
    call OpenDataBase()
    
	strSQL = "SELECT * FROM yma_employee "
	strSQL = strSQL & " WHERE emp_status = '1' "
	strSQL = strSQL & " 	AND emp_manager_id = '" & intID & "' "
	strSQL = strSQL & " ORDER BY emp_firstname"
	
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)
    
    strEmployeeList = strEmployeeList & "<option value=''>...</option>"
    
    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
			intEmployeeID			= rs("employee_id")
			strEmployeeUsername		= trim(rs("emp_username"))
        	strEmployeeFirstname 	= trim(rs("emp_firstname"))
			strEmployeeLastname 	= trim(rs("emp_lastname"))
			
            'if trim(session("employee_id")) = intEmployeeID then
            '    strEmployeeList = strEmployeeList & "<option selected value=" & intEmployeeID & ">" & strEmployeeFirstname & " " & strEmployeeLastname & "</option>"
            'else
                strEmployeeList = strEmployeeList & "<option value='" & intEmployeeID & "'>" & strEmployeeFirstname & " " & strEmployeeLastname & "</option>"
            'end if
            
        rs.Movenext
        loop
    end if
    
    call CloseDataBase()
end function

'-----------------------------------------------
' GET MANAGER LIST
'-----------------------------------------------
function getManagerList
    dim strSQL
    dim rs
	dim intEmployeeID
	dim strEmployeeFirstname
	dim strEmployeeLastname
	
    call OpenDataBase()
    
	strSQL = "SELECT * FROM yma_employee "
	strSQL = strSQL & " WHERE emp_status = '1' "
	strSQL = strSQL & " 	AND emp_admin = '1' "
	strSQL = strSQL & " ORDER BY emp_firstname"
	
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)
    
    strManagerList = strManagerList & "<option value=''>...</option>"
    
    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
			intEmployeeID			= rs("employee_id")
        	strEmployeeFirstname 	= trim(rs("emp_firstname"))
			strEmployeeLastname 	= trim(rs("emp_lastname"))
			
            if session("employee_manager_id") = rs("employee_id") then
                strManagerList = strManagerList & "<option selected value=" & intEmployeeID & ">" & strEmployeeFirstname & " " & strEmployeeLastname & "</option>"
            else
                strManagerList = strManagerList & "<option value='" & intEmployeeID & "'>" & strEmployeeFirstname & " " & strEmployeeLastname & "</option>"
            end if
            
        rs.Movenext
        loop
    end if
    
    call CloseDataBase()
end function

'-----------------------------------------------
' UPDATE EMPLOYEE
'-----------------------------------------------
function updateEmployee(intEmployeeID,strUsername,strFirstname,strLastname,strEmail,strDepartment,intAdmin,intManagerID,intStatus)
	dim strSQL
	
	Call OpenDataBase()
		
	strSQL = "UPDATE yma_employee SET "
	strSQL = strSQL & "emp_username = '" & strUsername & "',"
	strSQL = strSQL & "emp_firstname = '" & Server.HTMLEncode(strFirstname) & "',"
	strSQL = strSQL & "emp_lastname = '" & Server.HTMLEncode(strLastname) & "',"
	strSQL = strSQL & "emp_email = '" & strEmail & "',"
	strSQL = strSQL & "emp_department = '" & strDepartment & "',"
	strSQL = strSQL & "emp_admin = '" & intAdmin & "',"
	strSQL = strSQL & "emp_manager_id = '" & intManagerID & "',"	
	strSQL = strSQL & "emp_status = '" & intStatus & "',"
	strSQL = strSQL & "emp_date_modified = getdate(),"
	strSQL = strSQL & "emp_modified_by = '" & session("logged_username") & "' "
	strSQL = strSQL & "	WHERE employee_id = " & intEmployeeID
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' GET REQUESTER DETAILS (using username)
'-----------------------------------------------
Function getRequesterDetails(strUsername)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_employee "
	strSQL = strSQL & " WHERE emp_username = '" & strUsername & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("requester_id") 		= Trim(rs("employee_id"))
		session("requester_username") 	= Trim(rs("emp_username"))
		session("requester_firstname") 	= Trim(rs("emp_firstname"))
		session("requester_lastname") 	= Trim(rs("emp_lastname"))
		session("requester_email") 		= Trim(rs("emp_email"))
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET RECIPIENT DETAILS (using Account Code)
'-----------------------------------------------
Function getRecipientDetails(strAccountCode)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_employee "
	strSQL = strSQL & " WHERE emp_initial = '" & strAccountCode & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("recipient_username") 	= Trim(rs("emp_username"))
		session("recipient_email") 	  	= Trim(rs("emp_email"))
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' GET MARKETING MGR DETAILS (using username)
'-----------------------------------------------
Function getMarketingManagerDetails(strUsername)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM yma_employee "
	strSQL = strSQL & " WHERE emp_username = '" & strUsername & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("marketing_manager_id") 		= Trim(rs("employee_id"))
		session("marketing_manager_username") 	= Trim(rs("emp_username"))
		session("marketing_manager_firstname") 	= Trim(rs("emp_firstname"))
		session("marketing_manager_lastname") 	= Trim(rs("emp_lastname"))
		session("marketing_manager_email") 		= Trim(rs("emp_email"))
    end if

    call CloseDataBase()
end function
%>