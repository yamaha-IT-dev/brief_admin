<%
'-----------------------------------------------
' GET USER DETAILS (using username)
'-----------------------------------------------
Function getUser(strUsername)
	dim rs
	dim strSQL
	
    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM tbl_office365 "
	strSQL = strSQL & " WHERE usrUsername = '" & strUsername & "'"

	rs.Open strSQL, conn
	'Response.Write strSQL
	
    if not DB_RecSetIsEmpty(rs) Then
		session("usrUserID") 	= rs("usrUserID")
		session("usrPassword") 	= rs("usrPassword")
		session("usrName") 		= rs("usrName")
		session("usrStatus") 	= rs("usrStatus")
	else
		response.redirect("not-found.html")
    end if

    call CloseDataBase()
end function

'-----------------------------------------------
' UPDATE USER STATUS
'-----------------------------------------------
function updateUserStatus(strUserID)
	dim strSQL
	
	Call OpenDataBase()
		
	strSQL = "UPDATE tbl_office365 SET "
	strSQL = strSQL & "usrStatus = '1',"	
	strSQL = strSQL & "usrDateModified = GetDate() "	
	strSQL = strSQL & "	WHERE usrUserID = '" & strUserID & "'"
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "Thank you."
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' ADD USER LOG
'-----------------------------------------------
function addUserLog(strUsername)
	dim strSQL

	Call OpenDataBase()
	
	strSQL = "INSERT INTO tbl_office365_log (logUsername) VALUES ("
	strSQL = strSQL & " '" & strUsername & "')"

	on error resume next
	conn.Execute strSQL	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "IT Department."
	end if
	
	Call CloseDataBase()
end function
%>