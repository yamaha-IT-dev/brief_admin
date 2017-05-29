<%
'-----------------------------------------------
' ADD LOCATION
'-----------------------------------------------
function addLocation(stockOrderNo, stockOrderLine, stockLocation, stockCreatedBy)
    dim strSQL

    Call OpenDataBase()

    strSQL = "INSERT INTO tbl_loan_location (stockOrderNo, stockOrderLine, stockLocation, stockCreatedBy) VALUES ("
    strSQL = strSQL & " '" & Server.HTMLEncode(stockOrderNo) & "',"
    strSQL = strSQL & " '" & Server.HTMLEncode(stockOrderLine) & "',"
    strSQL = strSQL & " '" & Server.HTMLEncode(stockLocation) & "',"
    strSQL = strSQL & " '" & Trim(stockCreatedBy) & "')"

    on error resume next
    conn.Execute strSQL

    'response.Write strSQL

    if err <> 0 then
        strMessageText = err.description
    else
        Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
        strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Location has been saved.</div>"
    end if

    Call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOAN LOCATION
' UPDATED: Victor Samson 2016-08-10
'-----------------------------------------------
function addLoanLocation(stockOrderNo, stockOrderLine, stockSetSequence, stockLocation, stockCreatedBy)
    Dim cmdObj, paraObj
    Session("anchorTag") = strOrderNo & "" & strOrderLine

    call OpenDataBase

    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spAddLoanLocation"
    cmdObj.CommandType = AdCmdStoredProc

    Set paraObj = cmdObj.CreateParameter("@stockOrderNo",AdInteger,AdParamInput,4,stockOrderNo)
    cmdObj.Parameters.Append paraObj
    Set paraObj = cmdObj.CreateParameter("@stockOrderLine",AdInteger,AdParamInput,4,stockOrderLine)
    cmdObj.Parameters.Append paraObj
    Set paraObj = cmdObj.CreateParameter("@stockSetSequence",AdVarChar,AdParamInput,2,stockSetSequence)
    cmdObj.Parameters.Append paraObj
    Set paraObj = cmdObj.CreateParameter("@stockLocation",AdVarChar,AdParamInput,120,stockLocation)
    cmdObj.Parameters.Append paraObj
    Set paraObj = cmdObj.CreateParameter("@stockCreatedBy",AdVarChar,AdParamInput,50,stockCreatedBy)
    cmdObj.Parameters.Append paraObj

    On Error Resume Next
        Dim rs
        Dim id
        set rs = cmdObj.Execute
        id = rs(0)
        set rs = nothing
    On error Goto 0

    if CheckForSQLError(conn,"Add",MessageText) = TRUE then
        addLoanLocation = FALSE
        strMessageText = MessageText
        'strMessageText = err.description
    else
        addLoanLocation = TRUE
        Response.Redirect(Request.ServerVariables("HTTP_REFERER") & "#" & Session("anchorTag"))
        strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Location has been saved.</div>"
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)

    call CloseDataBase
end function

'----------------------------------------------------------------------------------------
' UPDATE LOAN LOCATION
'----------------------------------------------------------------------------------------
Function updateLocation(stockID, stockLocation, stockModifiedBy)
    dim strSQL

    Call OpenDataBase()

    strSQL = "UPDATE tbl_loan_location SET "
    strSQL = strSQL & "stockLocation = '" & Server.HTMLEncode(stockLocation) & "',"
    strSQL = strSQL & "stockDateModified = GetDate(),"
    strSQL = strSQL & "stockModifiedBy = '" & Trim(stockModifiedBy) & "' WHERE stockID = " & stockID

    'response.Write strSQL
    on error resume next
    conn.Execute strSQL

    'On error Goto 0

    if err <> 0 then
        strMessageText = err.description
    else
        Response.Redirect(Request.ServerVariables("HTTP_REFERER") & "#" & Session("anchorTag"))
        strMessageText = "<div align=""center"" class=""notification_text""><img src=""images/icon_check.png""> The Location has been updated.</div>"
    end if

    Call CloseDataBase()
end function
%>