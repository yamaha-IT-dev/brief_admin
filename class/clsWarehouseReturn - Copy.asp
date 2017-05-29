<%
'-----------------------------------------------
' GET REASON CODES
'-----------------------------------------------
function getReasonCodeList
    dim arrReturnReasonCodeFillText
    dim arrReturnReasonCodeFillID
    dim intCounter

    arrReturnReasonCodeFillText = split(arrReturnReasonCodeText, ",")
    arrReturnReasonCodeFillID 	= split(arrReturnReasonCodeID, ",")
   
    if isarray(arrReturnReasonCodeFillID) then
        if ubound(arrReturnReasonCodeFillID) > 0 then
            for intCounter = 0 to ubound(arrReturnReasonCodeFillID)
                if trim(session("damage_search_type")) = trim(arrReturnReasonCodeFillID(intCounter)) then
                    strReasonCodeList = strReasonCodeList & "<option selected value=" & arrReturnReasonCodeFillID(intCounter) & ">" & arrReturnReasonCodeFillText(intCounter) & "</option>"
                else
                   	strReasonCodeList = strReasonCodeList & "<option value=" & arrReturnReasonCodeFillID(intCounter) & ">" & arrReturnReasonCodeFillText(intCounter) & "</option>"
                end if
            next
        end if
    end if
end function

'-----------------------------------------------
' GET REASON CODES - update_warehouse-return.asp
'-----------------------------------------------
function getReasonCode
    dim arrReturnReasonCodeFillText
    dim arrReturnReasonCodeFillID
    dim intCounter

    arrReturnReasonCodeFillText = split(arrReturnReasonCodeText, ",")
    arrReturnReasonCodeFillID 	= split(arrReturnReasonCodeID, ",")
	
   	strReasonCodeList = strReasonCodeList & "<option value=''>...</option>"
	
    if isarray(arrReturnReasonCodeFillID) then
        if ubound(arrReturnReasonCodeFillID) > 0 then
			
            for intCounter = 0 to ubound(arrReturnReasonCodeFillID)                
                if trim(session("reason_code")) = trim(arrReturnReasonCodeFillID(intCounter)) then
                    strReasonCodeList = strReasonCodeList & "<option selected value=" & arrReturnReasonCodeFillID(intCounter) & ">" & arrReturnReasonCodeFillText(intCounter) & "</option>"
                else
                   	strReasonCodeList = strReasonCodeList & "<option value=" & arrReturnReasonCodeFillID(intCounter) & ">" & arrReturnReasonCodeFillText(intCounter) & "</option>"
                end if        
            next
        end if
    end if
end function

'----------------------------------------------------------------------------------------
' UPDATE RETURN
'----------------------------------------------------------------------------------------
Function updateReturn(intID, intReasonID, strGRA, strUsername)
	dim strSQL

	Call OpenDataBase()

	strSQL = "UPDATE yma_quarantines SET "
	strSQL = strSQL & "reason_code = '" & intReasonID & "',"
	strSQL = strSQL & "gra = '" & Server.HTMLEncode(strGRA) & "',"
	strSQL = strSQL & "date_modified = getdate(),"
	strSQL = strSQL & "modified_by = '" & strUsername & "' WHERE quarantine_id = " & intID

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