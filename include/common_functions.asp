<%
'-----------------------------------------------
' GET DAMAGE TYPES - list_damages.asp
'-----------------------------------------------
function getDamageTypeList
    dim arrDamageTypeFillText
    dim arrDamageTypeFillID
    dim intCounter

    arrDamageTypeFillText 	= split(arrDamageTypeText, ",")
    arrDamageTypeFillID 	= split(arrDamageTypeID, ",")
   
    if isarray(arrDamageTypeFillID) then
        if ubound(arrDamageTypeFillID) > 0 then        
            for intCounter = 0 to ubound(arrDamageTypeFillID)                
                if trim(session("strDamageType")) = trim(arrDamageTypeFillID(intCounter)) then
                    strDamageTypeList = strDamageTypeList & "<option selected value=" & arrDamageTypeFillID(intCounter) & ">" & arrDamageTypeFillText(intCounter) & "</option>"
                else
                   	strDamageTypeList = strDamageTypeList & "<option value=" & arrDamageTypeFillID(intCounter) & ">" & arrDamageTypeFillText(intCounter) & "</option>"
                end if             
            next
        end if    
    end if

end function


'-----------------------------------------------
' GET DAMAGE TYPES - update_damage.asp
'-----------------------------------------------
function getDamageType
    dim arrDamageTypeFillText
    dim arrDamageTypeFillID
    dim intCounter

    arrDamageTypeFillText 	= split(arrDamageTypeText, ",")
    arrDamageTypeFillID 	= split(arrDamageTypeID, ",")
   
    if isarray(arrDamageTypeFillID) then
        if ubound(arrDamageTypeFillID) > 0 then        
            for intCounter = 0 to ubound(arrDamageTypeFillID)                
                if trim(session("damage_type")) = trim(arrDamageTypeFillID(intCounter)) then
                    strDamageTypeList = strDamageTypeList & "<option selected value=" & arrDamageTypeFillID(intCounter) & ">" & arrDamageTypeFillText(intCounter) & "</option>"
                else
                   	strDamageTypeList = strDamageTypeList & "<option value=" & arrDamageTypeFillID(intCounter) & ">" & arrDamageTypeFillText(intCounter) & "</option>"
                end if             
            next
        end if    
    end if

end function

'-----------------------------------------------
' GET COURSE OF DAMAGE - list_damages.asp
'-----------------------------------------------
function getCourseDamageList
    dim arrCourseDamageFillText
    dim arrCourseDamageFillID
    dim intCounter

    arrCourseDamageFillText	= split(arrCourseDamageText, ",")
    arrCourseDamageFillID 	= split(arrCourseDamageID, ",")
        
    if isarray(arrCourseDamageFillID) then
        if ubound(arrCourseDamageFillID) > 0 then        
            for intCounter = 0 to ubound(arrCourseDamageFillID)                
                if trim(session("strCourseDamage")) = trim(arrCourseDamageFillID(intCounter)) or trim(session("course_damage")) = trim(arrCourseDamageFillID(intCounter)) then
                    strCourseDamageList = strCourseDamageList & "<option selected value=" & arrCourseDamageFillID(intCounter) & ">" & arrCourseDamageFillText(intCounter) & "</option>"
                else
                   	strCourseDamageList = strCourseDamageList & "<option value=" & arrCourseDamageFillID(intCounter) & ">" & arrCourseDamageFillText(intCounter) & "</option>"
                end if             
            next
        end if    
    end if

end function

'-----------------------------------------------
' GET COURSE OF DAMAGE - update_damage.asp
'-----------------------------------------------
function getCourseDamage
    dim arrCourseDamageFillText
    dim arrCourseDamageFillID
    dim intCounter

    arrCourseDamageFillText	= split(arrCourseDamageText, ",")
    arrCourseDamageFillID 	= split(arrCourseDamageID, ",")
        
    if isarray(arrCourseDamageFillID) then
        if ubound(arrCourseDamageFillID) > 0 then        
            for intCounter = 0 to ubound(arrCourseDamageFillID)                
                if trim(session("course_damage")) = trim(arrCourseDamageFillID(intCounter)) then
                    strCourseDamageList = strCourseDamageList & "<option selected value=" & arrCourseDamageFillID(intCounter) & ">" & arrCourseDamageFillText(intCounter) & "</option>"
                else
                   	strCourseDamageList = strCourseDamageList & "<option value=" & arrCourseDamageFillID(intCounter) & ">" & arrCourseDamageFillText(intCounter) & "</option>"
                end if             
            next
        end if    
    end if

end function
%>