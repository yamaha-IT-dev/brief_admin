<%
'setup for Australian Date/Time
session.lcid = 2057
session.timeout = 420

const projectModuleID = 10

session("logged_username") = Mid(Lcase(Request.ServerVariables("REMOTE_USER")),12,20)
'call getEmployeeDetails(session("logged_username"))

function displayDateFormatted(strDateInput)	
	if IsNull(strDateInput) or strDateInput = "01/01/1900" or strDateInput = "1/1/1900"  then 
		response.write "N/A"
	else
		response.write "" & WeekDayName(WeekDay(strDateInput)) & ", " & FormatDateTime(strDateInput,1) & " at " & FormatDateTime(strDateInput,3)	
	end if
end function
%>
<!--<div align="left"><br>
<a href="http://intranet:96/list_auction.asp" title="Home"><img src="images/ybay.gif" border="0" alt="Home" /></a></div>-->
