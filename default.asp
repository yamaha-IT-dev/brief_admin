<%
'setup for Australian Date/Time
session.lcid = 2057
session.timeout = 420

Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include file="../include/connection_it.asp " -->
<!--#include file="class/clsEmployee.asp " -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>GD Briefs</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script>
function searchBrief() {
    var strSearch       = document.forms[0].txtSearch.value;
    var strDepartment   = document.forms[0].cboDepartment.value;
    var strUser         = document.forms[0].cboUser.value;
    var strStatus       = document.forms[0].cboStatus.value;
    var strSort         = document.forms[0].cboSort.value;

    document.location.href = 'default.asp?type=search&txtSearch=' + strSearch + '&cboDepartment=' + strDepartment + '&cboUser=' + strUser + '&cboStatus=' + strStatus + '&cboSort=' + strSort;
}

function resetSearch() {
    document.location.href = 'default.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
    select case Trim(Request("type"))
        case "reset"
            session("brief_search")         = ""
            session("brief_department")     = ""
            session("brief_user")           = ""
            session("brief_status")         = ""
            session("brief_sort")           = ""
            session("brief_initial_page")   = 1
        case "search"
            session("brief_search")         = Trim(Request("txtSearch"))
            session("brief_department")     = Trim(Request("cboDepartment"))
            session("brief_user")           = Trim(Request("cboUser"))
            session("brief_status")         = Trim(Request("cboStatus"))
            session("brief_sort")           = Trim(Request("cboSort"))
            session("brief_initial_page")   = 1
    end select
end sub

sub displayBrief
    dim iRecordCount
    iRecordCount = 0
    dim strSortBy
    dim strSortItem
    dim strSearch
    dim strSQL
    dim strType
    dim strSort
    dim strPageResultNumber
    dim strRecordPerPage
    dim intRecordCount
    dim strModifiedDate

    dim strTodayDate
    strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

    set rs = Server.CreateObject("ADODB.recordset")

    rs.CursorLocation = 3   'adUseClient
    rs.CursorType = 3       'adOpenStatic
    rs.PageSize = 100

    if session("brief_sort") = "" then
        session("brief_sort") = "project_deadline"
    end if

    strSQL = "SELECT * FROM yma_project "
    strSQL = strSQL & " WHERE (project_created_by LIKE '%" & session("brief_search") & "%' "
    strSQL = strSQL & "   OR project_output_details LIKE '%" & session("brief_search") & "%' "
    strSQL = strSQL & "   OR project_job_no LIKE '%" & session("brief_search") & "%' "
    strSQL = strSQL & "   OR project_title LIKE '%" & session("brief_search") & "%') "
    strSQL = strSQL & " AND project_department LIKE '%" & session("brief_department") & "%' "
    strSQL = strSQL & " AND project_created_by LIKE '%" & session("brief_user") & "%' AND project_status <> '99' "

    if session("brief_status") = "" then
        strSQL = strSQL & "	AND project_status <> '0' "
    else
        strSQL = strSQL & "	AND project_status LIKE '%" & session("brief_status") & "%' "
    end if

    ' AV briefs are hidden if user is Jamie - Requested by Jamie Bosanko 2016-05-03
    If (session("brief_department") <> "AV") And Session("logged_username") = "jamieb" Then
        strSQL = strSQL & " AND project_department <> 'AV' "
    End If

    strSQL = strSQL & "	ORDER BY " & session("brief_sort")

    rs.Open strSQL, conn
    'Response.Write strSQL

    intPageCount = rs.PageCount
    intRecordCount = rs.recordcount

    Select Case Request("Action")
        case "<<"
            intpage = 1
            session("brief_initial_page") = intpage
        case "<"
            intpage = Request("intpage") - 1
            session("brief_initial_page") = intpage

            if session("brief_initial_page") < 1 then session("brief_initial_page") = 1
        case ">"l
            intpage = Request("intpage") + 1
            session("brief_initial_page") = intpage

            if session("brief_initial_page") > intPageCount then session("brief_initial_page") = IntPageCount
        Case ">>"
            intpage = intPageCount
            session("brief_initial_page") = intpage
    end select

    strDisplayList = ""

    if not DB_RecSetIsEmpty(rs) Then

        rs.AbsolutePage = session("brief_initial_page")

        For intRecord = 1 To rs.PageSize
            if (DateDiff("d",rs("project_date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("project_date_created"), strTodayDate) = 0) then
                if iRecordCount Mod 2 = 0 then
                    strDisplayList = strDisplayList & "<tr class=""updated_today"">"
                else
                    strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
                end if
            else
                if iRecordCount Mod 2 = 0 then
                    strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
                else
                    strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
                end if
            end if

            '1
            if rs("project_status") = 1 then
                strDisplayList = strDisplayList & "<td nowrap><a href=""update_draft_brief.asp?id=" & rs("project_id") & """>Edit Plan</a></td>"
            else
                strDisplayList = strDisplayList & "<td nowrap><a href=""update_brief.asp?id=" & rs("project_id") & """>View</a></td>"
            end if

            '2
            strDisplayList = strDisplayList & "<td>" & rs("project_id") & "</td>"

            '3
            strDisplayList = strDisplayList & "<td nowrap>" & rs("project_created_by") & ", " & FormatDateTime(rs("project_date_created"),1) & "</td>"

            '4
            strDisplayList = strDisplayList & "<td nowrap>" & rs("project_department") & "</td>"

            '5
            strDisplayList = strDisplayList & "<td><span title=""" & rs("project_output_details") & """>"
            strDisplayList = strDisplayList & "" & rs("project_title") & ""
            if DateDiff("d",rs("project_date_created"), strTodayDate) = 0 then
                strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=0>"
            end if

            if (DateDiff("d",rs("project_deadline"), strTodayDate) > 0) and rs("project_status") <> 0 and rs("project_status") <> 7 and rs("project_status") <> 8 then
                strDisplayList = strDisplayList & " <span style=""color:red"">(Overdue)</span>"
            end if
            strDisplayList = strDisplayList & "</span></td>"

            '6
            strDisplayList = strDisplayList & "<td>"
            Select Case	rs("project_priority")
                case 1
                    strDisplayList = strDisplayList & "<font class=""low_font"">Low</font>"
                case 2
                    strDisplayList = strDisplayList & "<font class=""medium_font"">Medium</font>"
                case 3
                    strDisplayList = strDisplayList & "<font class=""high_font"">High</font>"
            end select
            strDisplayList = strDisplayList & "</td>"

            '7
            strDisplayList = strDisplayList & "<td>"
            if rs("project_output_printed") = 1 then
                strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"
            else
                strDisplayList = strDisplayList & "<img src=""images/cross.gif"">"
            end if
            strDisplayList = strDisplayList & "</td>"

            '8
            strDisplayList = strDisplayList & "<td>"
            if rs("project_output_web") = 1 then
                strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"
            else
                strDisplayList = strDisplayList & "<img src=""images/cross.gif"">"
            end if
            strDisplayList = strDisplayList & "</td>"

            '9
            strDisplayList = strDisplayList & "<td>" & rs("project_quote") & "</td>"

            '10
            if IsNull(rs("project_first_deadline")) or rs("project_first_deadline") = "01/01/1900" or rs("project_first_deadline") = "1/1/1900" then
                strDisplayList = strDisplayList & "<td class=""orange_text"">NA</td>"
            else
                strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("project_first_deadline"),1) & "</td>"
            end if

            '11
            if IsNull(rs("project_second_deadline")) or rs("project_second_deadline") = "01/01/1900" or rs("project_second_deadline") = "1/1/1900" then
                strDisplayList = strDisplayList & "<td class=""orange_text"">NA</td>"
            else
                strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("project_second_deadline"),1) & "</td>"
            end if

            '12
            strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("project_deadline"),1) & "</td>"

            '13
            strDisplayList = strDisplayList & "<td>"
            select case rs("product_manager_approval")
                case 0
                    strDisplayList = strDisplayList & "..."
                case 1
                    strDisplayList = strDisplayList & "<font color=""green"">Approved</font>"
                case 2
                    strDisplayList = strDisplayList & "<font color=""red"">Rejected</font>"	
            end select
            strDisplayList = strDisplayList & "</td>"

            '14
            strDisplayList = strDisplayList & "<td>"
            select case rs("marketing_manager_approval")
                case 0
                    strDisplayList = strDisplayList & "..."
                case 1
                    strDisplayList = strDisplayList & "<font color=""green"">Approved</font>"
                case 2
                    strDisplayList = strDisplayList & "<font color=""red"">Rejected</font>"	
            end select
            strDisplayList = strDisplayList & "</td>"

            '15
            if rs("project_progress") = 0 then
                strDisplayList = strDisplayList & "<td align=""left""><table class=""progress_table_red"" border=""0"" cellpadding=""3"" cellspacing=""0"" width=""100%""><tr><td>0%</td></tr></table></td>"
            else
                strDisplayList = strDisplayList & "<td align=""left""><table class=""progress_table_red"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr><td><table class=""progress_table_green"" width=""" & rs("project_progress") & "%""><tr><td>" & rs("project_progress") & "%</td></tr></table></td></tr></table></td>"
            end if

            '16
            strDisplayList = strDisplayList & "<td>"
            Select Case	rs("project_status")
                case 1
                    strDisplayList = strDisplayList & "<font color=""red"">Plan</font>"
                case 2
                    strDisplayList = strDisplayList & "<font color=""red"">Submitted</font>"
                case 3
                    strDisplayList = strDisplayList & "<font color=""red"">Viewed</font>"
                case 4
                    strDisplayList = strDisplayList & "<font color=""orange"">Concept</font>"
                case 5
                    strDisplayList = strDisplayList & "<font color=""orange"">Draft</font>"
                case 6
                    strDisplayList = strDisplayList & "<font color=""orange"">Changes</font>"
                case 7
                    strDisplayList = strDisplayList & "<font color=""green"">Pending approval</font>"
                case 8
                    strDisplayList = strDisplayList & "<font color=""green"">On-hold</font>"
                case 0
                    strDisplayList = strDisplayList & "<font color=""green"">Completed</font>"
                case else
                    strDisplayList = strDisplayList & rs("project_status")
            end select
            strDisplayList = strDisplayList & "</td>"

            '17
            strDisplayList = strDisplayList & "<td nowrap>" & rs("project_job_no") & "</td>"
            if rs("project_status") = 1 then
                strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete this draft titled: - " & rs("project_title") & " ?');"" href='delete_brief.asp?id=" & rs("project_id") & "'><img src=""images/icon_trash.png"" border=""0""></a></td>"
            else
                strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Confirm to delete - " & rs("project_title") & " ?');"" href='deactivate_brief.asp?id=" & rs("project_id") & "'><img src=""images/icon_trash.png"" border=""0""></a></td>"
            end if
            strDisplayList = strDisplayList & "</tr>"

            rs.movenext
            iRecordCount = iRecordCount + 1
            If rs.EOF Then Exit For
        next
    else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""18"" align=""center"">No briefs found.</td></tr>"
    end if

    strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
    strDisplayList = strDisplayList & "<td colspan=""18"" class=""recordspaging"">"
    strDisplayList = strDisplayList & "<form name=""MovePage"" action=""default.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("brief_initial_page") & ">"

    if session("brief_initial_page") = 1 then
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
    else
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
    end if

    if session("brief_initial_page") = intpagecount or intRecordCount = 0 then
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
        strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
    else
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
        strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
    end if

    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strDepartment & ">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strSort & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("brief_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " briefs"
    strDisplayList = strDisplayList & "</h2></form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
    call getEmployeeDetails(session("logged_username"))
    call setSearch

    if trim(session("brief_initial_page"))  = "" then
        session("brief_initial_page") = 1
    end if

    call displayBrief
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table border="0" cellpadding="0" cellspacing="0" class="main_content_table">
    <!-- #include file="include/header.asp" -->
    <tr>
        <td valign="top" class="maincontent">
            <table width="100%" cellpadding="4" cellspacing="0">
                <tr>
                    <td width="10%" valign="top">
                        <div class="alert alert-success"><h2><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_brief.asp">New Brief</a></h2></div>
                        <div align="left"><img src="images/legend-blue.gif" border="1" /> = updated today</div>
                        <p>
                            <a href="export_brief.asp">Export this</a>
                        </p>
                    </td>
                    <td width="90%" valign="top">
                        <div class="alert alert-search">
                            <form name="frmSearch" id="frmSearch" action="default.asp?type=search" method="post" onsubmit="searchBrief()">
                                <h3>Brief Search:</h3>
                                Created by / Title / Job no / Details:
                                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                                <select name="cboDepartment" onchange="searchBrief()">
                                    <option value="">All Departments</option>
                                    <option <% if session("brief_department") = "MPD" then Response.Write " selected" end if%> value="MPD">All MPD</option>
                                    <option <% if session("brief_department") = "MPD - PRO" then Response.Write " selected" end if%> value="MPD - PRO">PRO</option>
                                    <option <% if session("brief_department") = "MPD - TRAD" then Response.Write " selected" end if%> value="MPD - TRAD">TRAD</option>
                                    <option <% if session("brief_department") = "CA" then Response.Write " selected" end if%> value="CA">CA</option>
                                    <option <% if session("brief_department") = "YMEC" then Response.Write " selected" end if%> value="YMEC">YMEC</option>
                                    <option <% if session("brief_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                                </select>
                                <select name="cboUser" onchange="searchBrief()">
                                    <option value="">All Users</option>
                                    <option <% if session("brief_user") = "alex" then Response.Write " selected" end if%> value="alex">Alex</option>
                                    <option <% if session("brief_user") = "atsuko" then Response.Write " selected" end if%> value="atsuko">Atsuko</option> 
                                    <option <% if session("brief_user") = "carolyn" then Response.Write " selected" end if%> value="carolyn">Carolyn</option>
                                    <option <% if session("brief_user") = "cameron" then Response.Write " selected" end if%> value="cameron">Cameron</option>
                                    <option <% if session("brief_user") = "dion" then Response.Write " selected" end if%> value="dion">Dion</option>
                                    <option <% if session("brief_user") = "eric" then Response.Write " selected" end if%> value="eric">Eric</option>
                                    <option <% if session("brief_user") = "euan" then Response.Write " selected" end if%> value="euan">Euan</option>
                                    <option <% if session("brief_user") = "jaclyn" then Response.Write " selected" end if%> value="jaclyn">Jaclyn</option>
                                    <option <% if session("brief_user") = "jamie" then Response.Write " selected" end if%> value="jamie">Jamie</option>
                                    <option <% if session("brief_user") = "julian" then Response.Write " selected" end if%> value="julian">Julian</option>
                                    <option <% if session("brief_user") = "leon" then Response.Write " selected" end if%> value="leon">Leon</option>
                                    <option <% if session("brief_user") = "mick" then Response.Write " selected" end if%> value="mick">Mick</option>
                                    <option <% if session("brief_user") = "mattd" then Response.Write " selected" end if%> value="mattd">Matt Dawkins</option>
                                    <option <% if session("brief_user") = "mattl" then Response.Write " selected" end if%> value="mattl">Matt Livingstone</option>
                                    <option <% if session("brief_user") = "nathan" then Response.Write " selected" end if%> value="nathan">Nathan</option>
                                    <option <% if session("brief_user") = "peta" then Response.Write " selected" end if%> value="peta">Peta</option>
                                </select>
                                <select name="cboStatus" onchange="searchBrief()">
                                    <option <% if session("brief_status") = "" then Response.Write " selected" end if%> value="">All Status (Exclude Completed)</option>
                                    <option <% if session("brief_status") = "1" then Response.Write " selected" end if%> value="1" style="color:red">Status: Plan</option>
                                    <option <% if session("brief_status") = "2" then Response.Write " selected" end if%> value="2" style="color:red">Status: Submitted</option>
                                    <option <% if session("brief_status") = "3" then Response.Write " selected" end if%> value="3" style="color:red">Status: Viewed</option>
                                    <option <% if session("brief_status") = "4" then Response.Write " selected" end if%> value="4" style="color:orange">Status: Concept</option>
                                    <option <% if session("brief_status") = "5" then Response.Write " selected" end if%> value="5" style="color:orange">Status: Draft</option>
                                    <option <% if session("brief_status") = "6" then Response.Write " selected" end if%> value="6" style="color:orange">Status: Changes</option>
                                    <option <% if session("brief_status") = "7" then Response.Write " selected" end if%> value="7" style="color:green">Status: Pending approval</option>
                                    <option <% if session("brief_status") = "8" then Response.Write " selected" end if%> value="8" style="color:green">Status: On-hold</option>
                                    <option <% if session("brief_status") = "0" then Response.Write " selected" end if%> value="0" style="color:green">Status: Completed</option>
                                </select>
                                <select name="cboSort" onchange="searchBrief()">
                                    <option <% if session("brief_sort") = "project_deadline" then Response.Write " selected" end if%> value="project_deadline">Sort by: Deadline</option>
                                    <option <% if session("brief_sort") = "project_date_created" then Response.Write " selected" end if%> value="project_date_created">Sort by: Date created</option>
                                    <option <% if session("brief_sort") = "project_created_by" then Response.Write " selected" end if%> value="project_created_by">Sort by: Requested by</option>
                                </select>
                                <input type="button" name="btnSearch" value="Search" onclick="searchBrief()" />
                                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
                            </form>
                        </div>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td>
            <table cellspacing="0" cellpadding="5" class="database_records" width="100%" border="0">
                <tr class="innerdoctitle">
                    <td></td>
                    <td>ID</td>
                    <td>Created</td>
                    <td>Dept</td>
                    <td>Title</td>
                    <td>Priority</td>
                    <td>Print</td>
                    <td>Online</td>
                    <td>Quote</td>
                    <td>1<sup>st</sup> draft</td>
                    <td>2<sup>nd</sup> draft</td>
                    <td>Print/publish</td>
                    <td>Requester</td>
                    <td>Marketing</td>
                    <td>Progress</td>
                    <td>Status</td>
                    <td>Job #</td>
                    <td></td>
                </tr>
                <%= strDisplayList %>
            </table>
        </td>
    </tr>
</table>
</body>
</html> 