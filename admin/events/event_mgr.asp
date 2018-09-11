<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i, j
Dim iNumRsrvd, iNumPndng, iYear
Dim sFtnssErr, sSortBy
Dim sngEventBal, sngBottomLine
Dim Events(), Meets(), EventDir()
Dim fs, fname, sFileName

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

sSortBy = Request.QueryString("sort_by")
If sSortBy = vbNullString Then sSortBy = "EventDate"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
		
iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
							
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim EventDir(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventDirID, FirstName, LastName FROM EventDir ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    EventDir(0, i) = rs(0).Value
	EventDir(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve EventDir(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.form.Item("submit_fitness") = "submit_fitness" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, Status, ShowOnline, EventDirID FROM Events WHERE EventDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear & "'"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		rs(1).Value = Request.Form.Item("status_" & rs(0).Value)
		rs(2).Value = Request.Form.Item("show_online_" & rs(0).Value)
		rs(3).Value = Request.Form.Item("event_dir_id_" & rs(0).Value)
  		rs.Update
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End If

i = 0
iNumPndng = 0
iNumRsrvd = 0
ReDim Events(9, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate, Location, Status, ShowOnline, EventDirID, Edition FROM Events "
sql  = sql & "WHERE EventDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear & "' ORDER BY " & sSortBy
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	For j = 0 To 7
		If Not rs(j).Value & "" = "" Then Events(j, i) = Replace(rs(j).Value, "''", "'")
	Next

    If ChkCntrct(rs(0).Value, rs(2).Value) = True Then 
        Events(8, i) = "View"
        Events(9, i) = "/contracts/" & Year(rs(2).Value) & "/" & rs(0).Value & ".pdf"
    End If

	i = i + 1
	ReDim Preserve Events(9, i)

    If rs(4).Value = "reserved" Then
        iNumRsrvd = CInt(iNumRsrvd) + 1
    Else
        iNumPndng = CInt(iNumPndng) + 1
    End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sngBottomLine = 0

Private Function Balance(lThisEvent)
    Dim lEventGrp
    Dim sEventGrp

    Balance = 0

    'get event grp
    sql = "SELECT EventGrp FROM Events WHERE EventID = " & lThisEvent
    Set rs = conn.Execute(sql)
    lEventGrp = rs(0).Value
    Set rs = Nothing

    'get event in grp
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID FROM Events WHERE EventGrp = " & lEventGrp
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sEventGrp = sEVentGrp & rs(0).Value & ","
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    sEventGrp = Left(sEventGrp, Len(sEventGrp) - 1)

    'get income
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT AmtRcvd FROM FinanceIncome WHERE EventID IN (" & sEventGrp & ") AND IncomeType IN ('Race Deposit', 'Invoice Payment')"
    sql = sql & "AND Sport = 'Fitness Event'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Not rs(0).Value & "" = "" Then Balance = CSng(Balance) + CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get invoice amount
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice FROM FinanceEvents WHERE EventID IN (" & sEventGrp & ") AND Sport = 'Fitness Event'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Not rs(0).Value & "" = "" Then Balance = CSng(Balance) - CSng(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Balance = Round(Balance, 2)
    sngBottomLine = CSng(sngBottomLine) + CSng(Balance)
End Function

Private Function ChkCntrct(lThisEvent, dEventDate)
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
    sFileName = "C:\Inetpub\h51web\gopherstateevents\contracts\" & Year(dEventDate) & "\" & lThisEvent & ".pdf"
    ChkCntrct = fs.FileExists(sFileName)
    Set fs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Event Manager</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
        <!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h4 class="h4">Event Management Utility</h4>

            <%If Not sFtnssErr = vbNullString Then%>
                <p><%=sFtnssErr%></p>
            <%End If%>

            <ul class="nav">
                <li class="nav-item">Total Events: <%=UBound(Events, 2)%></li>
                <li class="nav-item">Pending: <%=iNumPndng%></li>
                <li class="nav-item">Reserved: <%=iNumRsrvd%></li>
            </ul>

            <ul class="nav">
                <%For i = 2002 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="event_mgr.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
            </ul>

            <ul class="nav">
               <li class="nav-item">
                   <a class="nav-link" href="javascript:pop('/admin/events/dwnld_events.asp?year=<%=iYear%>',1000,750)" rel="nofollow">Download</a>
                </li>
                <li class="nav-item"><span style="font-weight: bold;">Sort By:</span></li>

                <li class="nav-item">
                    <a class="nav-link" href="/admin/events/event_mgr.asp?sort_by=EventDate&amp;year=<%=iYear%>" rel="nofollow">Date</a>
                </li>
                <li class="nav-item">
                    <a class="nav-link" href="/admin/events/event_mgr.asp?sort_by=EventName&amp;year=<%=iYear%>" rel="nofollow">Name</a>
                </li>
                <li class="nav-item">
                    <a class="nav-link" href="/admin/events/event_mgr.asp?sort_by=Status&amp;year=<%=iYear%>" rel="nofollow">Status</a>
                </li>
                <li class="nav-item ">
                    <a class="nav-link" href="/admin/events/event_mgr.asp?sort_by=Edition&amp;year=<%=iYear%>" rel="nofollow">Edition</a>
                </li>
                <li class="nav-item">
                    <a class="nav-link" href="/admin/events/event_mgr.asp?sort_by=EventDirID&amp;year=<%=iYear%>" rel="nofollow">Event Dir</a>
                </li>
            </ul>

			<form class="form" name="event_status" method="Post" action="event_mgr.asp?year=<%=iYear%>">
            <div class="table-responsive">
                <table class="table table-striped">
                    <tr>
                        <td style="text-align:center;" colspan="10">
                            <input type="hidden" name="submit_fitness" id="submit_fitness" value="submit_fitness">
                            <input class="form-control" type="submit" name="submit1" id="submit1" value="Submit Changes">
                        </td>
                    </tr>
                    <tr>
                        <th>No.</th>
                        <th style="white-space:nowrap;">Event Name</th>
                        <th>Date</th>
                        <th>Location</th>
                        <th style="white-space:nowrap;">Event Status</th>
                        <th>Visible</th>
                        <th style="white-space:nowrap;">Event Director</th>
                        <th>Bal</th>
                        <th>Cntrct</th>
                        <th>Edtn</th>
                    </tr>
                    <%For i = 0 To UBound(Events, 2) - 1%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td><a href="edit_event.asp?event_id=<%=Events(0, i)%>"><%=Events(1, i)%></a></td>
                            <td><%=Events(2, i)%></td>
                            <td><%=Events(3, i)%></td>
                            <td>
                                <%If Events(4, i) = "pending" Then%>
                                    <select class="form-control" name="status_<%=Events(0, i)%>" id="status_<%=Events(0, i)%>" style="background-color: yellow">
                                        <%If Events(4, i) = "pending" Then%>
                                            <option value="pending" selected>pending</option>
                                            <option value="reserved">rsvrd</option>
                                        <%Else%>
                                            <option value="pending">pending</option>
                                            <option value="reserved" selected>rsvrd</option>
                                        <%End If%>
                                    </select>
                                <%Else%>
                                    <select class="form-control" name="status_<%=Events(0, i)%>" id="status_<%=Events(0, i)%>">
                                        <%If Events(4, i) = "pending" Then%>
                                            <option value="pending" selected>pending</option>
                                            <option value="reserved">rsvrd</option>
                                        <%Else%>
                                            <option value="pending">pending</option>
                                            <option value="reserved" selected>rsvrd</option>
                                        <%End If%>
                                    </select>
                                <%End If%>
                            </td>
                            <td>
                                <%If Events(5, i) = "n" Then%>
                                    <select class="form-control" name="show_online_<%=Events(0, i)%>" id="show_online_<%=Events(0, i)%>" style="background-color: yellow;">
                                        <%If Events(5, i) = "n" Then%>
                                            <option value="n" selected>n</option>
                                            <option value="y">y</option>
                                        <%Else%>
                                            <option value="n">n</option>
                                            <option value="y" selected>y</option>
                                        <%End If%>
                                    </select>
                                <%Else%>
                                    <select class="form-control" name="show_online_<%=Events(0, i)%>" id="show_online_<%=Events(0, i)%>">
                                        <%If Events(5, i) = "n" Then%>
                                            <option value="n" selected>n</option>
                                            <option value="y">y</option>
                                        <%Else%>
                                            <option value="n">n</option>
                                            <option value="y" selected>y</option>
                                        <%End If%>
                                    </select>
                                <%End If%>
                            </td>
                            <td>
                                <select class="form-control" name="event_dir_id_<%=Events(0, i)%>" id="event_dir_id_<%=Events(0, i)%>">
                                    <%For j = 0 To UBound(EventDir, 2) - 1%>
                                        <%If CLng(Events(6, i)) = CLng(EventDir(0, j)) Then%>
                                            <option value="<%=EventDir(0, j)%>" selected><%=EventDir(1, j)%></option>
                                        <%Else%>
                                            <option value="<%=EventDir(0, j)%>"><%=EventDir(1, j)%></option>
                                        <%End If%>
                                    <%Next%>
                                </select>
                            </td>
                            <%sngEventBal = Balance(Events(0, i))%>
                            <%If CSng(sngEventBal) < 0 Then%>
                                <td style="text-align:right;color: red;">$<%=sngEventBal%></td>
                            <%ElseIf CSng(sngEventBal) = 0 Then%>
                                <td style="text-align:right;color: #ffd800;">$<%=sngEventBal%></td>
                            <%Else%>
                                <td style="text-align:right;">$<%=sngEventBal%></td>
                            <%End If%>
                            <td style="text-align: center;">
                                <%If Events(8, i) ="View" Then%>
                                    <a href="javascript:pop('<%=Events(9, i)%>',800,600)"><%=Events(8, i)%></a>
                                <%Else%>
                                    &nbsp;
                                <%End If%>
                            </td>
                            <td style="text-align:center;"><%=Events(7, i)%></td>
                        </tr>
                    <%Next%>
                    <tr>
                        <th style="text-align: right;" colspan="7">Bottom Line:</th>
                        <th>$<%=Round(sngBottomLine, 2)%></th>
                        <td colspan="2">&nbsp;</td>
                    </tr>
                </table>
            </div>
			</form>
		</div>
	</div>
</div>

<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>