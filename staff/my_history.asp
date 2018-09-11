<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j, k
Dim iYear
Dim MyHist(), TempArr(3)
Dim dMeetDate

If Not Session("role") = "staff" Then Response.Redirect "/default.asp?sign_out=y"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'first get fitness events
i = 0
ReDim MyHist(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventName, e.EventDate, sa.Comments, e.EventID FROM StaffAsgmt sa INNER JOIN Events e "
sql = sql & "ON sa.EventID = e.EventID WHERE sa.StaffID = " & Session("staff_id") & " AND sa.EventType = 'Fitness Event' ORDER BY e.EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If Year(rs(1).Value) = CInt(iYear) Then
        MyHist(0, i) = Replace(rs(0).Value, "''", "'")
        MyHist(1, i) = rs(1).Value
        MyHist(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve MyHist(2, i)
    End If

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'now get cc/nordic
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, Comments FROM StaffAsgmt WHERE StaffID = " & Session("staff_id") & " AND EventType <> 'Fitness Event'"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    dMeetDate = GetMeetDate(rs(0).Value)
    If Year(dMeetDate) = CInt(iYear) Then
        MyHist(0, i) = GetMeetName(rs(0).Value)
        MyHist(1, i) = dMeetDate
        MyHist(2, i) = rs(1).Value
        i = i + 1
        ReDim Preserve MyHist(2, i)
    End If

    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

For i = 0 To UBound(MyHist, 2) - 2
    For j = i + 1 To UBound(MyHist, 2) - 1
        If CDate(MyHist(1, i)) > CDate(MyHist(1, j)) THen
            For k = 0 To 6
                TempArr(k) = MyHist(k, i)
                MyHist(k, i) = MyHist(k, j)
                MyHist(k, j) = TempArr(k)
            Next
        End If
    Next
Next

Private Function GetMeetDate(lThisMeet)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
    rs2.Open sql2, conn2, 1, 2
    GetMeetDate = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetMeetName(lThisMeet)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT MeetName FROM Meets WHERE MeetsID = " & lThisMeet
    rs2.Open sql2, conn2, 1, 2
    GetMeetName = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>My GSE History</title>
<!--#include file = "../includes/js.asp" -->

<style type="text/css">
     th, td {
        padding-left: 5px;
    }
    
    th {
        text-align: left;
    }
</style>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
  	<div id="row">
		<!--#include file = "staff_menu.asp" -->
		<div class="col-md-10">
	        <h3 class="h3">My GSE&copy; History for <%=iYear%></h3>

            <div style="background-color: #fff;border:none;text-align: right;margin: 0 10px 0 0;padding: 2px;font-size: 0.85em;">
                <a href="/admin/staff/event_assign.asp?year=<%=iYear%>">View As List</a>
                &nbsp;|&nbsp;
                <a href="calendar.asp?year=<%=iYear%>" style="font-size: 0.85em;">Calendar</a>
                &nbsp;|&nbsp;
                <%For i = 2013 To Year(Date) + 1%>
                    <a href="my_history.asp?year=<%=i%>" style="font-size: 0.85em;"><%=i%></a>
                    <%If Not i = Year(Date) + 1 Then%>
                        &nbsp;|&nbsp;
                    <%End If%>
                <%Next%>
            </div>

            <table>
                <tr>
                    <th>No.</th>
                    <th>Event</th>
                    <th>Date</th>
                    <th>Comments</th>
                </tr>
                <%For i = 0 To UBound(MyHist, 2) - 1%>
                    <%If i mod 2 = 0 Then%>
                        <tr>
                            <td class="alt" valign="top"><%=i + 1%>)</td>
                            <td class="alt" style="white-space: nowrap;" valign="top"><%=MyHist(0, i)%></td>
                            <td class="alt" valign="top"><%=MyHist(1, i)%></td>
                            <td class="alt"><%=MyHist(2, i)%></td>
                        </tr>
                    <%Else%>
                        <tr>
                            <td valign="top"><%=i + 1%>)</td>
                            <td style="white-space: nowrap;" valign="top"><%=MyHist(0, i)%></td>
                            <td valign="top"><%=MyHist(1, i)%></td>
                            <td><%=MyHist(2, i)%></td>
                        </tr>
                    <%End If%>
                <%Next%>
            </table>
        </div>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
