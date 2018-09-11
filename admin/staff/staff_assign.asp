<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j
Dim iYear
Dim Staff(), MyHist()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
									
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(1, 0)
sql = "SELECT StaffID, LastName, FirstName FROM Staff WHERE  Active = 'y' ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
	i = i + 1
	ReDim Preserve Staff(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub GetMyHist(lThisStaff)
    Dim x, y, z
    Dim TempArr(3)

    x = 0
    ReDim MyHist(3, 0)

    'first get fitness events
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT e.EventName, e.EventDate, sa.Role, sa.Comments, e.EventID FROM StaffAsgmt sa INNER JOIN Events e "
    sql = sql & "ON sa.EventID = e.EventID WHERE sa.StaffID = " & lThisStaff & " AND sa.EventType = 'Fitness Event' ORDER BY e.EventDate"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Year(rs(1).Value) = CInt(iYear) Then
            MyHist(0, x) = Replace(rs(0).Value, "''", "'")
            MyHist(1, x) = rs(1).Value
            MyHist(2, x) = rs(2).Value
            MyHist(3, x) = rs(3).Value
            x = x + 1
            ReDim Preserve MyHist(3, x)
        End If

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'now get cc/nordic
    Dim dMeetDate
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, Role, Comments FROM StaffAsgmt WHERE StaffID = " & lThisStaff & " AND EventType <> 'Fitness Event'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        dMeetDate = GetMeetDate(rs(0).Value)
        If Year(dMeetDate) = CInt(iYear) Then
            MyHist(0, x) = GetMeetName(rs(0).Value)
            MyHist(1, x) = dMeetDate
            MyHist(2, x) = rs(1).Value
            MyHist(3, x) = rs(2).Value
            x = x + 1
            ReDim Preserve MyHist(3, x)
        End If

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For x = 0 To UBound(MyHist, 2) - 2
        For y = x + 1 To UBound(MyHist, 2) - 1
            If CDate(MyHist(1, x)) > CDate(MyHist(1, y)) THen
                For z = 0 To 3
                    TempArr(z) = MyHist(z, x)
                    MyHist(z, x) = MyHist(z, y)
                    MyHist(z, y) = TempArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Function GetMeetDate(lThisMeet)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
    rs2.Open sql2, conn2, 1, 2
    If rs2.RecordCount > 0 Then GetMeetDate = rs2(0).Value
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
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Staff Assignment</title>

<!--#include file = "../../includes/js.asp" -->

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
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">GSE Staff Assignments for <%=iYear%></h4>

            <div style="background-color: #ececd8;text-align: right;margin: 0 10px 0 10px;padding: 2px;font-size: 0.85em;">
                <a href="event_assign.asp?year=<%=iYear%>">Assignments By Event</a>
                &nbsp;|&nbsp;
                <%For i = 2013 To Year(Date) + 1%>
                    <a href="staff_assign.asp?year=<%=i%>"><%=i%></a>
                    <%If Not i = Year(Date) + 1 Then%>
                        &nbsp;|&nbsp;
                    <%End If%>
                <%Next%>
            </div>

            <%For i = 0 To UBound(Staff, 2) - 1%>
                <h4 class="h4"><%=Staff(1, i)%></h4>
                <%Call GetMyHist(Staff(0, i))%>
                <table>
                    <tr>
                        <th colspan="7"></th>
                    </tr>
                    <tr>
                        <th>No.</th>
                        <th>Event</th>
                        <th>Date</th>
                        <th>Role</th>
                        <th>Comments</th>
                    </tr>
                    <%For j = 0 To UBound(MyHist, 2) - 1%>
                        <tr>
                            <td valign="top"><%=j + 1%>)</td>
                            <td style="white-space: nowrap;" valign="top"><%=MyHist(0, j)%></td>
                            <td valign="top"><%=MyHist(1, j)%></td>
                            <td valign="top"><%=MyHist(2, j)%></td>
                            <td><%=MyHist(3, j)%></td>
                        </tr>
                    <%Next%>
                </table>
            <%Next%>
   		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
