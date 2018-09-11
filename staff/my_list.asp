<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim sFirstName, sLastName
Dim iThisYear, iNumTech, iNumSupp

If Not Session("role") = "staff" Then Response.Redirect "/default.asp?sign_out=y"

iThisYear = Request.QueryString("this_year")
If CStr(iThisYear) = vbNullString Then iThisYear = Year(Date)

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(5, 0)
sql = "SELECT EventID, EventName, EventDate, Location, TimingMethod FROM Events WHERE EventDate >= '" & Date & "' ORDER BY EventDate, EventName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    If Year(rs(2).Value) = CInt(iThisYear) Then
	    Events(0, i) = rs(0).Value
	    Events(1, i) = Replace(rs(1).Value, "''","'")
	    Events(2, i) = rs(2).Value
        Events(3, i) = rs(3).Value
        Events(4, i) = rs(4).Value
        Events(5, i) = "na"
	    i = i + 1
	    ReDim Preserve Events(5, i)
    End If
	rs.MoveNext
Loop
Set rs = Nothing

iNumTech = 0
iNumSupp = 0
For i = 0 To UBound(Events, 2) - 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Tech1, Tech2, Support1, Support2, Support3 FROM EventAsgmt WHERE EventID = " & Events(0, i)
    rs.Open sql, conn, 1, 2
    For j = 0 To 4
        If CLng(rs(j).Value) = CLng(Session("staff_id")) Then
            If j = 0 Or j = 1 Then
                iNumTech = CInt(iNumTech) + 1
            Else
                iNumSupp = CInt(iNumSupp) + 1
            End If
            Events(5, i) = rs(j).Name
            Exit For
        End If
    Next
    rs.Close
    Set rs = Nothing
Next
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Staff History</title>
<meta name="description" content="Gopher State Events staff history page.">
<!--#include file = "../includes/js.asp" -->
<style type="text/css">
    th, td {
        padding-right: 10px;
        text-align: left;
    }
</style>
</head>

<body>
<div style="margin: 10px;padding: 10px;background-color: #fff;">
	<h3>My GSE Assignments</h3>

    <div style="text-align: right;margin: 5px 0 0 0;padding: 0;font-size: 0.9em;">
        <ul style="display: inline-block;">
            <%For i = 2013 To Year(Date)%>
                <li style="display: inline-block;"><a href="my_list.asp?this_year=<%=i%>"><%=i%></a>&nbsp;&nbsp;&nbsp;</li>
            <%Next%>
            <li style="display: inline-block;"><a href="javascript:window.print();">Print</a></li>
        </ul>
    </div>

    <table>
        <tr>
            <th>No.</th>
            <th>Event</th>
            <th>Date</th>
            <th>Location</th>
            <th>Timing</th>
        </tr>
        <%j = 1%>
        <%For i = 0 To UBound(Events, 2) - 1%>
            <%If Not Events(5, i) = "na" Then%>
                <tr>
                    <td valign="top"><%=j%>)</td>
                    <td valign="top"><%=Events(1, i)%></td>
                    <td valign="top"><%=Events(2, i)%></td>
                    <td valign="top"><%=Events(3, i)%></td>
                    <td valign="top"><%=Events(4, i)%></td>
                </tr>
                <%j = j + 1%>
            <%End If%>
        <%Next%>
    </table>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>