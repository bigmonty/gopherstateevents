<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim sTeamSenders
Dim MsgLog()

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 30
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Session("role") = "coach" Then
    sTeamSenders = Session("my_id") & ", "
Else
    sTeamSenders = Session("team_coach_id") & ", " & Session("my_id") & ", "
End If

'get team staff
Set rs = Server.CreateObject("ADODB.Recordset")
If Session("role") = "coach" Then
    sql = "SELECT TeamStaffID FROM TeamStaff WHERE CoachesID = " & Session("my_id")
Else
    sql = "SELECT TeamStaffID FROM TeamStaff WHERE CoachesID = " & Session("team_coach_id") & " AND TeamStaffID <> " & Session("my_id")
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sTeamSenders = sTeamSenders & rs(0).Value & ", "
	rs.MoveNext
Loop
rs.Close
Set rs=Nothing

sTeamSenders = Left(sTeamSenders, Len(sTeamSenders) - 2)

i = 0
ReDim MsgLog(3,0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MsgLogID, SenderID, Subject, WhenSent, SenderRole FROM MsgLog WHERE MsgType = 'Email' AND SenderID IN (" & sTeamSenders 
sql = sql & ") ORDER BY WhenSent DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	MsgLog(0, i) = rs(0).Value
	MsgLog(1, i) = SenderName(rs(1).Value, rs(4).Value)
	MsgLog(2, i) = rs(2).Value
	MsgLog(3, i) = rs(3).Value
	i = i + 1
	ReDim Preserve MsgLog(3, i)
	rs.MoveNext
Loop
rs.Close
Set rs=Nothing

Private Function SenderName(lMyID, sMyRole)
    If sMyRole = "coach" Then
        sql2 = "SELECT FirstName, LastName FROM Coaches WHERE CoachesID = " & lMyID
    Else
        sql2 = "SELECT FirstName, LastName FROM TeamStaff WHERE TeamStaffID = " & lMyID
    End If
    Set rs2 = conn.Execute(sql2)
    If rs2.EOF = rs2.BOF Then
        SenderName = "unknown"
    Else
        SenderName = Replace(rs2(0).Value, "''", "'") & " " & Replace(rs2(1).Value, "''", "'")
    End If
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>Gopher State Events&reg; Group Email Log</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<!--#include file = "communications_nav.asp" -->
            <h4 class="h4">Gopher State Events<sup>&reg;</sup> Group Email Log</h4>	

            <%If UBound(MsgLog, 2) = 0 Then%>
                <p>Your staff has no messages logged.</p>
            <%Else%>
                <table class="table table-striped">
                    <tr>
                        <tr>
                            <th>No.</th>
                            <th>Sender</th>
                            <th>Subject</th>
                            <th>When Sent</th>
                            <th>Details</th>
                        </tr>
                        <%For i = 0 To UBound(MsgLog, 2) - 1%>
                            <tr>
                                <td><%=i + 1%>)</td>
                                <td><%=MsgLog(1, i)%></td>
                                <td><%=MsgLog(2, i)%></td>
                                <td><%=MsgLog(3, i)%></td>
                                <td><a href="javascript:pop('log_details.asp?msg_log_id=<%=MsgLog(0, i)%>',1000,750)">View</a></td>
                            </tr>
                        <%Next%>
                    </tr>
                </table>
            <%End If%>
        </div>
    </div>
</div>
<!--#include file = "../../../../includes/footer.asp" --> 
<%
conn.Close
Set conn=Nothing
%>
</body>
</html>
