<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet
Dim cdoMessage, cdoConfig
Dim sMsg, sMessage, sSubject, sMeetName
Dim SendTo(), Coaches()
Dim dMeetDate
Dim bEmailSent

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

'get meet teams array
i = 0
ReDim MeetTeams(1, 0)
sql = "SELECT mt.TeamsID, t.TeamName, t.Gender FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0,  i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve MeetTeams(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim Coaches(5, 0)
For j = 0 To UBound(MeetTeams, 2) - 1
    sql = "SELECT c.CoachesID,  c.FirstName, c.LastName, c.Email FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID WHERE t.TeamsID = "
    sql = sql & MeetTeams(0, j) & " ORDER BY c.LastName, c.FirstName"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    Coaches(0, i) = rs(0).Value
	    Coaches(1, i) = Replace(rs(1).Value, "''", "'") 
        Coaches(2, i) = Replace(rs(2).Value, "''", "'")
        Coaches(3, i) = rs(3).Value
        Coaches(4, i) = MeetTeams(1, j)
        Coaches(5, i) = MeetTeams(0, j)
	    i = i + 1
	    ReDim Preserve Coaches(5, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing
Next

If Request.Form.Item("submit_this") = "submit_this" Then
	sSubject = Request.Form.Item("subject")
    sMessage = Request.Form.Item("message")

    i = 0
    ReDim SendTo(1, 0)
	For j = 0 To UBound(Coaches, 2) - 1
        If Request.Form.Item("send_all") = "on" Then
            If Not Coaches(3, j) & "" = "" Then
		        If ValidEmail(Coaches(3, j)) = True Then
                    If EmailExists(Coaches(3, j)) = "n" Then
			            SendTo(0, i) = Coaches(2, j)
                        SendTo(1, i) = Coaches(3, j)
			            i = i + 1
			            ReDim Preserve SendTo(1, i)
                    End If
                End If
            End If
        Else
		    If Request.Form.Item("request_" & Coaches(0, j)) = "on" Then
                If Not Coaches(3, j) & "" = "" Then
		            If ValidEmail(Coaches(3, j)) = True Then
                        If EmailExists(Coaches(3, j)) = "n" Then
			                SendTo(0, i) = Coaches(2, j)
                            SendTo(1, i) = Coaches(3, j)
			                i = i + 1
			                ReDim Preserve SendTo(1, i)
                        End If
                    End If
                End If
		    End If
        End If
    Next

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

	For i = 0 to UBound(SendTo, 2) - 1
		sMsg = vbNullString
		sMsg = "Dear Coach " & SendTo(0, i) & ": " & vbCrLf & vbCrLf
	
        sMsg = sMsg & sMessage & vbCrLf & vbCrLf

		sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
		sMsg = sMSg & "Bob Schneider " & vbCrLf
		sMsg = sMSg & "GSE (Gopher State Events) " & " " & vbCrLf 
		sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
		sMsg = sMsg & "612-720-8427 " & vbCrLf

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
			 
		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = SendTo(1, i)
			If i = 0 Then .BCC = "bob.schneider@gopherstateevents.com"
			.From = "bob.schneider@gopherstateevents.com"
			.Subject = sSubject
			.TextBody = sMsg
			.Send
		End With
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing
	Next
	
	Set cdoConfig = Nothing
End If

%>
<!--#include file = "../../includes/valid_email.asp" -->
<%

Private Function EmailExists(sThisEmail)
    Dim x

    EmailExists = "n"

    For x = 0 To UBound(SendTo, 2) - 1
        If CStr(SendTo(1, x)) = CStr(sThisEmail) Then
            EmailExists = "y"
            Exit For
        End If
    Next 
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Email CC Coaches</title>
<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>

			<h4 class="h4">Email Coaches</h4>		
            	
			<form class="form" name="email_coaches" method="post" action="email_coaches.asp?meet_id=<%=lThisMeet%>">
			<div class="col-sm-6">
				<table class="table table-striped">
					<tr>
						<td colspan="4"><input type="checkbox" name="send_all" id="send_all">&nbsp;Send All</td>
					</tr>
					<tr>
                        <th>No.</th>
                        <th>Team</th>
                        <th>Coach</th>
						<th>Send</th>
					</tr>
					<%For i = 0 to UBound(Coaches, 2) - 1%>
						<tr>
							<td><%=i + 1%>)</td>
							<td><%=Coaches(4, i)%></td>
							<td><%=Coaches(2, i)%>, <%=Coaches(1, i)%></td>
							<td><input type="checkbox" name="request_<%=Coaches(0, i)%>" id="request_<%=Coaches(0, i)%>"></td>
						</tr>
					<%Next%>
				</table>
            </div>
            <div class="col-sm-6">
				<table class="table">
					<tr>
						<th>Subject:</th>
						<td><input type="text" class="form-control" name="subject" id="subject"></td>
					</tr>
					<tr>
						<th valign="top">Message:</th>
						<td><textarea class="form-control" name="message" id="message" rows="15" style="font-size:1.2em;"></textarea></td>
					</tr>
					<tr>
						<td style="text-align:center;" colspan="2">
							<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
							<input type="submit" class="form-control" name="submit" id="submit" value="Send Email">
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
