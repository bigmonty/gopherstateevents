<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, rs, conn
Dim i, j
Dim lThisMeet
Dim sMeetName, sPageToSend
Dim SendTo(), MeetTeams()
Dim cdoMessage, cdoConfig, objEmail, xmlhttp, EmailContents
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetDate, MeetName FROM Meets WHERE MeetsID = " & lThisMeet 
Set rs = conn.Execute(sql)
dMeetDate = rs(0).Value
sMeetName = Replace(rs(1).Value, "''", "")
Set rs = Nothing

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

If Request.Form.Item("send_notif") = "send_notif" Then
    i = 0
    ReDim SendTo(0)
	For j = 0 To UBound(MeetTeams, 2) - 1
        'get coach email
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT c.Email FROM Coaches c INNER JOIN Teams t ON c.CoachesID = t.CoachesID WHERE t.TeamsID = " & MeetTeams(0, j)
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            If Not rs(0).Value & "" = "" Then
                SendTo(i) = rs(0).Value
                i = i + 1
                ReDim Preserve SendTo(i)
            End If
        End If
        rs.Close
        Set rs = Nothing

        'get followers 
    Next

	For j = 0 To UBound(SendTo) - 1
		If ValidEmail(SendTo(j)) = True Then
            sPageToSend = "http://www.gopherstateevents.com/ccmeet_admin/manage_meet/pix-vids_notif.asp?meet_id=" & lThisMeet 

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

            Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
	            xmlhttp.open "GET", sPageToSend, false
	            xmlhttp.send ""
	            EmailContents = xmlhttp.responseText
            Set xmlhttp = nothing

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
    			.To = SendTo(j)
'                .To = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
                If j = 0 Then .BCC = "bob.schneider@gopherstateevents.com;"
   				.Subject = "Pictures Are Ready for " & sMeetName
				.HTMLBody = EmailContents
				.Send
			End With
			Set cdoMessage = Nothing
            Set cdoConfig = Nothing
		End If
	Next
End If

%>
<!--#include file = "../../includes/valid_email.asp" -->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; CC/Nordic Media Notification</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
    <!--#include file = "../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
	        <!--#include file = "manage_meet_nav.asp" -->
			
			<h4 class="h4">GSE CC-Nordic Media Notification</h4>
				
			<form class="form-inline" name="send_all" Method="Post" action="media_notif.asp?meet_id=<%=lThisMeet%>">
			<input type="hidden" name="send_notif" id="send_notif" value="send_notif">
			<input type="submit" class="form-control" name="submit1" id="submit1" value="Send Notification">
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
