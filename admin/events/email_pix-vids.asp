<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, rs, conn, rs2, sql2
Dim i, j
Dim lEventID, lEventDirID
Dim sMyEmail, sEventName, sPageToSend, sEventDirEmail, sEventRaces
Dim FinishersArr(), SendTo()
Dim cdoMessage, cdoConfig, objEmail, xmlhttp, EmailContents
Dim bSentBCC

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 2400

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim Events
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("send_notif") = "send_notif" Then
    Call EventInfo()

    j = 0
    ReDim FinishersArr(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT ir.ParticipantID, pr.Bib FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID WHERE ir.RaceID IN (" 
    sql = sql & sEventRaces & ") AND pr.RaceID IN (" & sEventRaces & ")"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
        FinishersArr(0, j) = rs(0).Value
        FinishersArr(1, j) = rs(1).Value
        j = j + 1
        ReDim Preserve FinishersArr(2, j)
        rs.MoveNext
    Loop
    rs.Close
  	Set rs = Nothing

	For j = 0 To UBound(FinishersArr, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT Email FROM Participant WHERE ParticipantID = " & FinishersArr(0, j)
	    rs.Open sql, conn, 1, 2
	    Do While Not rs.EOF
            If Not rs(0).Value & "" = "" Then FinishersArr(2, j) = rs(0).Value
            rs.MoveNext
        Loop
        rs.Close
  	    Set rs = Nothing
    Next

    i = 0
    ReDim SendTo(0)
	For j = 0 To UBound(FinishersArr, 2) - 1
        If Not FinishersArr(2, j) & "" = "" Then
            If EmailExists(FinishersArr(2, j)) = "n" Then
                SendTo(i) = FinishersArr(2, j)
                i = i + 1
                ReDim Preserve SendTo(i)
            End If
        End If
    Next

    bSentBCC = False

	For j = 0 To UBound(SendTo) - 1
		'get email address
		sMyEmail = vbNullString
				
		If DontSend(SendTo(j)) = False Then sMyEmail = SendTo(j)

		If Not sMyEmail = vbNullString Then
			If ValidEmail(sMyEmail) = True Then
                sPageToSend = "http://www.gopherstateevents.com/misc/pix-vids_notif.asp?event_id=" & lEventID 

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
    				.To = sMyEmail
'                    .To = "bob.schneider@gopherstateevents.com"
				    .From = "bob.schneider@gopherstateevents.com"
                    If bSentBCC = False Then 
                        .BCC = "bob.schneider@gopherstateevents.com;" & sEventDirEmail
                        bSentBCC = True
                    End If
   				    .Subject = "Pictures Are Ready for " & sEventName
				    .HTMLBody = EmailContents
				    .Send
			    End With
			    Set cdoMessage = Nothing
                Set cdoConfig = Nothing
		    End If
        End If
	Next
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

If CLng(lEventID) > 0 Then Call EventInfo()

Private Sub EventInfo()
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT e.EventName, ed.EventDirID, ed.Email FROM Events e INNER JOIN EventDir ed ON e.EventDirID = ed.EventDirID WHERE e.EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    sEventName = Replace(rs(0).Value, "''", "'")
    lEventDirID = rs(1).Value
    sEventDirEmail = rs(2).Value
    rs.Close
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sEventRaces = sEventRaces & rs(0).Value & ", "
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)
End Sub

%>
<!--#include file = "../../includes/valid_email.asp" -->
<%

Function DontSend(sThisEmail) 
	Dim rs2, sql2

    DontSend = False

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Email FROM DontSend WHERE Email = '" & sThisEmail & "' AND (DontSend = 'all' OR DontSend = 'results')"
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then DontSend = True
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function EmailExists(sThisEmail)
    Dim x

    EmailExists = "n"

    For x = 0 To UBound(SendTo) - 1
        If CStr(SendTo(x)) = CStr(sThisEmail) Then
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
<title>GSE&copy; Pix-Vids Ready Notification</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">GSE Pictures Ready Notification for <%=sEventName%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="email_pix-vids.asp?event_id=<%=lEventID%>">
			<label for="events">Events:</label>
			<select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Events, 2)%>
					<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%Else%>
						<option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			<input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
			</form>
			<br>

			<%If Not Clng(lEventID) = 0 Then%>
                <!--#include file = "../../includes/event_nav.asp" -->
                <!--#include file = "email_nav.asp" -->

			    <h4 style="margin:0 0 10px 0;">Send Notification</h4>
				
			    <div style="background-color:#ececd8;margin-bottom: 10px;">
				    <form class="form-inline" name="send_all" Method="Post" action="email_pix-vids.asp?event_id=<%=lEventID%>">
				    <input type="hidden" name="send_notif" id="send_notif" value="send_notif">
				    <input class="form-control" type="submit" name="submit1" id="submit1" value="Send Email">
				    </form>
			    </div>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
