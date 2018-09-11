<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID
Dim i
Dim sEventName, sEventRaces, sSubject, sMsg, sMessage, sEventDirEmail
Dim cdoMessage, cdoConfig
Dim EmailArr()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

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
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
    Call EventInfo()

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ed.Email FROM EventDir ed INNER JOIN Events e ON ed.EventDirID = e.EventDirID WHERE e.EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    sEventDirEmail = rs(0).Value
    rs.Close
    Set rs = Nothing

	sSubject = Request.Form.Item("subject")
	sMsg = Request.Form.Item("message")
	
	i = 0
	ReDim EmailArr(1, 0)
	sql = "SELECT p.Email, p.FirstName, p.LastName FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID "
    sql = sql & "WHERE pr.RaceID IN (" & sEventRaces & ") ORDER BY p.LastName, p.FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
        If Not rs(0).Value & "" = "" Then
            If EmailExists(rs(0).Value) = "n" Then
		        EmailArr(0, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
		        EmailArr(1, i) = rs(0).Value
		        i = i + 1
		        ReDim Preserve EmailArr(1, i)
            End If
        End If
		rs.MoveNext
	Loop
	Set rs = Nothing

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
    	
	For i = 0 to UBound(EmailArr, 2) - 1
        If ValidEmail(EmailArr(1, i)) = True Then
            sMessage = "Dear " & EmailArr(0, i) & ": " & vbCrLf & vbCrLf
            sMessage = sMessage & sMsg & vbCrLf & vbCrLf
            sMessage = sMessage & "Sincerely; " & vbCrLf & vbCrLf
            sMessage = sMessage & "Bob Schneider " & vbCrLf
            sMessage = sMessage & "Gopher State Events, LLC " & vbCrLf
            sMessage = sMessage & "612.720.8427 "

		    Set cdoMessage = CreateObject("CDO.Message")
		    With cdoMessage
			    Set .Configuration = cdoConfig
    			.To = EmailArr(1, i)
'                .To = "bobs@h51software.net"
			    .FROM = "bob.schneider@gopherstateevents.com"
    			If i = 0 Then .BCC = "bob.schneider@gopherstateevents.com;" & sEventDirEmail
			    .Subject = sSubject
			    .TextBody = sMessage
			    .Send
		    End With
		    Set cdoMessage = Nothing

        End If
	Next
	
	Set cdoConfig = Nothing
End If

If CStr(lEventID) = vbNullString Then lEventID = 0
If CLng(lEventID) > 0 Then Call EventInfo()

Private Sub EventInfo()
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventName FROM Events WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    sEventName = replace(rs(0).Value, "''", "'")
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

Private Function EmailExists(sThisEmail)
    Dim x

    EmailExists = "n"

    For x = 0 To UBound(EmailArr, 2) - 1
        If CStr(EmailArr(1, x)) = CStr(sThisEmail) Then
            EmailExists = "y"
            Exit For
        End If
    Next 
End Function
%>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Participant Group Email</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h3 class="h3">Participant Email Utility for <%=sEventName%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="partic_email.asp?event_id=<%=lEventID%>">
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

                <p class="bg-success">Note: This utility will inject the salutation for each participant as well as the GSE signature line.</p>

	            <form role="form" class="form-horizontal" name="send_email" method="post" action="partic_email.asp?event_id=<%=lEventID%>">			
	            <div class="form-group">
		            <label for="subject" class="control-label col-xs-4">Subject:</label>
		            <div class="col-xs-8">
                        <input type="text" class="form-control" name="subject" id="subject" value="<%=sSubject%>">
                    </div>
	            </div>
	            <div class="form-group">
		            <label for="message" class="control-label col-xs-4">Message:</label>
		            <div class="col-xs-8">
                        <textarea class="form-control" name="message" id="message" rows="10"><%=sMsg%></textarea>
                    </div>
	            </div>
	            <div class="form-group">
		            <input type="hidden" class="form-control" name="submit_this" id="submit_this" value="submit_this">
		            <input type="submit" class="form-control" name="submit1" id="submit1" value="Send Email">
	            </div>
	            </form>
            <%End If%>
        </div>
    </div>
<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
