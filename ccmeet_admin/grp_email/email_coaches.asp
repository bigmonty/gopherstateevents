<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k, m
Dim CoachesArr(), MeetsArr(), EmailArr(), ThisMeetArr(), EmailSent()
Dim cdoMessage, cdoConfig
Dim lCoachID
Dim sMsg, sSubject
Dim AttachArr(), sAttachments
Dim iStart, iLength
Dim bEmailSent

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	sSubject = Request.Form.Item("subject")
	
	'get attachment(s)
	sAttachments = Request.Form.Item("attachments")
	iStart = 1
	iLength = 1
	m = 0
	ReDim AttachArr(0)
	For j = 1 to Len(sAttachments)
		If j = Len(sAttachments) Then
			AttachArr(m) = Mid(sAttachments, iStart, iLength + 1) 	
		Else
			If Mid(sAttachments, j,1) = "," Then
				AttachArr(m) = Mid(sAttachments, iStart, iLength - 1) 	
				iStart = j + 1
				iLength = 1
				m = m + 1
				ReDim Preserve AttachArr(m)
			Else
				iLength = iLength + 1
			End If
		End If
	Next
		
	If Request.Form.Item("all") = "on" Then
		i = 0
		ReDim EmailArr(1, 0)
		sql = "SELECT LastName, Email FROM Coaches"
		Set rs = conn.Execute(sql)
		Do While Not rs.EOF
			If Not rs(1).Value & "" = "" Then
				EmailArr(0, i) = Replace(rs(0).Value, "''", "'")
				EmailArr(1, i) = rs(1).Value
				i = i + 1
				ReDim Preserve EmailArr(1, i)
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
	Else	
		'first send email to all who are selected by meet
		'start by getting all meet ids into an array if they were selected
		i = 0
		ReDim ThisMeetArr(0)
		sql = "SELECT MeetsID FROM Meets"
		Set rs = conn.Execute(sql)
		Do While Not rs.EOF
			If Request.Form.Item("meet_" & rs(0).Value) = "on" Then
				ThisMeetArr(i) = rs(0).Value
				i = i + 1
				ReDim Preserve ThisMeetArr(i)
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
		
		'now get the data for the coaches of the teams in these meets
		i = 0
		ReDim EmailArr(1, 0)
		For j = 0 to UBound(ThisMeetArr) - 1
			sql = "SELECT DISTINCT c.Email, c.LastName FROM Coaches c INNER JOIN Teams t ON t.CoachesID = c.CoachesID "
			sql = sql & "INNER JOIN MeetTeams mt ON mt.TeamsID = t.TeamsID WHERE mt.MeetsID = " & ThisMeetArr(j)
			Set rs = conn.Execute(sql)
			Do While Not rs.EOF
				If Not rs(0).Value = vbNullString Then
					EmailArr(0, i) = Replace(rs(1).Value, "''", "'")
					EmailArr(1, i) = rs(0).Value
					i = i + 1
					ReDim Preserve EmailArr(1, i)
				End If
				rs.MoveNext
			Loop
			Set rs = Nothing
		Next
		
		'then send email to all who are selected by name
		sql = "SELECT CoachesID, LastName, Email FROM Coaches"
		Set rs = conn.Execute(sql)
		Do While Not rs.EOF
			If Not rs(2).Value = vbNullString Then
				If Request.Form.Item("coach_" & rs(0).value) = "on" Then
					EmailArr(0, i) = Replace(rs(1).Value, "''", "'")
					EmailArr(1, i) = rs(2).Value
					i = i + 1
					ReDim Preserve EmailArr(1, i)
				End If
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
	End If

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	k = 0
	ReDim EmailSent(0)
	For i = 0 to UBound(EmailArr, 2) - 1
		bEmailSent = False
		For j = 0 To UBound(EmailSent) - 1
			If CStr(EmailArr(1, i)) = CStr(EmailSent(j)) Then
				bEmailSent = True
				Exit For
			End If
		Next
		
		If bEmailSent = False Then
			EmailSent(k) = EmailArr(1, i)
			k = k + 1
			ReDim Preserve EmailSent(k)
			
			sSubject = Request.Form.Item("subject")
			
			sMsg & "Dear Coach " & EmailArr(0, i) & ": " & vbCrLf & vbCrLf
		
			sMsg = sMsg & Request.Form.Item("msg") & vbCrLf & vbCrLf

			sMsg = sMsg & "Sincerely; " & vbCrLf & vbCrLf
			sMsg = sMSg & "Bob Schneider " & vbCrLf
			sMsg = sMSg & "Gopher State Events " & " " & vbCrLf 
			sMsg = sMsg & "www.gopherstateevents.com " & vbCrLf
			sMsg = sMsg & "612-720-8427 " & vbCrLf
	
 			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = EmailArr(1, i)
'				.To = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
				.Subject = sSubject
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
		End If
	Next
	
	Set cdoConfig = Nothing
End If

i = 0
ReDim CoachesArr(1, 0)
sql = "SELECT CoachesID,  FirstName, LastName FROM Coaches ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	CoachesArr(0, i) = rs(0).Value
	CoachesArr(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve CoachesArr(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim MeetsArr(1, 0)
sql = "SELECT MeetsID,  MeetName, MeetDate FROM Meets ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetsArr(0, i) = rs(0).Value
	MeetsArr(1, i) = Replace(rs(1).Value, "''", "'") & " (" & Year(rs(2).Value) & ")"
	i = i + 1
	ReDim Preserve MeetsArr(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim AttachArr(0)
sql = "SELECT Attachment FROM Attach"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	AttachArr(i) = rs(0).Value
	i = i + 1
	ReDim Preserve AttachArr(i)
	rs.MoveNext
Loop
Set rs = Nothing

Function GetTeams(lCoachID)
	sql = "SELECT TeamName FROM Teams WHERE CoachesID = " & lCoachID
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		If GetTeams = vbNullString Then
			GetTeams = rs(0).Value & "<br>"
		Else
			GetTeams =GetTeams & rs(0).Value & "<br>"
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing
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
			<form name="email_coaches" method="post" action="email_coaches.asp">
			<h4 class="h4">Email Cross-Country Coaches</h4>
			<table style="font-size:1.0em;">
				<tr>
					<td valign="top">
						<table>
							<tr>
								<td>
									<span style="font-weight:bold"><input type="checkbox" name="all" id="all">All</span><br>
									<hr>
									<span style="font-weight:bold">By Meet Registered For:</span><br>
									<%For i = 0 to UBound(MeetsArr, 2) - 1%>
										 <input type="checkbox" name="meet_<%=MeetsArr(0, i)%>">
										 <a href="javascript:pop('../../../events/cross_ctry/ccmeet_info.asp?meet_id=<%=MeetsARr(0, i)%>',650,400)"><%=MeetsArr(1, i)%></a><br>
									<%Next%>
									<hr>
									<span style="font-weight:bold">By Name:</span><br>
									<%For i = 0 to UBound(CoachesArr, 2) - 1%>
										<span style="font-weight:bold"><%=i + 1%>)</span>
										 <input type="checkbox" name="coach_<%=CoachesArr(0, i)%>">
										<span style="font-weight:bold"><%=CoachesArr(1, i)%></span><br>
										 <%=GetTeams(CoachesArr(0, i))%><br>
									<%Next%>
								</td>
							</tr>
						</table>
					</td>
					<td valign="top">
						<table>
							<tr>
								<th>
									Subject:
								</th>
								<td>
									<input name="subject" id="subject" style="width:250px">
								</td>
							</tr>
							<tr>
								<th valign="top">
									Attachments:
								</th>
								<td>
									<select name="attachments" id="attachments" size="4" multiple>
										<%For i = 0 to UBound(AttachArr) - 1%>
											<option value="<%="c:\inetpub\h51web\gopherstateevents\ccmeet_admin\grp_email\attachments\" & AttachArr(i)%>"><%=AttachArr(i)%></option>
										<%Next%>
									</select>
								</td>
							</tr>
							<tr>
								<th valign="top">
									Message:
								</th>
								<td>
									<textarea name="msg" id="msg" cols="70" rows="15" style="font-size:1.2em;"></textarea>
								</td>
							</tr>
							<tr>
								<td style="text-align:center;" colspan="2">
									<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
									<input type="submit" name="submit" id="submit" value="Send Email">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
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
