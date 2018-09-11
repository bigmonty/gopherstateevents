<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k, m
Dim MeetDirArr()
Dim cdoMessage, cdoConfig
Dim lMeetDirID
Dim EmailArr(), sMsg, sSubject
Dim AttachArr(), sAttachments
Dim iStart, iLength

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
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
	
	i = 0
	ReDim EmailArr(1, 0)
	sql = "SELECT MeetDirID, LastName, Email FROM MeetDir ORDER BY LastName, FirstName"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		If Request.Form.Item("meet_dir_" & rs(0).Value) = "on" Then
			EmailArr(0, i) = Replace(rs(1).Value, "''", "'")
			EmailArr(1, i) = rs(2).Value
			i = i + 1
			ReDim Preserve EmailArr(1, i)
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing
	
%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	For i = 0 to UBound(EmailArr, 2) - 1
		sSubject = Request.Form.Item("subject")
		
		sMsg = sMsg & "Dear Meet Director " & EmailArr(0, i) & ": " & vbCrLf & vbCrLf
	
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
			.From = "bob.schneider@gopherstateevents.com"
			.BCC = "bob.schneider@gopherstateevents.com"
			For j = 0 to UBound(AttachArr)
				.AddAttachment AttachArr(j)
			Next
			
			.Subject = sSubject
			.TextBody = sMsg
			.Send
		End With
		Set cdoMessage = Nothing
	Next
	
	Set cdoConfig = Nothing
End If

i = 0
ReDim MeetDirArr(1, 0)
sql = "SELECT MeetDirID,  FirstName, LastName FROM MeetDir ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetDirArr(0, i) = rs(0).Value
	MeetDirArr(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	i = i + 1
	ReDim Preserve MeetDirArr(1, i)
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

Function GetMeets(lMeetDirID)
	sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetDirID = " & lMeetDirID & " ORDER BY MeetDate DESC"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		If GetMeets = vbNullString Then
			GetMeets = rs(0).Value & " (" & Year(CDate(rs(1).Value)) & ")<br>"
		Else
			GetMeets =GetMeets & rs(0).Value & " (" & Year(CDate(rs(1).Value)) & ")<br>"
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

<title>GSE Email Cross-Country/Nordic Ski Meet Directors</title>

<!--#include file = "../../includes/js.asp" -->
</head>
<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		
		<div class="col-md-10">
			<form name="email_meet_directors" method="post" action="email_meet_dirs.asp">
			<h4 class="h4">Email Cross-Country Coaches</h4>			
            <table style="font-size:1.0em;">
				<tr>
					<td valign="top">
						<table>
							<tr>
								<td>
									<%For i = 0 to UBound(MeetDirArr, 2) - 1%>
										 <input type="checkbox" name="meet_dir_<%=MeetDirArr(0, i)%>">
										 <span style="font-weight:bold"><%=MeetDirArr(1, i)%></span><br>
										 <%=GetMeets(MeetDirArr(0, i))%><br>
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
											<option value="<%="c:\inetpub\h51web\gopherstateevents\admin\ccmeet_admin\grp_email\attachments\" & AttachArr(i)%>"><%=AttachArr(i)%></option>
										<%Next%>
									</select>
								</td>
							</tr>
							<tr>
								<th valign="top">
									Message:
								</th>
								<td>
									<textarea name="msg" id="msg" rows="15" cols="75" style="font-size:1.1em;"></textarea>
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
