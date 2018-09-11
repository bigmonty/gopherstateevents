<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim EventDir(), AttachArr(), EmailArr()
Dim i, m, j
Dim sSubject, sMsg, sAttachments, iStart, iLength
Dim cdoMessage, cdoConfig
Dim sErrMsg
Dim bRecipients
Dim lEventDirID, sUserName, sPassword
Dim sSendTo

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

sSendTo = Request.QueryString("send_to")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("view_this") = "view_this" Then
	sSendTo = Request.Form.Item("send_to")
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	sSubject = Request.Form.Item("subject")
	sMsg = Request.Form.Item("message")
	sAttachments = Request.Form.Item("attachments")
	bRecipients = False
	
	If Not sAttachments & "" = "" Then
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
	End If
	
	i = 0
	ReDim EmailArr(2, 0)
	If sSendTo = "all" Then
		sql = "SELECT EventDirID, FirstName, LastName, Email FROM EventDir ORDER BY LastName, FirstName"
	Else
		sql = "SELECT EventDirID, FirstName, LastName, Email FROM EventDir WHERE Active = 'y' "
		sql = sql & "ORDER BY LastName, FirstName"
	End If
	Set rs = conn.Execute(sql)
	If Request.Form.Item("email_all") = "on" Then
		bRecipients = True
		Do While Not rs.EOF
			EmailArr(0, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
			EmailArr(1, i) = rs(3).Value
			EmailArr(2, i) = rs(0).Value
			i = i + 1
			ReDim Preserve EmailArr(2, i)
			rs.MoveNext
		Loop
		Set rs = Nothing
	Else
		Do While Not rs.EOF
			If Request.Form.Item("event_dir_" & rs(0).Value) = "on" Then
				EmailArr(0, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
				EmailArr(1, i) = rs(3).Value
				EmailArr(2, i) = rs(0).Value
				i = i + 1
				ReDim Preserve EmailArr(2, i)
				If bRecipients = False Then bRecipients = True
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
	End If

	If bRecipients = False Then
		sErrMsg = "Please select at least one recipient."
	Else
%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
		For i = 0 to UBound(EmailArr, 2) - 1
			sSubject = Request.Form.Item("subject")
			
            sMsg = "Dear " & EmailArr(0, i) & ":" & vbCrLf & vbCrLf

			sMsg = sMsg & Request.Form.Item("message") & vbCrLf & vbCrLf
			
			If Request.Form.Item("send_login") = "on" Then
				Call GetLogin(EmailArr(2, i))
				sMsg = sMsg & "Login Information:" & vbCrLf
				sMsg = sMsg & "User Name: " & sUserName & vbCrLf
				sMsg = sMsg & "Password: " & sPassword & vbCrLf & vbCrLf
			End If
			
            sMsg = sMsg & "Sincerely: " & vbCrLf
            sMsg = sMsg & "Bob Schneider" & vbCrLf
            sMsg = sMsg & "Gopher State Events, LLC" & vbCrLf
            sMsg = sMsg & "612-720-8427"

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
'                .To = "bob.schneider@gopherstateevents.com"
				.To = EmailArr(1, i)
				If i = 0 Then .BCC = "bob.schneider@gopherstateevents.com"
				.From = "bob.schneider@gopherstateevents.com"
				
'				If IsArray(AttachArr) Then
'					For j = 0 to UBound(AttachArr)
'						.AddAttachment AttachArr(j)
'					Next
'				End If

				.Subject = sSubject
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
		Next
	
		Set cdoConfig = Nothing
	End If
End If

If sSendTo = vbNullString Then sSendTo = "active"

i = 0
ReDim EventDir(3, 0)
If sSendTo = "all" Then
	sql = "SELECT EventDirID, FirstName, LastName, Email FROM EventDir ORDER BY LastName, FirstName"
Else
	sql = "SELECT EventDirID, FirstName, LastName, Email FROM EventDir WHERE Active = 'y' "
	sql = sql & "ORDER BY LastName, FirstName"
End If
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventDir(0, i) = rs(0).Value
	EventDir(1, i) = Replace(rs(1).Value, "''","'") & " " & Replace(rs(2).Value, "''", "'")
	EventDir(2, i) = rs(3).Value
    EventDir(3, i) = MyLastEvent(rs(0).Value)
	i = i + 1
	ReDim Preserve EventDir(3, i)
	rs.MoveNext
Loop
Set rs = Nothing

i = 0
ReDim AttachArr(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Attachment FROM Attach"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	AttachArr(i) = rs(0).Value
	i = i + 1
	ReDim Preserve AttachArr(i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub GetLogin(lEventDirID)
	sql = "SELECT UserID, Password FROM EventDir WHERE EventDirID = " & lEventDirID
	Set rs = conn.Execute(sql)
	sUserName = rs(0).Value
	sPassword = rs(1).Value
	Set rs = Nothing
End Sub

Private Function MyLastEvent(lEventDirID)
    Dim rs2, sql2

    MyLastEvent = "n/a"
    Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT EventName, EventDate FROM Events WHERE EventDirID = " & lEventDirID & " ORDER BY EventDate DESC"
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then MyLastEvent = Replace(rs2(0).Value, "''", "'") & " (" & rs2(1).Value & ")"
    rs2.Close
	Set rs2 = Nothing
End Function

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Admin Group Email</title>

<!--#include file = "../../includes/js.asp" -->

</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE&copy; Group Email</h4>
			
			<%If Not sErrMsg = vbNullString then%>
				<p><%=sErrMsg%></p>
			<%End If%>
			
			<table>
				<tr>
					<td valign="top">
						<h4 class="h4">Select Recipients</h4>
						<table style="width:250px">
							<tr>
								<td>
									<form name="view_who" method="post" action="group_email.asp">
									<fieldset>
										<legend>Send To</legend>
										<input type="radio" name="send_to" id="send_to" value="active" checked>Active Event Directors<br>
										<input type="radio" name="send_to" id="send_to" value="all"
										<%If sSendTo = "all" Then%>
											checked					
										<%End If%>>All Event Directors<br>
										<input type="hidden" name="view_this" id="view_this" value="view_this">
										<input type="submit" name="submit1" id="submit1" value="View These" style="margin-left:50px;">
									</fieldset>
									</form>
								</td>
							</tr>
							<form name="send_email" method="post" action="group_email.asp?send_to=<%=sSendTo%>" onsubmit="return chkFlds()">
							<tr>
								<td>
									1)
									<input type="checkbox" name="email_all" id="email_all"> Email All
								</td>
							</tr>
							<%For i = 0 to UBound(EventDir, 2) - 1%>
								<tr>
									<td white-space:nowrap;">
										<%=i + 2%>)
										<input type="checkbox" name="event_dir_<%=EventDir(0, i)%>" id="event_dir_<%=EventDir(0, i)%>">
										<a href="mailto:<%=EventDir(2, i)%>"><%=EventDir(1, i)%> (<%=EventDir(3, i)%>)</a>
									</td>
								</tr>
							<%Next%>
						</table>
					</td>
					<td valign="top">
                        <p>Includes personal salutation and signature.</p>
						<div style="background-color:#ececec;padding: 10px;">
							<h4 class="h4">Create Message</h4>
							<table style="width:450px;">
								<tr>
									<td style="text-align:right" valign="top">
										Subject:
									</td>
									<td>
										<input type="text" name="subject" id="subject" size="50" value="<%=sSubject%>">
									</td>
								</tr>
								<tr>
									<td style="text-align:right" valign="top">
										Attach File:
									</td>
									<td>
										<select name="attachments" id="attachments" style="width:200px" size="3" multiple>
											<%For i = 0 to UBound(AttachArr) - 1%>
												<option value="<%="c:\inetpub\h51web\gopherstateevents\admin\grp_email\attachments\" & AttachArr(i)%>"><%=AttachArr(i)%></option>
											<%Next%>
										</select>
									</td>
								</tr>
								<tr>
									<td style="text-align:right;white-space:nowrap;">
										Send Login Info:
									</td>
									<td style="text-align:left">
										<input type="checkbox" name="send_login" id="send_login">
										Send site login information with email.
									</td>
								</tr>
								<tr>
									<td style="text-align:right" valign="top">
										Message:
									</td>
									<td>
										<textarea name="message" id="message" rows="10" cols="50" style="font-size:1.35em;"><%=sMsg%></textarea>
									</td>
								</tr>
								<tr>
									<td style="text-align:center" colspan="2">
										<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
										<input type="submit" name="submit2" id="submit2" value="Send">
									</td>
								</tr>
							</table>
							</form>
						</div>
					</td>
				</tr>
			</table>
		</div>
	</div>
</div>
</body>
</html>
