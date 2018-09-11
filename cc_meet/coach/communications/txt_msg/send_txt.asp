<%@ Language=VBScript%>

<%
Option Explicit

Dim sql, rs, conn, sql2, rs2
Dim i, x, j, k
Dim lPartID, lThisProvider, lCoachProvider, lMyProvider, lCoachID
Dim Recips()
Dim sMsg, sErrMsg, sCoachNumber, sThisNumber, sMyNumber, sTeamIDs
Dim cdoMessage, cdoConfig
Dim bHasError

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

sMsg = Request.Form.Item("Message")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 30
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Session("role") = "coach" Then
    lCoachID = Session("my_id")
Else
    lCoachID = Session("team_coach_id")
End If

'get team staff
Set rs = Server.CreateObject("ADODB.Recordset")
If Session("role") = "coach" Then
    sql = "SELECT TeamsID FROM Teams WHERE CoachesID = " & Session("my_id")
Else
    sql = "SELECT TeamsID FROM Teams WHERE CoachesID = " & Session("team_coach_id")
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sTeamIDs = sTeamIDs & rs(0).Value & ", "
	rs.MoveNext
Loop
rs.Close
Set rs=Nothing

sTeamIDs = Left(sTeamIDs, Len(sTeamIDs) - 2)

'get my wireless
bHasError = False
Set rs=Server.CreateObject("ADODB.Recordset")
If Session("role") = "coach" Then
	sql = "SELECT CellProvidersID, CellPhone FROM Coaches WHERE CoachesID = " & Session("my_id")
Else
    sql = "SELECT Provider, MobilePhone FROM TeamStaff WHERE TeamStaffID = " & Session("my_id")
End If
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    If CLng(rs(0).Value) > 0 And rs(1).Value & "" <> "" Then
	    lMyProvider = rs(0).Value
	    sMyNumber = Replace(rs(1).Value, "-", "")
    Else
        bHasError = True
    End If
Else
    bHasError = True
End If
rs.Close
Set rs = Nothing

If bHasError = True Then
	sErrMsg = "I am sorry.  You may not send a text message until your wireless information is in our database.  You may enter this information by "
	sErrmsg = sErrMsg & "clicking the 'My Wirelss Info' on your Welcome page."
End If

If sErrMsg = vbNullString Then
	'get head coach wireless if a staff member
    If Not Session("role") = "coach" Then
	    Set rs=Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT CellProvidersID, CellPhone FROM Coaches WHERE CoachesID = " & lCoachID
	    rs.Open sql, conn, 1, 2
	    If rs.RecordCount > 0 Then
		    lCoachProvider = rs(0).Value
		    sCoachNumber = Replace(rs(1).Value, "-", "")
	    End If
	    rs.Close
	    Set rs = Nothing
    End If
    	
	'get participants for the text
	i = 0	
	ReDim Recips(2, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RosterID, CellProvidersID, CellPhone FROM Roster WHERE TeamsID IN (" & sTeamIDs & ") AND Archive = 'n'"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		If Request.Form.Item("all_parts") = "on" Then
			If Request.Form.Item("part_exists_" & rs(0).Value) = "y" Then
				Recips(0, i) = rs(0).Value
				Recips(1, i) = rs(1).Value
				Recips(2, i) = Replace(rs(2).Value, "-", "")
				i = i + 1
				ReDim Preserve Recips(2, i)
			End If
		Else
			If Request.Form.Item("part_" & rs(0).Value) = "on" Then
				Recips(0, i) = rs(0).Value
				Recips(1, i) = rs(1).Value
				Recips(2, i) = Replace(rs(2).Value, "-", "")
				i = i + 1
				ReDim Preserve Recips(2, i)
			End If
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs =Nothing
	
	'get staff recipients for the text
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT TeamStaffID, Provider, MobilePhone FROM TeamStaff WHERE CoachesID = " & lCoachID
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		If Request.Form.Item("all_staff") = "on" Then
			If Request.Form.Item("staff_exists_" & rs(0).Value) = "y" Then
				Recips(0, i) = rs(0).Value
				Recips(1, i) = rs(1).Value
				Recips(2, i) = Replace(rs(2).Value, "-", "")
				i = i + 1
				ReDim Preserve Recips(2, i)
			End If
		Else
			If Request.Form.Item("staff_" & rs(0).Value) = "on" Then
				Recips(0, i) = rs(0).Value
				Recips(1, i) = rs(1).Value
				Recips(2, i) = Replace(rs(2).Value, "-", "")
				i = i + 1
				ReDim Preserve Recips(2, i)
			End If
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs =Nothing
	
	'get team contact recipients for the text
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT TeamContactsID, CellProvidersID, CellPhone FROM TeamContacts WHERE TeamsID IN (" & sTeamIDs  & ")"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		If Request.Form.Item("all_contacts") = "on" Then
			If Request.Form.Item("contact_exists_" & rs(0).Value) = "y" Then
				Recips(0, i) = rs(0).Value
				Recips(1, i) = rs(1).Value
				Recips(2, i) = Replace(rs(2).Value, "-", "")
				i = i + 1
				ReDim Preserve Recips(2, i)
			End If
		Else
			If Request.Form.Item("contacts_" & rs(0).Value) = "on" Then
				Recips(0, i) = rs(0).Value
				Recips(1, i) = rs(1).Value
				Recips(2, i) = Replace(rs(2).Value, "-", "")
				i = i + 1
				ReDim Preserve Recips(2, i)
			End If
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
			
%>
<!--#include file = "../../../../includes/cdo_connect.asp" -->
<%
	
	Set cdoMessage = Server.CreateObject("CDO.Message")
	Set cdoMessage.Configuration = cdoConfig

	'first send to head coach
    If Not Session("role") = "coach" Then
	    If Not lCoachProvider & "" = "" Then
		    With cdoMessage
			    .From = sMyNumber & GetSendURL(lMyProvider)
			    .To = sCoachNumber & GetSendURL(lCoachProvider)
			    .TextBody = sMsg
			    .Send
		    End With
	    End If
    End If
    	
	'now send to designated recipients
	For i = 0 To UBound(Recips, 2) - 1
		If Not Recips(1, i) & "" = "" Then
			With cdoMessage
				.From = sMyNumber & GetSendURL(lMyProvider)
				.To = Recips(2, i) & GetSendURL(Recips(1, i))
				.TextBody = sMsg
				.Send
			End With
		End If
	Next
		
	Set cdoMessage = Nothing
End If

Private Function GetSendURL(lProviderID)
	If Not lProviderID & "" = ""  Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT SendURL FROM CellProviders WHERE CellProvidersID = " & lProviderID
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then GetSendURL = rs(0).Value
		Set rs = Nothing
	End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>Gopher State Events&reg; Send Group Email</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<h4 class="h4">Gopher State Events<sup>&reg;</sup> Text Message Utility</h4>	
			<%If sErrMsg = vbNullString Then%>
				<p>Your text message has been sent.</p>
			<%Else%>
				<p><%=sErrMsg%></p>
			<%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
