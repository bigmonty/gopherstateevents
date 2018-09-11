<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j
Dim lEventID, lProvider
Dim sEventName, sEventRaces, sMobileNum, sProvider, sMessage
Dim dEventDate
Dim PartArray(), SendTo()

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    sEventRaces = sEventRaces & rs(0).Value & ", "
	rs.MoveNext
Loop
Set rs = Nothing

If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

i = 0
ReDim PartArray(8, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, rc.Bib, p.Gender, rc.Age, rc.RaceID FROM Participant p INNER JOIN PartRace rc "
sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID IN (" & sEventRaces & ") ORDER BY p.LastName, p.FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Call GetMyMobile(rs(0).Value)
    If Not sMobileNum = vbNullString Then
        PartArray(0, i) = rs(0).Value
        PartArray(1, i) = Replace(rs(1).Value, "''", "'")
        PartArray(2, i) = Replace(rs(2).Value, "''", "'")
        PartArray(3, i) = rs(3).Value
        PartArray(4, i) = rs(4).Value
        PartArray(5, i) = rs(5).Value
        PartArray(6, i) = GetRaceName(rs(6).Value)
        PartArray(7, i) = sMobileNum
        PartArray(8, i) = sProvider

        i = i + 1
        ReDim Preserve PartArray(8, i)
    End If
    rs.MoveNext
Loop
rs.Close
Set rs=Nothing

If Request.Form.Item("submit_this") = "submit_this" Then
    sMessage = Request.Form.Item("message")

    Dim cdoConfig, cdoMessage
    Set cdoConfig = CreateObject("CDO.Configuration")
    With cdoConfig.Fields
        .Item(cdoSendUsingMethod) = cdoSendUsingPort
        .Item(cdoSMTPServer) = "smtp.mandrillapp.com"
        .Item(cdoSMTPAuthenticate) = 1
        .Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
        .Item(cdoSendUsername) = "bob.schneider@gopherstateevents.com"
        .Item(cdoSendPassword) = "H49iry1SZKdY7PQ5afpfyg"
        .Update
    End With

    For i = 0 To UBound(PartArray, 2) - 1
        sMobileNum = vbNullString
        lProvider = 0

        If Request.Form.Item("send_" & PartArray(0, i)) = "on" Then
	        sql = "SELECT MobileNumber, CellProvider FROM MobileSettings WHERE EventID = " & lEventID & " AND PartID = " & PartArray(0, i)
            Set rs = conn.Execute(sql)
            sMobileNum = rs(0).Value
            lProvider = rs(1).Value
            Set rs = Nothing

            Set cdoMessage = Server.CreateObject("CDO.Message")
            Set cdoMessage.Configuration = cdoConfig
		    With cdoMessage
                .From = "bob.schneider@gopherstateevents.com"
			    .To = sMobileNum & GetSendURL(lProvider)
			    .TextBody = sMessage
			    .Send
		    End With
	        Set cdoMessage = Nothing
        End If
    Next
    Set cdoConfig = Nothing
End If

Private Function GetRaceName(lWhichRace)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT RaceName FROM RaceData WHERE RaceID = " & lWhichRace
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetRaceName = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

Private Sub GetMyMobile(lMyID)
    lProvider = 0
    sMobileNum = vbNullString
    sProvider = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT MobileNumber, CellProvider FROM MobileSettings WHERE EventID = " & lEventID & " AND PartID = " & lMyID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then 
        sMobileNum = rs2(0).Value
        lProvider = rs2(1).Value
    End If
    rs2.Close
    Set rs2 = Nothing

    If Not CLng(lProvider) = 0 Then
	    sql2 = "SELECT Provider FROM CellProviders WHERE CellProvidersID = " & lProvider
        Set rs2 = conn2.Execute(sql2)
        sProvider = rs2(0).Value
        Set rs2 = Nothing
    End If
End Sub

Private Function GetSendURL(lProviderID)
	If Not CStr(lProviderID) & "" = ""  Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT SendURL FROM CellProviders WHERE CellProvidersID = " & lProviderID
		rs.Open sql, conn2, 1, 2
		If rs.RecordCount > 0 Then GetSendURL = rs(0).Value
        rs.Close
		Set rs = Nothing
	End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Participant SMS Data</title>

<script>
    function chkFlds() {
        if (document.send_sms.message.value == '') {
            alert('You must supply a message!');
            return false;
        }
        else
            return true;
    }
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h2 class="h2"><%=sEventName%> Participant SMS Data</h2>

			<!--#include file = "../../includes/event_nav.asp" -->
			<!--#include file = "part_nav.asp" -->

            <form role="form" class="form-horizontal" name="send_sms" method="Post" action="part_sms.asp?event_id=<%=lEventID%>" 
                onsubmit="return chkFlds();">
            <div class="form-group">
				<label for="message" class="control-label col-xs-1">Message:</label>
				<div class="col-xs-11">
                    <textarea class="form-control" name="message" id="message" rows="4"></textarea>
                </div>
            </div>

			<table class="table table-striped">
				<tr>
					<th>No</th>
					<th>Name</th>
					<th>Bib</th>
					<th>M/F</th>
					<th>Age</th>
					<th>Race</th>
					<th>Mobile</th>
					<th>Provider</th>
                    <th>Send</th>
				</tr>

				<%For i = 0 to UBound(PartArray, 2) - 1%>
					<tr>
						<td><%=i+1%>)</td>
                        <td><%=PartArray(2, i)%>, <%=PartArray(1, i)%></td>
                        <td><%=PartArray(3, i)%></td>
                        <td><%=PartArray(4, i)%></td>
                        <td><%=PartArray(5, i)%></td>
						<td><%=PartArray(6, i)%></td>
						<td><%=PartArray(7, i)%></td>
                        <td><%=PartArray(8, i)%></td>
                        <td><input type="checkbox" name="send_<%=PartArray(0, i)%>" id="send_<%=PartArray(0, i)%>"></td>
					</tr>
				<%Next%>
			</table>
            <div class="form-group">
                <input class="form-control" type="hidden" name="submit_this" id="submit_this" value="submit_this">
                <input class="form-control" type="submit" name="submit1" id="submt1" value="Send Text">
            </div>
            </form>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>