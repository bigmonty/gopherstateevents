<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, rs, conn, rs2, sql2
Dim i, j, k
Dim iNumSent
Dim lTargetEvent, lPromoID, lRecipEvent
Dim sMessage, sOtherRecips, sPartString, sWhichPage, sSeriesName, sEventRaces
Dim Events(), PartArr(), InfoSheets(), Recips(), SendRslts(), SeriesEvents()
Dim cdoMessage, cdoConfig, objEmail, xmlhttp, EmailContents, sPageToSend
Dim dWhenSent

If Not Session("role") = "admin" Then Response.Redirect "/index.html"

lTargetEvent = Request.QueryString("target_event")
sWhichPage = Request.QueryString("which_page")

Server.ScriptTimeout = 1200

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("get_target") = "get_target" Then
    lTargetEvent = Request.Form.Item("target_event")
    sWhichPage = Request.Form.Item("which_page")
ElseIf Request.Form.Item("submit_promo") = "submit_promo" Then
	lRecipEvent = Request.Form.Item("recip_event")
    If Not Request.Form.Item("message") & "" = "" Then sMessage = Replace(Request.Form.Item("message"), "'", "''")
    If Not Request.Form.Item("location") & "" = "" Then sLocation = Replace(Request.Form.Item("location"), "'", "''")
    sOtherRecips = Request.Form.Item("other_recips")

	'add this event to the promo table
	sql = "INSERT INTO PromoEmail(TargetEvent, WhenSent, Message, EventRecips, OtherRecips) VALUES (" 
    sql = sql & lTargetEvent & ", '" & Now() & "', '" & sMessage & "', '" & lRecipEvent & "', '" & sOtherRecips & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

    'get promo id
    sql = "SELECT PromoEmailID FROM PromoEmail WHERE TargetEvent = " & lTargetEvent & " ORDER BY PromoEmailID DESC"
    Set rs = conn.Execute(sql)
    lPromoID = rs(0).Value
    Set rs = Nothing

    If sWhichPage = "series" Then
        'get series name
        sql = "SELECT s.SeriesName FROM Series s INNER JOIN SeriesEvents se ON s.SeriesID = se.SeriesID WHERE se.EventID = " & lTargetEvent
        Set rs = conn.Execute(sql)
        sSeriesName = rs(0).Value
        Set rs = Nothing
    End If

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lRecipEvent
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sEventRaces = sEventRaces & rs(0).Value & ", "
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)
	
    j = 0
    ReDim Recips(1, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT p.FirstName, p.LastName, p.Email FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID "
    sql = sql & "WHERE pr.RaceID IN (" & sEventRaces & ") ORDER BY p.LastName, p.FirstName"
	rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
		If Not rs(2).Value & "" = "" Then
            Recips(0, j) = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
            Recips(1, j) = rs(2).Value
            j = j + 1
            ReDim Preserve Recips(1, j)
        End If
        rs.MoveNext
    Loop
	rs.Close
	Set rs = Nothing

    If Len(sOtherRecips) > 0 Then
        sOtherRecips = Trim(sOtherRecips)

        If Right(sOtherRecips, 1) = ";" Or Right(sOtherRecips, 1) = "," Then sOtherRecips = Left(sOtherRecips, Len(sOtherRecips) - 1)

        For i = 1 To Len(sOtherRecips)
            If i = Len(sOtherRecips) Then
                Recips(0, j) = "Other Recipient"
                Recips(1, j) = sPartString
                j = j + 1
                ReDim Preserve Recips(1, j)
           Else
                If Mid(sOtherRecips, i, 1) = ";" Or Mid(sOtherRecips, i, 1) = "," Then
                    Recips(0, j) = "Other Recipient"
                    Recips(1, j) = sPartString
                    j = j + 1
                    ReDim Preserve Recips(1, j)

                    sPartString = vbNullString
                Else
                    sPartString = sPartString & Mid(sOtherRecips, i, 1)
                End If
            End If
        Next
    End If

'    If sWhichPage = "series" Then  
'        sPageToSend = "http://www.gopherstateevents.com/admin/events_promo/series_promo.asp?promo_id=" & lPromoID
'    Else              
'        sPageToSend = "http://www.gopherstateevents.com/admin/events_promo/promo_page.asp?promo_id=" & lPromoID
'    End If

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

'    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
'	    xmlhttp.open "GET", sPageToSend, false
'	    xmlhttp.send ""
'	    EmailContents = xmlhttp.responseText
'    Set xmlhttp = nothing

    iNumSent = 0

    For i = 0 To UBound(Recips, 2) - 1
		If ValidEmail(Recips(1, i)) = True Then
            If DontSend(Recips(1, i)) = False Then
                If AlreadySent(Recips(1, i)) = False Then
                    iNumSent = CInt(iNumSent) + 1

			        Set cdoMessage = CreateObject("CDO.Message")
			        With cdoMessage
				        Set .Configuration = cdoConfig
				        .To = Recips(1, i)
                        If iNumSent = 1 Then .BCC = "bob.schneider@gopherstateevents.com"
				        .From = "bob.schneider@gopherstateevents.com"

                        If sWhichPage = "series" Then  
   				            .Subject = "The " & sSeriesName & " Continues!"
                        Else              
   				            .Subject = "Your next race?"
                        End If
                    
                        .HTMLBody = EmailContents
                        .CreateMHTMLBody "http://www.gopherstateevents.com/admin/events_promo/promo_page.asp?promo_id=" & lPromoID
				        .Send
			        End With
			        Set cdoMessage = Nothing

                    sql = "INSERT INTO PromoRecips(TargetEvent, RecipName, Email, WhenSent, PromoEmailID) VALUES (" & lTargetEvent & ", '" 
                    sql = sql & Replace(Recips(0, i), "'", "''") & "', '" & Recips(1, i) & "', '" & Now() & "', " & lPromoID & ")"
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing

                    Set rs = Server.CreateObject("ADODB.Recordset")
                    sql = "SELECT NumSent FROM PromoEmail WHERE PromoEmailID = " & lPromoID
                    rs.Open sql, conn, 1, 2
                    rs(0).Value = iNumSent
                    rs.Update
                    rs.Close
                    Set rs = Nothing
               End If
		    End If
        End If
	Next
    
    Set cdoConfig = Nothing

    sMessage = vbNullString
    sOtherRecips = vbNullString
End If

If CStr(lTargetEvent) = vbNullString Then lTargetEvent = 0

i = 0
ReDim Events(2, 0)
sql = "SELECT EventID, EventName, EventDate, Location FROM Events WHERE EventDate > '" & Date - 375 & "' ORDER BY EventDate"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & " - " & rs(3).Value & ")"
    Events(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve Events(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

ReDim SendRslts(2, 0)
If Not CLng(lTargetEvent) = 0 Then
    'get event recips
    i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventRecips, WhenSent, NumSent FROM PromoEmail WHERE TargetEvent = " & lTargetEvent 
	rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        SendRslts(0, i) = GetEventName(rs(0).Value)
        SendRslts(1, i) = rs(1).Value
        SendRslts(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve SendRslts(2, i)
        rs.MoveNext
    Loop
  	rs.Close
	Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Comments FROM Events WHERE EventID = " & lTargetEvent 
	rs.Open sql, conn, 1, 2
    If Not rs(0).Value & "" = "" Then sMessage = Replace(rs(0).Value, "''", "'")
  	rs.Close
	Set rs = Nothing
End If

i = 0
Dim MyDirectory, MyDocs, MyFile
ReDim InfoSheets(0)
Set MyDirectory=Server.CreateObject("Scripting.FileSystemObject" )
Set MyDocs=MyDirectory.GetFolder("c:/inetpub/h51web/gopherstateevents/admin/events_promo")
For each MyFile in MyDocs.files
    If Not Right(MyFile.Name, 4) = ".asp" Then
        InfoSheets(i) = MyFile.Name
        i = i + 1
        ReDim Preserve InfoSheets(i)
    End If
Next

Private Function GetEventName(lThisEvent)
    sql2 = "SELECT EventName FROM Events WHERE EventID = " & lThisEvent
    Set rs2 = conn.Execute(sql2)
    GetEventName = Replace(rs2(0).Value, "''", "'")
     Set rs2 = Nothing
End function

Private Function AlreadySent(sThisEmail)
    AlreadySent = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PromoRecipsID FROM PromoRecips WHERE Email = '" & sThisEmail & "' AND TargetEvent = " & lTargetEvent
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then AlreadySent = True
    rs.Close
    Set rs = Nothing
End Function

Private Function DontSend(sThisEmail)
    DontSend = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT DontSendID FROM DontSend WHERE Email = '" & sThisEmail & "' AND (DontSend = 'all' OR DontSend = 'promo')"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then DontSend = True
    rs.Close
    Set rs = Nothing
End Function

%>
<!--#include file = "../../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Promotional Email Utility</title>

<script>
function chkPage(){
 	if (document.select_event.which_page.value == '')
		{
  		alert('Please select a page to send.');
  		return false
  		}
	else
   		return true
}

function chkFlds(){
 	if (document.send_promo.recip_events.value == '')
		{
  		alert('Please select event recipients.');
  		return false
  		}
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <h4 class="h4">GSE Event Promotion Utility</h4>
		
            <div class="row">
                <form class="form-inline" role="form" name="select_event" method="post" action="email_event_promo.asp" onsubmit="return chkPage();">
                <label for="target_event">Promote:</span>&nbsp;
                <select class="form-control" name="target_event" id="target_event">
                    <option value="">&nbsp;</option>
                    <%For i = 0 To UBound(Events, 2) - 1%>
                        <%If CDate(Events(2, i)) > Date Then%>
                            <%If CLng(lTargetEvent) = CLng(Events(0, i)) Then%>
                                <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
                            <%Else%>
                                <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                            <%End If%>
                        <%End If%>
                    <%Next%>
                </select>
                &nbsp;&nbsp;<label for="which_page">Email Type:</span>&nbsp;
                <select class="form-control" name="which_page" id="which_page">
                    <option value="">&nbsp;</option>
                    <%If sWhichPage = "promo" Then%>
                        <option value="promo" selected>Promo</option>
                        <option value="series">Series</option>
                    <%ElseIf sWhichPage = "series" Then%>
                        <option value="promo">Promo</option>
                        <option value="series"selected>Series</option>
                    <%Else%>
                        <option value="promo">Promo</option>
                        <option value="series">Series</option>
                    <%End If%>
                </select>
                <input type="hidden" name="get_target" id="get_target" value="get_target">
                <input class="form-control" type="submit" name="submit2" id="submit2" value="Get This">
                </form>
            </div>
            <div class="row">
                <div class="col-sm-6">
                    <form class="form" name="send_promo" method="post" action="email_event_promo.asp?target_event=<%=lTargetEvent%>&amp;which_page=<%=sWhichPage%>" 
                    onsubmit="return chkFlds();">
                    <div class="form-group">
                        <label for="recip_event">Select Event Recipients:</label>
                        <select class="form-control" name="recip_event" id="recip_event" size="10">
                            <%For i = 0 To UBound(Events, 2) - 1%>
                                <%If CDate(Events(2, i)) <= Date Then%>
                                    <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>   
                                <%End If%>
                            <%Next%>
                        </select>  
                    </div>  
                    <div class="form-group">
                        <label for="reg_fee">Other Recips:</label>
                        <textarea class="form-control" name="other_recips" id="other_recips" rows="3">bob.schneider@gopherstateevents.com;</textarea>
                    </div>
                    <div class="form-group">
                        <label for="message">Message:</label>
                        <textarea class="form-control" name="message" id="message" rows="8"><%=sMessage%></textarea>
                    </div>
                    <div class="form-group">   
                        <input type="hidden" name="submit_promo" id="submit_promo" value="submit_promo">
                        <input class="form-control" type="submit" name="submit1" id="submit1" value="Send Promotional Email">
                    </div>
                    </form>
                </div>
                 <div class="col-sm-6">
                    <h5 class="h5">Send Results</h5>
                    <table class="table table-striped">
                        <tr>
                            <th>No.</th>
                            <th>Event</th>
                            <th>When Sent</th>
                            <th>Num</th>
                        </tr>

                        <%For k = 0 To UBound(SendRslts, 2) - 1%>
                            <tr>
                                <td><%=k + 1%></td>
                                <%For j = 0 To 2%>
                                    <td><%=SendRslts(j, k)%></td>
                                <%Next%>
                            </tr>
                        <%Next%>
                    </table>
                </div>
           </div>
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
