<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim lEventID, lMeetID
Dim i
Dim sEventName, sLocation
Dim RaceGallery(), Events(), Meets()
Dim dEventDate

lEventID = Request.QueryString("event_id")
lMeetID = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get fitness events
i = 0
ReDim Events(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate <= '" & Date & "' AND EventDate > '9/1/2013' ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    If HasGallery(rs(0).Value, "fitness") = "y" Then
	    Events(0, i) = rs(0).Value
	    Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
        Events(2, i) = "fitness"
	    i = i + 1
	    ReDim Preserve Events(2, i)
    End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get cc events
i = 0
ReDim Meets(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate <= '" & Date & "' AND MeetDate > '9/1/2013' ORDER By MeetDate DESC"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    If HasGallery(rs(0).Value, "cc") = "y" Then
	    Meets(0, i) = rs(0).Value
	    Meets(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
        Meets(2, i) = "cc"
	    i = i + 1
	    ReDim Preserve Meets(2, i)
    End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_order") = "submit_order" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'finish_pix' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_finish_pix") = "y"
	rs.Close
	Set rs = Nothing

    'send email
	If Session("access_finish_pix") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
        Dim iBibNum
        Dim sEmail, sMsg
        Dim cdoMessage, cdoConfig

        iBibNum = CleanInput(Trim(Request.Form.Item("bib_num")))
        If sHackMsg = vbNullString Then sEmail = CleanInput(Trim(Request.Form.Item("email")))

        If sHackMsg = vbNullString Then
            'write to table
            If Not CLng(lEventID) = 0 Then
                sql = "INSERT INTO MediaOrder(BibNum, Email, WhenOrdered, IPAddress, EventID, MediaType) VALUES (" & iBibNum & ", '" & sEmail & "', '" 
                sql = sql & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', " & lEventID & ", 'both')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            Else
                sql = "INSERT INTO MediaOrder(BibNum, Email, WhenOrdered, IPAddress, EventID, MediaType) VALUES (" & iBibNum & ", '" & sEmail & "', '" 
                sql = sql & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', " & lMeetID & ", 'both')"
                Set rs = conn2.Execute(sql)
                Set rs = Nothing
            End If

			sMsg = "Thank you for ordering finish line media from Gopher State Events.  We have already begun processing your order.  The details of "
            sMsg = sMsg & "your order can be found below. Please verify that they are correct:" & vbCrLf & vbCrLf
			
			sMsg = sMsg & "Event Name: " & EventName() & vbCrLf
			sMsg = sMsg & "Bib Number: " & iBibNum & vbCrLf & vbCrLf

            sMsg = sMsg & "You will receive a link for online payment shortly.  Once payment is received your order will be completed and sent to you "
            sMsg = sMsg & "via the email address that you have supplied. " & vbCrLF & vbCrLf

            sMsg = sMsg & "Sincerely; " & vbCrLf
            sMsg = sMsg & "Bob Schneider " & vbCrLf
            sMsg = sMsg & "Gopher State Events, LLC " & vbCrLf
            sMsg = sMsg & "612.720.8427 " & vbCrLf

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = "bob.schneider@gopherstateevents.com;bob.bakken@gopherstateevents.com;" & sEmail
				.From = sEmail
				.Subject = "GSE Media Order"
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing
		End If

	    sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'finish_pix'"
	    Set rs = conn.Execute(sql)
	    Set rs = Nothing

	    Session.Contents.Remove("access_finish_pix")
	End If
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
    lEventID = Request.Form.Item("events")

    If CStr(lEventID) = vbNullString Then lEventID = 0
ElseIf Request.Form.Item("submit_meet") = "submit_meet" Then
    lMeetID = Request.Form.Item("meets")

    If CStr(lMeetID) = vbNullString Then lMeetID = 0
End If

'log this user if they are just entering the site
If Session("access_finish_pix") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'finish_pix')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

If CStr(lEventID) = vbNullString Then lEventID = 0
If CStr(lMeetID) = vbNullString Then lMeetID = 0

'get event information
ReDim RaceGallery(2, 0)

If Not CLng(lEventID) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventName, EventDate, Location FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then
        sEventName = Replace(rs(0).Value, "''", "'")
	    dEventDate = rs(1).Value
	    sLocation = rs(2).Value
    End If
	rs.Close
	Set rs = Nothing
	
	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceGalleryID, GalleryName, EmbedLink FROM RaceGallery WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While NOt rs.EOF
	    RaceGallery(0, i) = rs(0).Value
	    RaceGallery(1, i) = Replace(rs(1).Value, "''", "'")
        RaceGallery(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve RaceGallery(2, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

If Not CLng(lMeetID) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MeetName, MeetDate, Location FROM Meets WHERE MeetsID = " & lMeetID
	rs.Open sql, conn2, 1, 2
	sEventName = Replace(rs(0).Value, "''", "'")
	dEventDate = rs(1).Value
	sLocation = rs(2).Value
	rs.Close
	Set rs = Nothing
	
	i = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceGalleryID, GalleryName, EmbedLink FROM RaceGallery WHERE MeetsID = " & lMeetID
    rs.Open sql, conn2, 1, 2
    Do While NOt rs.EOF
	    RaceGallery(0, i) = rs(0).Value
	    RaceGallery(1, i) = Replace(rs(1).Value, "''", "'")
        RaceGallery(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve RaceGallery(2, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function EventName()
	Set rs = Server.CreateObject("ADODB.Recordset")

    If Not CLng(lEventID) = 0 Then
        sql = "SELECT EventName FROM Events WHERE EventID = " & lEventID
        rs.Open sql, conn, 1, 2
    Else
        sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lMeetID
        rs.Open sql, conn2, 1, 2
    End If

    EventName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End Function

Private Function HasGallery(lThisEvent, sThisEventType)
    HasGallery = "n"

    Set rs2 = Server.CreateObject("ADODB.Recordset")

    If sThisEventType = "fitness" Then
        sql2 = "SELECT RaceGalleryID FROM RaceGallery WHERE EventID = " & lThisEvent
        rs2.Open sql2, conn, 1, 2
    Else
        sql2 = "SELECT RaceGalleryID FROM RaceGallery WHERE MeetsID = " & lThisEvent
        rs2.Open sql2, conn2, 1, 2
    End If
        
    If rs2.RecordCount > 0 Then HasGallery = "y"
    rs2.Close
    Set rs2 = Nothing
End Function

%>
<!--#include file = "../includes/clean_input.asp" -->
<%
Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events Galleries</title>
<meta name="description" content="Gopher State Events race videos.">
<!--#include file = "../includes/js.asp" -->

<script>
function chkFlds() {
 	if (document.order_video.which_vid.value == '' || 
 	    document.order_video.clip_start.value == '' ||
	 	document.order_video.bib_num.value == ''|| 
	 	document.order_video.email.value == '')
		{
  		alert('To get you your media we need all fields filled out.');
  		return false
  		}
 	else
		if (isNaN(document.order_video.bib_num.value))
    		{
			alert('The bib number can not contain non-numeric values');
			return false
			}
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-md-10">
			<h1 style="margin:5px 0 15px 0;padding:5px;font-size:1.1em;">Gopher State Events Finish Line Pix</h1>
		
            <div style="float:left;width: 400px;">
                <h4 style="margin-top: 0;margin-bottom: 10px;">Fitness Event Pix</h4>

		        <form name="which_event" method="post" action="finish_pix.asp?event_id=<%=lEventID%>&amp;event_type=fitness">
		        <span style="font-weight:bold;">Event:</span>
		        <select name="events" id="events" onchange="this.form.get_video.click()" style="font-size:0.85em;">
			        <option value="">&nbsp;</option>
			        <%For i = 0 to UBound(Events, 2) - 1%>
				        <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
					        <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
				        <%Else%>
					        <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
				        <%End If%>
			        <%Next%>
		        </select>
		        <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
		        <input type="submit" name="get_video" id="get_video" value="View" style="font-size:0.8em;">
		        </form>
		    </div>
            <div style="margin-left:430px;">
                <h4 style="margin-bottom: 10px;">CC/Nordic Pix</h4>

		        <form name="which_meet" method="post" action="finish_pix.asp?meet_id=<%=lMeetID%>&amp;event_type=ccmeet">
		        <span style="font-weight:bold;">Meet:</span>
		        <select name="meets" id="meets" onchange="this.form.get_video1.click()" style="font-size:0.85em;">
			        <option value="">&nbsp;</option>
			        <%For i = 0 to UBound(Meets, 2) - 1%>
				        <%If CLng(lMeetID) = CLng(Meets(0, i)) Then%>
					        <option value="<%=Meets(0, i)%>" selected><%=Meets(1, i)%></option>
				        <%Else%>
					        <option value="<%=Meets(0, i)%>"><%=Meets(1, i)%></option>
				        <%End If%>
			        <%Next%>
		        </select>
		        <input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
		        <input type="submit" name="get_video1" id="get_video1" value="View" style="font-size:0.8em;">
		        </form>
            </div>

		    <%If UBound(RaceGallery, 2) > 0 Then%>
                <br>
                <table>
                    <%For i = 0 To UBound(RaceGallery, 2) - 1%>
                        <tr>
                            <th style="padding-right: 10px;" valign="top"><%=RaceGallery(1, i)%>:</th>
                            <td><%=RaceGallery(2, i)%></td>
                        </tr>
                    <%Next%>
                </table>

		        <div style="padding:0;position:absolute;left:840px;width:150px;top:200px;background:none;background-color:#cde6ff;font-size: 0.75em;">
<!-- 
                   <h4  style="background-color: #80bfff;padding-left: 5px;margin:4px;color: #000;">Media Order Form</h4>
                    <p style="margin: 4px 4px 0 4px;padding: 2px 2px 0 2px;">
                        <a href="javascript:pop('/graphics/demo_pic.png',477,640)"><img src="/graphics/demo_pic.png" 
                            style="width: 75px;float: right;margin-bottom:10px;" alt="Demo"></a>
                            Order 5-sec video clip and a finish line pic for $10.
                        <br><br>
                        <a href="javascript:pop('http://youtu.be/s7hNfF26vBw',1024,768)" style="font-weight: bold;">View Sample Video</a>
                    </p>
                    <div style="clear:both;"></div>

                    <form name="order_video" method="post" action="finish_pix.asp?event_id=<%=lEventID%>&amp;meet_id=<%=lMeetID%>" onsubmit="return chkFlds();">
                    <table style="background-color:#80bfff;margin:0 4px 4px 4px;">
                        <tr><th>Bib No:</th><td><input type="text" name="bib_num" id="bib_num" size="3"></td></tr>
                        <tr><th>Email:</th><td><input type="text" name="email" id="email" size="25"></td></tr>
                        <tr>
                            <td colspan="2" style="text-align: center;">
                                <input type="hidden" name="submit_order" id="submit_order" value="submit_order">
                                <input type="submit" name="submit4x" id="submit4x" value="Order Video">
                            </td>
                        </tr>
                    </table>
                    </form>
                    <p style="margin: 0 4px 0 4px;padding: 2px;font-size: 0.8em;">Your order will incude a finish line picture and a video.  We will 
                        verify the order and send an online payment linkr.  Once payment is received we will email your media.</p>
-->
   		        <%End If%>
            </div>
	    </div>	
    </div>
	<!--#include file = "../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>