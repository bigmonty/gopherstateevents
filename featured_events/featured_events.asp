<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sEventDir, sEmail, sDescr, sMsg, sEventName, sLocation, sWebUrl, sPhone, sErrMsg
Dim cdoMessage, cdoConfig
Dim dEventDate
Dim bMsgSent

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%
If Request.form.Item("submit_this") = "submit_this" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'featured_events' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_featured_events") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_featured_events") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
		
		sEventDir = CleanInput(Trim(Request.Form.Item("your_name")))
		If sHackMsg = vbNullString Then sEmail = CleanInput(Trim(Request.Form.Item("email")))
		If sHackMsg = vbNullString Then sEventDir = CleanInput(Trim(Request.Form.Item("event_dir")))
        If sHackMsg = vbNullString Then sPhone = CleanInput(Trim(Request.Form.Item("phone")))
        If sHackMsg = vbNullString Then sWebUrl = Trim(Request.Form.Item("web_url"))
        If sHackMsg = vbNullString Then dEventDate = Trim(Request.Form.Item("event_date"))
        If sHackMsg = vbNullString Then sEventName = CleanInput(Trim(Request.Form.Item("event_name")))
        If sHackMsg = vbNullString Then sLocation = CleanInput(Trim(Request.Form.Item("location")))
		If sHackMsg = vbNullString Then sDescr = CleanInput(Trim(Request.Form.Item("descr")))
		
        If IsDate(dEventDate) = False Then sErrMsg = "Please select a valid date for your event date."
        If CDate(dEventDate) <= Date Then sErrMsg = "Please select a future date for your event date."

		If sHackMsg = vbNullString AND sErrMsg = vbNullString Then
    		sMsg = "A Featured Event Request" & vbCrLf & vbCrLf
            sMsg = sMsg & "Event: " & sEventName & vbCrLf
            sMsg = sMsg & "Date: " & dEventDate & vbCrLf
            sMsg = sMsg & "Location: " & sLocation & vbCrLf
            sMsg = sMsg & "Website: " & sWebURL & vbCrLf
			sMsg = sMsg & "Event Director: " & sEventDir & vbCrLf
            sMsg = sMsg & "Phone: " & sPhone & vbCrLf
			sMsg = sMsg & "Email: " & sEmail & vbCrLf & vbCrLf
			sMsg = sMsg & "Description: " & sDescr & vbCrLf 

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = "bob.schneider@gopherstateevents.com;" '& sEmail
				.From = "" & sEmail & "<" & sEmail & ">"
				.Subject = "A Featured Event Request"
				.TextBody = sMsg
				.Send
			End With
			Set cdoMessage = Nothing

            'insert into table
		    sEmail = Replace(sEmail, "'", "''")
            sEventDir = Replace(sEventDir, "'", "''")
            sEventName = Replace(sEventName, "'", "''")
            sLocation = Replace(sLocation, "'", "''")
 		    If Not sDescr = vbNullString Then sDescr = Replace(sDescr, "'", "''")

            sql = "INSERT INTO FeaturedEvents (EventName, EventDate, Location, WebURL, EventDir, Phone, Email, Descr, WhenCreated) VALUES ('" & sEventName
            sql = sql & "', '" & dEventDate & "', '" & sLocation & "', '" & sWebURL & "', '" & sEventDir & "', '" & sPhone & "', '" & sEmail
            sql = sql & "', '" & sDescr & "', '" & Now() & "')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
			
			bMsgSent = True
		End If
	End If
End If

'log this user if they are just entering the site
If Session("access_featured_events") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'featured_events')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'contact'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_featured_events")
End If

%>
<!--#include file = "../includes/clean_input.asp" -->
<%

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; Featured Events</title>
<meta name="description" content="Gopher State Events featured events utility.">
<!--#include file = "../includes/js.asp" -->

<script>
function chkFlds() {
if (document.request_feature.event_dir.value == ''||
    document.request_feature.phone.value == '' ||
    document.request_feature.event_date.value == '' ||
    document.request_feature.location.value == '' ||
    document.request_feature.event_name.value == '' ||
    document.request_feature.descr.value == '' ||
    document.request_feature.web_url.value == '' ||
    document.request_feature.email.value == '') 
{
 	alert('All fields are required!');
 	return false
 	}
else
 	return true;
}

$(function() {
    $( "#event_date" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body onload="javascript:request_feature.your_name.focus()">
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<h1 class="h1">Gopher State Events "Featured Event" Request Form</h1>
		<div class="col-md-5">
            <p class="p">
                Promoting your event is the key to increasing participation numbers.  And while bigger isn't necessarily better, especially in
                events "for a cause", higher participation numbers can increase the value of an event to it's beneficiaries.  With this in mind we have
                created our "Featured Events" utility.
            </p>

            <p>
                As GSE has grown to become a major event management "player", we are seeing a dramatic increase in our website traffic.  This is very fertile
                soil for event promotion.  People that come to our site for results, race calendars, and related reasons run races.  Putting your race in
                front of these folks can be huge in terms of growing your race.
            </p>

            <p>Listing your event with us as a "Featured Event" carries the following benefits:</p>
            <ul>
                <li>Randomly appearing presence front-and-center on our web site's home page.</li>
                <li>Banner ad on our "Results" and "Mobile Results" pages.</li>
                <li>Block ad on our individual results emails, pre-race emails, and post-race picture gallery email.</li>
                <li>Listed whereever our "Your Next Race" feature shows up (on the "Calendar" page, on this page, etc)</li>
                <li>The ad can be displayed up to 6 months from your race date.</li>
            </ul>

            <p>
                In essence, we put your event in front of as many event participants as we can...literally thousands.  Unfortunately there is a cost for 
                this because of the prime "electronic real estate" that we allocate for it.  It is significantly less for events we manage but any event 
                can get listed here.
            </p>
            <p>
                To get your event featured, just fill out the form on this page.  We will send you a price quote and request a banner and a block image.  
                We choose not to publish our pricing but it is roughly equivalent to one to three entry fees for events that we manage and equivalent 
                to about three to five entry fees for events that we do not manage.
            </p>
        </div>
        <div class="col-md-5 bg-info">
			<%If bMsgSent = True Then%>
				<p>Your featured event request has been sent and it will be responded to in less than 24 hours.  Thank you for your interest in GSE!</p>
			<%Else%>
                <%If Not sErrMsg = vbNullString Then%>
                    <p class="bg-danger"><%=sErrMsg%></p>
                <%End If%>
				<form class="form" name="request_feature" method="post" action="featured_events.asp" onSubmit="return chkFlds();">
				<div class="form-group">
					<label for="event_name">Event Name:</label>
					<input class="form-control" name="event_name" id="event_name" value="<%=sEventName%>">
				</div>
				<div class="form-group">
					<label for="event_date">Event Date:</label>
					<input class="form-control" name="event_date" id="event_date" value="<%=dEventDate%>">
				</div>
				<div class="form-group">
					<label for="location">Location:</label>
					<input class="form-control" name="location" id="location" value="<%=sLocation%>">
				</div>
				<div class="form-group">
					<label for="web_url">Website:</label>
					<input class="form-control" name="web_url" id="web_url" value="<%=sWebURL%>">
				</div>
				<div class="form-group">
					<label for="event_dir">Event Director:</label>
					<input class="form-control" name="event_dir" id="event_dir" value="<%=sEventDir%>">
				</div>
				<div class="form-group">
					<label for="email">Email:</label>
					<input class="form-control" name="email" id="email" value="<%=sEmail%>">
				</div>
				<div class="form-group">
					<label for="phone">Phone:</label>
					<input class="form-control" name="phone" id="phone" value="<%=sPhone%>">
				</div>
                <div class="form-group">
                    <label for="descr">Description (this will appear on the GSE home page):</label>
					<textarea class="form-control" name="descr" id="descr" rows="5"><%=sDescr%></textarea>
                </div>
				<div class="form-group">
					<input class="form-control" type="hidden" name="submit_this" id="submit_this" value="submit_this">
					<input class="form-control" type="submit" name="submit" id="submit" value="Send">
				</div>
				</form>
		    <%End If%>
  		</div>
	</div>
	<!--#include file = "../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
