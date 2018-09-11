<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventDirID, lGSEEventID
Dim sEventDir, sEmail, sDescr, sMsg, sEventName, sLocation, sWebUrl, sPhone, sErrMsg
Dim Events
Dim cdoMessage, cdoConfig
Dim dEventDate

If Not Session("role") = "admin" Then 
    If Not Session("role") = "event_dir" Then Response.Redirect "http://www.google.com"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
If Session("role") = "event_dir" Then
    sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '" & Date & "' AND EventDirID = " & Session("my_id") & " ORDER By EventDate"
Else
    sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '" & Date & "' ORDER By EventDate"
End If
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    Events = rs.GetRows()
Else
    ReDim Events(2, 0)
End If
rs.Close
Set rs = Nothing

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
If Request.form.Item("submit_this") = "submit_this" Then
    lGSEEventID =Request.Form.Item("gse_event")
	sDescr = Trim(Request.Form.Item("descr"))

    If CStr(lGSEEventID) = vbNullString Then
        sErrMsg = "You must select an event."
    Else
        'get event info
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventName, EventDate, Location, WebSite, EventDirID FROM Events WHERE EventID = " & lGSEEventID
        rs.Open sql, conn, 1, 2
        sEventName = Replace(rs(0).Value, "''", "'")
        dEventDate = rs(1).Value
        If Not rs(2).Value & "" = "" Then sLocation = Replace(rs(2).Value, "''", "'")
        sWebURL = rs(3).Value
        lEventDirID = rs(4).Value
        rs.Close
        Set rs = Nothing

        'get event dir info
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT FirstName, LastName, Email, Phone FROM EventDir WHERE EventDirID = " & lEventDirID
        rs.Open sql, conn, 1, 2
        sEventDir = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
        sEmail = rs(2).Value
        sPhone = rs(3).Value
        rs.Close
        Set rs = Nothing

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
		    .To = "bob.schneider@gopherstateevents.com;"
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

        sql = "INSERT INTO FeaturedEvents (EventName, EventDate, Location, WebURL, EventDir, Phone, Email, Descr, WhenCreated, EventID) VALUES ('" 
        sql = sql & sEventName & "', '" & dEventDate & "', '" & sLocation & "', '" & sWebURL & "', '" & sEventDir & "', '" & sPhone & "', '" & sEmail
        sql = sql & "', '" & sDescr & "', '" & Now() & "', " & lGSEEventID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        Response.Redirect "featured_events.asp"
    End If
End If

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Featured Events</title>
<meta name="description" content="Gopher State Events featured events utility.">
</head>

<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<%If Session("role") = "event_dir" Then%>
            <!--#include file = "../../includes/event_dir_menu.asp" -->
        <%Else%>
            <!--#include file = "../../includes/admin_menu.asp" -->
        <%End If%>
		
        <div class="col-sm-10">
            <div style="text-align: right;">
                <a href="featured_events.asp">Featured Events</a>
            </div>

		    <h3 class="h3">GSE"Featured Event" Request Form</h3>

            <%If Not sErrMsg = vbNullString Then%>
                <p class="bg-danger"><%=sErrMsg%></p>
            <%End If%>

			<form class="form" name="request_feature" method="post" action="new_event.asp">
            <div class="form-group">
                <label for="gse_event">Which event would you like to feature?</label>
                <select class="form-control" name="gse_event" id="gse_event">
                    <option value="">&nbsp;</option>
                    <%For i = 0 To UBound(Events, 2)%>
                        <option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
                    <%Next%>
                </select>
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
	    </div>
    </div>
</div>
	<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
