<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lFeaturedEventID, lGSEEventID
Dim i
Dim sEventName, sLocation, sWebURL, sEventDir, sPhone, sDescr, sEmail, sBannerImage, sBlockImage, sActive
Dim iViews, iClicks
Dim Events
Dim dEventDate

If Not Session("role") = "admin" Then 
    If Not Session("role") = "event_dir" Then Response.Redirect "http://www.google.com"
End If

lFeaturedEventID = Request.QueryString("featured_event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
    sEventName = Request.Form.Item("event_name")
    dEventDate = Request.Form.Item("event_date")
    sLocation = Request.Form.Item("location")
    sWebURL = Request.Form.Item("web_url")
    sEventDir = Request.Form.Item("event_dir")
    sPhone = Request.Form.Item("phone")
    sEmail = Request.Form.Item("email")
    sDescr =Request.Form.Item("descr")
    lGSEEventID = Request.Form.Item("gse_event")
    sActive =Request.Form.Item("active")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventName, EventDate, Location, WebURL, EventDir, Phone, Email, Descr, EventID, Active FROM FeaturedEvents "
    sql = sql & "WHERE FeaturedEventsID = " & lFeaturedEventID
    rs.Open sql, conn, 1, 2
    If sEventName & "" = "" Then
        rs(0).Value = rs(0).OriginalValue
    Else
        rs(0).Value = Replace(sEventName, "'", "''")
    End If

    If dEventDate & "" = "" Then
        rs(1).Value = rs(1).OriginalValue
    Else
        If IsDate(dEventDate) Then
            rs(1).Value = dEventDate
        Else
            rs(1).Value = rs(1).OriginalValue
        End If
    End If

    If sLocation & "" = "" Then
        rs(2).Value = rs(2).OriginalValue
    Else
        rs(2).Value = Replace(sLocation, "'", "''")
    End If

    If sWebURL & "" = "" Then
        rs(3).Value = rs(3).OriginalValue
    Else
        rs(3).Value = sWebURL
    End If

    If sEventDir & "" = "" Then
        rs(4).Value = rs(4).OriginalValue
    Else
        rs(4).Value = Replace(sEventDir, "'", "''")
    End If    

    If sPhone & "" = "" Then
        rs(5).Value = rs(5).OriginalValue
    Else
        rs(5).Value = sPhone
    End If

    If sEmail & "" = "" Then
        rs(6).Value = rs(6).OriginalValue
    Else
        rs(6).Value = sEmail
    End If    

    If sDescr & "" = "" Then
        rs(7).Value = rs(7).OriginalValue
    Else
        rs(7).Value = Replace(sDescr, "'", "''")
    End If

    rs(8).Value = lGSEEventID
    rs(9).Value = sActive

    rs.Update
    rs.Close
    Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate, Location, Views, Clicks, WebURL, EventDir, Phone, Email, Descr, EventID, BannerImage, BlockImage, Active "
sql = sql & "FROM FeaturedEvents WHERE FeaturedEventsID = " & lFeaturedEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sLocation = Replace(rs(2).Value, "''", "'")
iViews = rs(3).Value
iClicks = rs(4).Value
sWebURL = rs(5).Value
sEventDir = Replace(rs(6).Value, "''", "'")
sPhone = rs(7).Value
sEmail = rs(8).Value
sDescr = Replace(rs(9).Value, "''", "'")
lGSEEventID = rs(10).Value
sBannerImage = rs(11).Value
sBlockImage = rs(12).Value
sActive = rs(13).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '5/7/2016' ORDER By EventDate"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    Events = rs.GetRows()
Else
    ReDim Events(2, 0)
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Featured Events: Edit Featured Event</title>
<meta name="description" content="Gopher State Events featured events admin utility.">

<script>
$(function() {
    $( "#event_date" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body>
<div class="container">
    <h3 class="h3">GSE Featured Event Control Panel</h3>

	<div class ="row">
         <div class="col-sm-3">
             <h4 class="h4">Banner Image</h4>

	         <a href="javascript:pop('upload_banner.asp?featured_event_id=<%=lFeaturedEventID%>',800,500)">Upload New</a>

             <small>(Preferred Size: 780px x 90px)</small>
        </div>
       <div class="col-sm-9">
            <img src="/featured_events/images/<%=sBannerImage%>" alt="Banner Image" class="img-responsive" style="height:75px;">
        </div>
    </div>
    <hr>
    <div class ="row">
        <div class="col-sm-3">
             <h4 class="h4">Block Image</h4>
             <a href="javascript:pop('upload_block.asp?featured_event_id=<%=lFeaturedEventID%>',800,500)">Upload New</a>
            <small>(Preferred Width: 150px)</small>
        </div>
        <div class="col-sm-3">
            <img src="/featured_events/images/<%=sBlockImage%>" alt="Block Image" class="img-responsive" style="height:75px;">
        </div>
        <div class="col-sm-3">
             <ul>
                 <li>Views: <%=iViews%></li>
                 <li>Clicks: <%=iClicks%></li>
            </ul>
        </div>
        <div class="col-sm-3">
             <a href="click_log.asp?featured_event_id=<%=lFeaturedEventID%>">View Click Log</a>
            |
            <a href="edit_event.asp?featured_event_id=<%=lFeaturedEventID%>">Refresh</a>
        </div>
    </div>

    <div class="row">
        <p class="bg-warning text-warning">
            NOTE:  All data is required.  Any fields left blank will revert to their original value.  If an 
            invalid date is entered it will revert to it's original value.
        </p>

        <form class="form-horizontal" name="edit_event" method="post" action="edit_event.asp?featured_event_id=<%=lFeaturedEventID%>">
        <div class="form-group row">
            <label for="event_name" class="control-label col-sm-2">Event Name:</label>
            <div class="col-sm-4">
                <input class="form-control" type="text" name="event_name" id="event_name" value="<%=sEventName%>">
            </div>
            <label for="event_date" class="control-label col-sm-2">Event Date:</label>
            <div class="col-sm-4">
                <input class="form-control" type="text" name="event_date" id="event_date" value="<%=dEventDate%>">
            </div>
        </div>
        <div class="form-group row">
            <label for="gse_event" class="control-label col-sm-2">GSE Event:</label>
            <div class="col-sm-4">
                <select class="form-control" name="gse_event" id="gse_event">
                    <option value="0">None</option>
                    <%For i = 0 To UBound(Events, 2)%>
                        <%If CLng(Events(0, i)) = CLng(lGSEEventID) Then%>
                            <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
                        <%Else%>
                            <option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
                        <%End If%>
                    <%Next%>
                </select>
            </div>
            <label for="location" class="control-label col-sm-2">Location:</label>
            <div class="col-sm-4">
                <input class="form-control" type="text" name="location" id="location" value="<%=sLocation%>">
            </div>
        </div>
        <div class="form-group row">
            <label for="web_url" class="control-label col-sm-2">Web URL:</label>
            <div class="col-sm-4">
                <input class="form-control" type="text" name="web_url" id="web_url" value="<%=sWebURL%>">
            </div>
            <label for="event_dir" class="control-label col-sm-2">Director:</label>
            <div class="col-sm-4">
                <input class="form-control" type="text" name="event_dir" id="event_dir" value="<%=sEventDir%>">
            </div>
        </div>
        <div class="form-group row">
            <label for="phone" class="control-label col-sm-2">Phone:</label>
            <div class="col-sm-4">
                <input class="form-control" type="text" name="phone" id="phone" value="<%=sPhone%>">
            </div>
            <label for="email" class="control-label col-sm-2">Email:</label>
            <div class="col-sm-4">
                <input class="form-control" type="text" name="email" id="email" value="<%=sEmail%>">
            </div>
        </div>
        <div class="form-group row">
            <label for="active" class="control-label col-sm-2">Active:</label>
            <div class="col-sm-10">
                <select class="form-control" name="active" id="active">
                    <%If sActive = "y" Then%>
                        <option value="y" selected>Yes</option>
                        <option value="n">No</option>
                    <%Else%>
                        <option value="y">Yes</option>
                        <option value="n" selected>No</option>
                    <%End If%>
                </select>
            </div>
        </div>
        <div class="form-group row">
            <label for="descr" class="control-label col-sm-2">Descripton:</label>
            <div class="col-sm-10">
                <textarea class="form-control" name="descr" id="descr" rows="8"><%=sDescr%></textarea>
            </div>
        </div>
        <div class="form-group">
            <input class="form-control" type="hidden" name="submit_this" id="submit_this" value="submit_this">
            <input class="form-control" type="submit" name="submit1" id="submit1" value="Save Changes">
        </div>
        </form>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
