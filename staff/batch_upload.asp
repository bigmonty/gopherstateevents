<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim Events, Races()

If Not Session("role") = "staff" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")

If Request.Form.Item("submit_event") = "submit_event" Then 
	lEventID = Request.Form.Item("events") 
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

Response.Buffer = true		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate > '" & Date - 7 & "' ORDER By EventDate"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

'get race ids for this event
i = 0
ReDim Races(1, 0)
If Not CLng(lEventID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    Races(0, i) = rs(0).Value
	    Races(1, i) = Replace(rs(1).Value, "''", "'")
	    i = i + 1
	    ReDim Preserve Races(1, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Batch Upload</title>
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
  	<div id="row">
		<!--#include file = "staff_menu.asp" -->
		<div class="col-md-10">
	        <h4 class="h4">Participant Batch Upload Format</h4>

			    <form class="form-inline" name="which_event" method="post" action="batch_upload.asp?event_id=<%=lEventID%>">
			    <label for="events">Select Event:</label>
			    <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				    <option value="">&nbsp;</option>
				    <%For i = 0 to UBound(Events, 2)%>
					    <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						    <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (On <%=Events(2, i)%>)</option>
					    <%Else%>
						    <option value="<%=Events(0, i)%>"><%=Events(1, i)%> (On <%=Events(2, i)%>)</option>
					    <%End If%>
				    <%Next%>
			    </select>
			    <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			    <input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event" style="font-size:0.8em;">
			    </form>
	
	        <p>In order for your data to upload seamlessly the following guidelines must be followed EXACTLY!</p>
	
            <div class="col-md-8">
	            <ul class="list-group">
		            <li class="list-group-item">The file MUST BE a tab-delimited text file (*.txt)</li>
		            <li class="list-group-item">The file CAN NOT have a header row.</li>
		            <li class="list-group-item">The file CAN NOT have any trailing spaces or rows after the final line.</li>
		            <li class="list-group-item">The file MUST HAVE ONLY the following fields IN THIS ORDER.  Required fields are so noted and optional fields MUST exist at 
			            least as an empty field.</li>
		            <li class="list-group-item">
			            <ul>
				            <li>First Name (reqd)</li>
				            <li>Last Name (reqd)</li>
				            <li>Gender ("M" or "F") (reqd)</li>
                            <li>Age (DOB or Age reqd)</li>
				            <li>DOB (DOB or Age reqd)</li>
				            <li>Phone</li>
				            <li>City</li>
				            <li>State</li>
				            <li>Email</li>
				            <li>Shirt Size</li>
				            <li>Bib (reqd)</li>
				            <li>Race ID # (reqd-see below)</li>
			            </ul>
		            </li>
	            </ul>
            </div>
            <div class="col-md-4">
	            <h5 class="h5">Race ID #s</h5>
	            <ul class="list-group">
		            <%For i = 0 To UBound(Races, 2)- 1%>
			            <li class="list-group-item list-group-item-success"><%=Races(1, i)%> - <%=Races(0, i)%></li>
		            <%Next%>
	            </ul>
            </div>
        </div>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>