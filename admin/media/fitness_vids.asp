<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID
Dim i
Dim sEventName, sVideoName, sEmbedLink, sVideoLink
Dim RaceVids(), Delete()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim Events
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.Form.Item("submit_new") = "submit_new" Then
    sVideoName = Replace(Request.Form.Item("video_name"), "'", "''")
    sVideoLink = Request.Form.Item("video_link")
    sEmbedLink = Request.Form.Item("embed_link")

    sql = "INSERT INTO RaceVids (EventID, VideoName, EmbedLink, VideoLink) VALUES (" & lEventID & ", '" & sVideoName & "', '" & sEmbedLink
    sql = sql & "', '" & sVideoLink & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceVidsID, VideoName, VideoLink, EmbedLink FROM RaceVids WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
            rs(1).Value = Replace(Request.Form.Item("video_name_" & rs(0).Value), "'", "''")
            rs(2).Value = Request.Form.Item("video_link_" & rs(0).Value)
            rs(3).Value = Request.Form.Item("embed_link_" & rs(0).Value)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM RaceVids WHERE RaceVidsID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

If Not CLng(lEventID) = 0 Then
    'get meet information
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventName FROM Events WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    sEventName = Replace(rs(0).Value, "''", "'")
    rs.Close
    Set rs = Nothing
	
    i = 0
    ReDim RaceVids(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceVidsID, VideoName, EmbedLink, VideoLink FROM RaceVids WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    Do While NOt rs.EOF
	    RaceVids(0, i) = rs(0).Value
	    RaceVids(1, i) = Replace(rs(1).Value, "''", "'")
        RaceVids(2, i) = rs(2).Value
        RaceVids(3, i) = rs(3).Value
        i = i + 1
        ReDim Preserve RaceVids(3, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Edit Race Videos For <%=sEventName%></title>
<!--#include file = "../../includes/js.asp" -->

<script>
function chkFlds() {
 	if (document.new_video.video_name.value == '' || 
	 	document.new_video.video_link.value == '')
		{
  		alert('Please supply a name and a link!');
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

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <h3 class="h3">Add/Edit Videos for <%=sEventName%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="fitness_vids.asp?event_id=<%=lEventID%>">
			<label for="events">Events:</label>
			<select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Events, 2)%>
					<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%Else%>
						<option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			<input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
			</form>
			<br>

			<%If Not Clng(lEventID) = 0 Then%>
				<!--#include file = "../../includes/event_nav.asp" -->
                <!--#include file = "media_nav.asp" -->

                <h4 class="h4">Add Video</h4>
                <form class="form-inline bg-success" name="new_video" method="post" action="fitness_vids.asp?event_id=<%=lEventID%>">
                <label for="video_name">Name:</label>
                <input type="text" class="form-control" name="video_name" id="video_name">
                <label for="embed_link">Embed:</label>
                <textarea class="form-control" name="embed_link" id="embed_link" rows="3"></textarea>
                <label for="video_link">Link:</label>
                <input type="text" class="form-control" name="video_link" id="video_link">
                <input type="hidden" name="submit_new" id="submit_new" value="submit_new">
                <input type="submit" class="form-control" name="submit1" id="submit1" value="Submit Video">
                </form>

                <form name="edit_videos" method="post" action="fitness_vids.asp?event_id=<%=lEventID%>">
                <div>
		            <h4 class="h4">Existing Videos</h4>
                    <table>
                         <tr>
                            <td style="text-align: center;" colspan="3" valign="top">
                                <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                                <input type="submit" name="submit2" id="submit2" value="Save Changes">
                            </td>
                        </tr>
                       <tr>
                            <th>Name</th>
                            <th>Embed Link</th>
                            <th>Video Link</th>
                        </tr>
                        <%For i = 0 To UBound(RaceVids, 2) - 1%>
                            <tr>
                                <td valign="top">
                                    <input type="text" name="video_name_<%=RaceVids(0, i)%>" id="video_name_<%=RaceVids(0, i)%>" value="<%=RaceVids(1, i)%>">
                                    <br>
                                    <input type="checkbox" name="delete_<%=RaceVids(0, i)%>" id="delete_<%=RaceVids(0, i)%>">&nbsp;Delete Video
                                </td>
                                <td>
                                    <textarea name="embed_link_<%=RaceVids(0, i)%>" id="embed_link_<%=RaceVids(0, i)%>" rows="3" cols="50" 
                                        style="font-size: 1.2em;"><%=RaceVids(2, i)%></textarea>
                                </td>
                                <td valign="top">
                                    <input type="text" name="video_link_<%=RaceVids(0, i)%>" id="video_link_<%=RaceVids(0, i)%>" value="<%=RaceVids(3, i)%>"
                                    size="30">
                                </td>
                            </tr>
                        <%Next%>
                    </table>
                </div>
                </form>
            <%End If%>
       </div>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>