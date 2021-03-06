<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisMeet
Dim i
Dim sMeetName, sVideoName, sVideoLink, sEmbedLink
Dim RaceVids(), Delete()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_new") = "submit_new" Then
    sVideoName = Replace(Request.Form.Item("video_name"), "'", "''")
    sVideoLink = Request.Form.Item("video_link")
    sEmbedLink = Request.Form.Item("embed_link")

    sql = "INSERT INTO RaceVids (MeetsID, VideoName, VideoLink, EmbedLink) VALUES (" & lThisMeet & ", '" & sVideoName & "', '" & sVideoLink
    sql = sql & "', '" & sEmbedLink & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceVidsID, VideoName, VideoLink, EmbedLink FROM RaceVids WHERE MeetsID = " & lThisMeet
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

'get meet information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lThisMeet
rs.Open sql, conn, 1, 2
sMeetName = Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing
	
i = 0
ReDim RaceVids(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceVidsID, VideoName, EmbedLink, VideoLink FROM RaceVids WHERE MeetsID = " & lThisMeet
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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Edit Race Videos For <%=sMeetName%></title>
<!--#include file = "../../includes/js.asp" -->
<script>
function chkFlds() {
 	if (document.new_video.video_name.value == '' || 
	 	document.new_video.embed_link.value == '')
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
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "manage_meet_nav.asp" -->
			<%End If%>

		    <h4 class="h4"><%=sMeetName%> Videos</h4>

            <form class="form bg-success" name="new_video" method="post" action="cc_vids.asp?meet_id=<%=lThisMeet%>">
		    <h4 class="h4">Add Video</h4>

            <div class="form-group">
                <label for="video_name">Name:</label>
                <input type="text" class="form-control" name="video_name" id="video_name">
            </div>
            <div class="form-group">
                <label for="embed_link">Embed Link:</label>
                <textarea class="form-control" name="embed_link" id="embed_link" rows="3"></textarea>
            </div>
            <div class="form-group">
                <label for="video_link">Video Link:</label>
                <input type="text"class="form-control" name="video_link" id="video_link">
            </div>
            <div class="form-group">
                <input type="hidden" name="submit_new" id="submit_new" value="submit_new">
                <input type="submit" class="form-control" name="submit1" id="submit1" value="Submit Video">
            </div>
            </form>
 
            <form class="form" name="edit_videos" method="post" action="cc_vids.asp?meet_id=<%=lThisMeet%>">
		    <h4 class="h4">Existing Videos</h4>
            <table class="table table-striped">
                <tr>
                    <td style="text-align: center;" colspan="3" valign="top">
                        <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                        <input type="submit" class="form-control" name="submit2" id="submit2" value="Save Changes">
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
                            <textarea class="form-control" name="embed_link_<%=RaceVids(0, i)%>" id="embed_link_<%=RaceVids(0, i)%>" rows="3"><%=RaceVids(2, i)%></textarea>
                        </td>
                        <td valign="top">
                            <input type="text" class="form-control" name="video_link_<%=RaceVids(0, i)%>" id="video_link_<%=RaceVids(0, i)%>"
                            value="<%=RaceVids(3, i)%>">
                        </td>
                    </tr>
                <%Next%>
            </table>
        </div>
        </form>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>