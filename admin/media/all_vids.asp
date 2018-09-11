<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim RaceVids(), Delete()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceVidsID, VideoName, VideoLink, EmbedLink FROM RaceVids ORDER BY EventID DESC"
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
	
i = 0
ReDim RaceVids(4, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT rv.RaceVidsID, rv.VideoName, rv.EmbedLink, rv.VideoLink, e.EventName FROM RaceVids rv INNER JOIN Events e ON rv.EventID = e.EventID "
sql = sql & "ORDER BY EventDate DESC"
rs.Open sql, conn, 1, 2
Do While NOt rs.EOF
	RaceVids(0, i) = rs(0).Value
	RaceVids(1, i) = Replace(rs(1).Value, "''", "'")
    RaceVids(2, i) = rs(2).Value
    RaceVids(3, i) = rs(3).Value
    RaceVids(4, i) = Replace(rs(4).Value, "''", "'")
    i = i + 1
    ReDim Preserve RaceVids(4, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Fitness Event Videos</title>
<meta name="link" content="Gopher State Events race videos.">

<!--#include file = "../../includes/js.asp" -->

</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">All Videos</h4>
				
             <form name="edit_videos" method="post" action="all_vids.asp">
            <div>
		        <h4 class="h4">Existing Videos</h4>
                <table>
                     <tr>
                        <td style="text-align: center;" colspan="4" valign="top">
                            <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                            <input type="submit" name="submit2" id="submit2" value="Save Changes">
                        </td>
                    </tr>
                   <tr>
                        <th>Event</th>
                        <th>Video</th>
                        <th>Embed Link</th>
                        <th>Video Link</th>
                    </tr>
                    <%For i = 0 To UBound(RaceVids, 2) - 1%>
                        <tr>
                            <td valign="top"><%=RaceVids(4, i)%></td>
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
       </div>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>