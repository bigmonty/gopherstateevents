<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim RacePix(), Delete()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceGalleryID, GalleryName, GalleryLink, EmbedLink FROM RaceGallery ORDER BY EventID DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
            rs(1).Value = Replace(Request.Form.Item("gallery_name_" & rs(0).Value), "'", "''")
            rs(2).Value = Request.Form.Item("Gallery_link_" & rs(0).Value)
            rs(3).Value = Request.Form.Item("embed_link_" & rs(0).Value)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM RaceGallery WHERE RaceGalleryID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
End If
	
i = 0
ReDim RaceGallery(4, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT rv.RaceGalleryID, rv.GalleryName, rv.EmbedLink, rv.GalleryLink, e.EventName FROM RaceGallery rv INNER JOIN Events e ON rv.EventID = e.EventID "
sql = sql & "ORDER BY EventDate DESC"
rs.Open sql, conn, 1, 2
Do While NOt rs.EOF
	RaceGallery(0, i) = rs(0).Value
	RaceGallery(1, i) = Replace(rs(1).Value, "''", "'")
    RaceGallery(2, i) = rs(2).Value
    RaceGallery(3, i) = rs(3).Value
    RaceGallery(4, i) = Replace(rs(4).Value, "''", "'")
    i = i + 1
    ReDim Preserve RaceGallery(4, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Fitness Event Galleries</title>
<meta name="link" content="Gopher State Events race galleries.">

<!--#include file = "../../includes/js.asp" -->

</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">Edit GSE Galleries</h4h3>
				
             <form name="edit_galleries" method="post" action="all_pix.asp">
            <div>
		        <h4 class="h4">Existing Galleries</h4>
                <table>
                     <tr>
                        <td style="text-align: center;" colspan="4" valign="top">
                            <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                            <input type="submit" name="submit2" id="submit2" value="Save Changes">
                        </td>
                    </tr>
                   <tr>
                        <th>Event</th>
                        <th>Gallery</th>
                        <th>Embed Link</th>
                        <th>Gallery Link</th>
                    </tr>
                    <%For i = 0 To UBound(RaceGallery, 2) - 1%>
                        <tr>
                            <td valign="top"><%=RaceGallery(4, i)%></td>
                            <td valign="top">
                                <input type="text" name="gallery_name_<%=RaceGallery(0, i)%>" id="gallery_name_<%=RaceGallery(0, i)%>" value="<%=RaceGallery(1, i)%>">
                                <br>
                                <input type="checkbox" name="delete_<%=RaceGallery(0, i)%>" id="delete_<%=RaceGallery(0, i)%>">&nbsp;Delete Gallery
                            </td>
                            <td>
                                <textarea name="embed_link_<%=RaceGallery(0, i)%>" id="embed_link_<%=RaceGallery(0, i)%>" rows="3" cols="50" 
                                    style="font-size: 1.2em;"><%=RaceGallery(2, i)%></textarea>
                            </td>
                            <td valign="top">
                                <input type="text" name="gallery_link_<%=RaceGallery(0, i)%>" id="gallery_link_<%=RaceGallery(0, i)%>" value="<%=RaceGallery(3, i)%>"
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