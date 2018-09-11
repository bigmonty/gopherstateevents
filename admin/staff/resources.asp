<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sVideoName, sVideoLink, sDescription
Dim StaffVids(), Delete()

If Not (Session("role") = "staff" OR Session("role") = "admin") Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_new") = "submit_new" Then
    sVideoName = Replace(Request.Form.Item("video_name"), "'", "''")
    sVideoLink = Request.Form.Item("video_link")
    If Not Request.Form.Item("description") & "" = "" Then sDescription = Replace(Request.Form.Item("description"), "'", "''")

    sql = "INSERT INTO StaffVids (VideoName, VideoLink, Description) VALUES ('" & sVideoName & "', '" & sVideoLink & "', '" & sDescription & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT StaffVidsID, VideoName, VideoLink, Description FROM StaffVids"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
            rs(1).Value = Replace(Request.Form.Item("video_name_" & rs(0).Value), "'", "''")
            rs(2).Value = Request.Form.Item("video_link_" & rs(0).Value)
            If Not Request.Form.Item("description_" & rs(0).Value) & "" = "" Then 
                rs(3).Value = Replace(Request.Form.Item("description_" & rs(0).Value), "'", "''")
            End If
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM StaffVids WHERE StaffVidsID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
End If
	
i = 0
ReDim StaffVids(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT StaffVidsID, VideoName, VideoLink, Description FROM StaffVids ORDER BY StaffVidsID DESC"
rs.Open sql, conn, 1, 2
Do While NOt rs.EOF
	StaffVids(0, i) = rs(0).Value
	StaffVids(1, i) = Replace(rs(1).Value, "''", "'")
    StaffVids(2, i) = rs(2).Value
    If Not rs(3).Value & "" = "" Then StaffVids(3, i) = Replace(rs(3).Value, "''", "'")
    i = i + 1
    ReDim Preserve StaffVids(3, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Staff Resources</title>
<meta name="description" content="Gopher State Events race videos.">

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
		<%If Session("role") = "admin" Then%>
            <!--#include file = "../../includes/admin_menu.asp" -->
        <%Else%>
            <!--#include file = "../../staff/staff_menu.asp" -->
        <%End If%>
		<div class="col-md-10">
			<h4 class="h4">Staff Resources</h4>

            <%If Session("role") = "admin" Then%>
                <form name="new_video" method="post" action="resources.asp">
                <div style="background-color: #ececd8;">
		            <h4 class="h4">Add Help Video</h4>

                   <table>
                        <tr>
                            <th valign="top">Name:</th>
                            <td valign="top"><input type="text" name="video_name" id="video_name"></td>
                            <th valign="top">Link:</th>
                            <td><textarea name="video_link" id="video_link" rows="3" cols="40" style="font-size: 1.2em;"></textarea></td>
                            <th valign="top">Description:</th>
                            <td><textarea name="description" id="description" rows="3" cols="30" style="font-size: 1.2em;"></textarea></td>
                        </tr>
                        <tr>
                            <td style="text-align: center;" colspan="6" valign="top">
                                <input type="hidden" name="submit_new" id="submit_new" value="submit_new">
                                <input type="submit" name="submit1" id="submit1" value="Submit Video">
                            </td>
                        </tr>
                    </table>
                </div>
                </form>

                <form name="edit_videos" method="post" action="resources.asp">
                <div>
		            <h4 class="h4">Existing Videos</h4>
                    <table>
                         <tr>
                            <td style="text-align: center;" colspan="2" valign="top">
                                <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                                <input type="submit" name="submit2" id="submit2" value="Save Changes">
                            </td>
                        </tr>
                        <%For i = 0 To UBound(StaffVids, 2) - 1%>
                           <tr>
                               <td>
                                    <input type="text" name="video_name_<%=StaffVids(0, i)%>" id="video_name_<%=StaffVids(0, i)%>" 
                                    value="<%=StaffVids(1, i)%>" size="40">
                               </td>
                               <td rowspan="4"><%=StaffVids(2, i)%></td>
                            </tr>
                            <tr>
                                </td>
                                    <input type="checkbox" name="delete_<%=StaffVids(0, i)%>" id="delete_<%=StaffVids(0, i)%>">&nbsp;Delete Video
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <textarea name="video_link_<%=StaffVids(0, i)%>" id="video_link_<%=StaffVids(0, i)%>" rows="3" cols="40" 
                                        style="font-size: 1.2em;"><%=StaffVids(2, i)%></textarea>
                                </td>
                             </tr>
                            <tr>
                               <td>
                                    <textarea name="description_<%=StaffVids(0, i)%>" id="description_<%=StaffVids(0, i)%>" rows="3" cols="40" 
                                        style="font-size: 1.2em;"><%=StaffVids(3, i)%></textarea>
                                </td>
                            </tr>
                        <%Next%>
                    </table>
                </div>
                </form>
            <%Else%>
                 <div>
		            <h4 class="h4">Existing Videos</h4>
                    <table>
                        <%For i = 0 To UBound(StaffVids, 2) - 1%>
                            <tr>
                                <th valign="top"><%=StaffVids(1, i)%></th>
                                <td rowspan="2"><%=StaffVids(2, i)%></td>
                            </tr>
                            <tr>
                                <td><%=StaffVids(3, i)%></td>
                            </tr>
                        <%Next%>
                    </table>
                </div>
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