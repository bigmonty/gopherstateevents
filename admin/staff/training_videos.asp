<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim sVideoName, sVideoLink, sDescription, sVideoType, sViewWhich, sVideoTypeStaff, sWasWatched
Dim TrngVids(), Delete(), VideoTypes(4), WhichViews(2)
Dim cdoMessage, cdoConfig

If Not (Session("role") = "staff" OR Session("role") = "admin") Then Response.Redirect "/default.asp?sign_out=y"

sViewWhich = Request.QueryString("view_which")
If sViewWhich = vbNullString Then sViewWhich = "Not Watched"

sVideoTypeStaff = Request.QueryString("video_type_staff")
If sVideoTypeStaff = vbNullString Then sVideoTypeStaff = "All"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
    
VideoTypes(0) = "RaceWare"
VideoTypes(1) = "CCMeet"
VideoTypes(2) = "Combined"
VideoTypes(3) = "Other"
VideoTypes(4) = "All"

WhichViews(0) = "Watched"
WhichViews(1) = "Not Watched"
WhichViews(2) = "All"

i = 0
ReDim TrngVids(5, 0)

If Request.Form.Item("submit_watched") = "submit_watched" Then
    Call GetVideos()

    For i = 0 To UBound(TrngVids, 2) - 1
        sWasWatched = WasWatched(TrngVids(0, i))

        If sWasWatched = "n" Then
            If Request.Form.Item("watched_" & TrngVids(0, i)) = "y" Then
                'if they watched a non-watched video insert it into the watched list
                sql = "INSERT INTO TrngVidsWatched (TrngVidsID, StaffID, WhenWatched) VALUES (" & TrngVids(0, i) & ", " & Session("staff_id")
                sql = sql & ", '" & Date & "')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            End If
        Else
            If Request.Form.Item("watched_" & TrngVids(0, i)) = "n" Then
                'if they didn't watch a video delete it from the watched list
                sql = "DELETE FROM TrngVidsWatched WHERE TrngVidsID = " & TrngVids(0, i) & " AND StaffID = " & Session("staff_id")
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            End If
        End If
    Next
ElseIf Request.Form.Item("submit_staff_vids") = "submit_staff_vids" Then
    sViewWhich = Request.Form.Item("view_which")
    sVideoTypeStaff = Request.Form.Item("video_type_staff")
ElseIf Request.Form.Item("submit_new") = "submit_new" Then
    sVideoName = Replace(Request.Form.Item("video_name"), "'", "''")
    sVideoLink = Request.Form.Item("video_link")
    If Not Request.Form.Item("description") & "" = "" Then sDescription = Replace(Request.Form.Item("description"), "'", "''")
    sVideoType = Request.Form.Item("video_type")

    sql = "INSERT INTO TrngVids (VideoName, VideoLink, Description, VideoType) VALUES ('" & sVideoName & "', '" & sVideoLink & "', '" & sDescription
    sql = sql & "', '" & sVideoType & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%

	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.com;nateengelg.92@gmail.com;Ashley.pethan@gmail.com;bolstad@q.com;kurthunter57@yahoo.com;nathanhowe18@gmail.com;solveigkc@gmail.com;jacummings2@gmail.com;taren.weyer@gmail.com"
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "New Training Video Online"
		.TextBody = "A new Gopher State Events training video is online.  This one is entitled " & sVideoName
		.Send
	End With
	Set cdoMessage = Nothing
    Set cdoConfig = Nothing
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TrngVidsID, VideoName, VideoLink, Description, VideoType FROM TrngVids"
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
            rs(4).Value = Request.Form.Item("video_type_" & rs(0).Value)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM TrngVids WHERE TrngVidsID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
End If
	
Call GetVideos()

Private Sub GetVideos()
    Set rs = Server.CreateObject("ADODB.Recordset")
    Select Case sVideoTypeStaff
        Case "All"
            sql = "SELECT TrngVidsID, VideoName, VideoLink, Description, VideoType FROM TrngVids ORDER BY TrngVidsID"
        Case Else
            sql = "SELECT TrngVidsID, VideoName, VideoLink, Description, VideoType FROM TrngVids WHERE VideoType = '" & sVideoTypeStaff 
            sql = sql & "' ORDER BY TrngVidsID DESC"
    End Select
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        sWasWatched = vbNullString
        If Session("role") = "staff" Then sWasWatched = WasWatched(rs(0).Value)

        If sViewWhich = "Not Watched" Then
            If sWasWatched = "n" Then
	            TrngVids(0, i) = rs(0).Value
	            TrngVids(1, i) = Replace(rs(1).Value, "''", "'")
                TrngVids(2, i) = rs(2).Value
                If Not rs(3).Value & "" = "" Then TrngVids(3, i) = Replace(rs(3).Value, "''", "'")
                TrngVids(4, i) = rs(4).Value
                TrngVids(5, i) = sWasWatched
                i = i + 1
                ReDim Preserve TrngVids(5, i)
            End If
        ElseIf sViewWhich = "Watched" Then
            If sWasWatched = "y" Then
	            TrngVids(0, i) = rs(0).Value
	            TrngVids(1, i) = Replace(rs(1).Value, "''", "'")
                TrngVids(2, i) = rs(2).Value
                If Not rs(3).Value & "" = "" Then TrngVids(3, i) = Replace(rs(3).Value, "''", "'")
                TrngVids(4, i) = rs(4).Value
                TrngVids(5, i) = sWasWatched
                i = i + 1
                ReDim Preserve TrngVids(5, i)
            End If
        Else
	        TrngVids(0, i) = rs(0).Value
	        TrngVids(1, i) = Replace(rs(1).Value, "''", "'")
            TrngVids(2, i) = rs(2).Value
            If Not rs(3).Value & "" = "" Then TrngVids(3, i) = Replace(rs(3).Value, "''", "'")
            TrngVids(4, i) = rs(4).Value
            TrngVids(5, i) = sWasWatched
            i = i + 1
            ReDim Preserve TrngVids(5, i)
        End If
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Function WasWatched(lThisVideo)
    Dim rs2, sql2

    WasWatched = "n"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT TrngVidsID FROM TrngVidsWatched WHERE TrngVidsID = " & lThisVideo & " AND StaffID = " & Session("staff_id")
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then WasWatched = "y"
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Staff Training Videos</title>
<meta name="description" content="Gopher State Events staff training videos.">

<!--#include file = "../../includes/js.asp" -->

<script>
function chkFlds() {
 	if (document.new_video.video_name.value == '' || 
        document.new_video.video_type.value == '' || 
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
			<h2 class="h2">Staff Training Videos</h2>

            <%If Session("role") = "admin" Then%>
                <form name="new_video" method="post" action="training_videos.asp">
                <div style="background-color: #ececd8;">
		            <h4 class="h4">Add Training Video</h4>

                   <table>
                        <tr>
                            <th valign="top">Name:</th>
                            <td valign="top"><input type="text" name="video_name" id="video_name"></td>
                            <th valign="top">Type:</th>
                            <td valign="top">
                                <select name="video_type" id="video_type">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(VideoTypes) - 1   'don't show "all" because that is just a viewing parameter%>
                                        <option value="<%=VideoTypes(i)%>"><%=VideoTypes(i)%></option>
                                    <%Next%>
                                </select>
                            </td>
                            <th valign="top">Link:</th>
                            <td><textarea name="video_link" id="video_link" rows="3" cols="30" style="font-size: 1.2em;"></textarea></td>
                            <th valign="top">Description:</th>
                            <td><textarea name="description" id="description" rows="3" cols="25" style="font-size: 1.2em;"></textarea></td>
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

                <form name="edit_videos" method="post" action="training_videos.asp">
                <div>
		            <h4 class="h4">Existing Videos</h4>
                    <table>
                         <tr>
                            <td style="text-align: center;" colspan="2" valign="top">
                                <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                                <input type="submit" name="submit2" id="submit2" value="Save Changes">
                            </td>
                        </tr>
                        <%For i = 0 To UBound(TrngVids, 2) - 1%>
                           <tr>
                               <td>
                                    <input type="text" name="video_name_<%=TrngVids(0, i)%>" id="video_name_<%=TrngVids(0, i)%>" 
                                    value="<%=TrngVids(1, i)%>" size="40">
                               </td>
                               <td rowspan="4"><%=TrngVids(2, i)%></td>
                            </tr>
                            <tr>
                                </td>
                                    <input type="checkbox" name="delete_<%=TrngVids(0, i)%>" id="delete_<%=TrngVids(0, i)%>">&nbsp;Delete Video
                                </td>
                            </tr>
                            <tr>
                                </td>
                                    <select name="video_type_<%=TrngVids(0, i)%>" id="video_type_<%=TrngVids(0, i)%>">
                                        <%For j = 0 To UBound(VideoTypes)%>
                                            <%If TrainingVids(4, i) = VideoTypes(j) Then%>
                                                <option value="<%=VideoTypes(j)%>" selected><%=VideoTypes(j)%></option>
                                            <%Else%>
                                                <option value="<%=VideoTypes(j)%>"><%=VideoTypes(j)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <textarea name="video_link_<%=TrngVids(0, i)%>" id="video_link_<%=TrngVids(0, i)%>" rows="3" cols="40" 
                                        style="font-size: 1.2em;"><%=TrngVids(2, i)%></textarea>
                                </td>
                             </tr>
                            <tr>
                               <td>
                                    <textarea name="description_<%=TrngVids(0, i)%>" id="description_<%=TrngVids(0, i)%>" rows="3" cols="40" 
                                        style="font-size: 1.2em;"><%=TrngVids(3, i)%></textarea>
                                </td>
                            </tr>
                        <%Next%>
                    </table>
                </div>
                </form>
            <%Else%>
                 <div>
                     <div class="row">
                         <div class="bg-warning">
                            <form role="form" class="form-inline" name="show_what" method="post" action="training_videos.asp">
			                <div class="form_group">
                                <label for="view_which">View Which Videos:</label>
			                    <select class="form-control" name="view_which" id="view_which" onchange="this.form.submit3.click()">
                                    <%For i = 0 To UBound(WhichViews)%>
                                        <%If sViewWhich = WhichViews(i) Then%>
                                            <option value="<%=WhichViews(i)%>" selected><%=WhichViews(i)%></option>
                                        <%Else%>
                                            <option value="<%=WhichViews(i)%>"><%=WhichViews(i)%></option>
                                        <%End If%>
                                    <%Next%>
			                    </select>
                                <label for="video_type_staff">Video Type:</label>
			                    <select class="form-control" name="video_type_staff" id="video_type_staff" onchange="this.form.submit3.click()">
                                    <%For i = 0 To UBound(VideoTypes)%>
                                        <%If sVideoTypeStaff = VideoTypes(i) Then%>
                                            <option value="<%=VideoTypes(i)%>" selected><%=VideoTypes(i)%></option>
                                        <%Else%>
                                            <option value="<%=VideoTypes(i)%>"><%=VideoTypes(i)%></option>
                                        <%End If%>
                                    <%Next%>
			                    </select>
			                    <input class="form-control" type="hidden" name="submit_staff_vids" id="submit_staff_vids" value="submit_staff_vids">
			                    <input class="form-control" type="submit" name="submit3" id="submit3" value="Get These">
			                </div>
                            </form>
                         </div>
                     </div>

                     <h4 class="h4">Existing Videos</h4>

                     <%If UBound(TrngVids, 2) = 0 Then%>
                        <div class="bg-danger">
                            There are no available videos for the parameters you have selected.  If you believe this is in error, please re-set
                            the parameters above.
                        </div>
                     <%Else%>
                         <form role="form" method="post" action="training_videos.asp?view_which=<%=sViewWhich%>&amp;video_type_staff=<%=sVideoTypeStaff%>">
                        <table class="table table-striped">
                            <tr>
                                <th>No.</th>
                                <th>Video</th>
                                <th>Type</th>
                                <th>Thumbnail</th>
                                <th>Description</th>
                                <th>Watched?</th>
                            </tr>
                            <%For i = 0 To UBound(TrngVids, 2) - 1%>
                                <tr>
                                    <th><%=i + 1%></th>
                                    <th><%=TrngVids(1, i)%></th>
                                    <td><%=TrngVids(4, i)%></td>
                                    <td><%=TrngVids(2, i)%></td>
                                    <td><%=TrngVids(3, i)%></td>
                                    <td>
                                        <select class="form-control" name="watched_<%=TrngVids(0, i)%>" id="watched_<%=TrngVids(0, i)%>">
                                            <%If TrngVids(5, i) = "y" Then%>
                                                <option value="y" selected>y</option>
                                                <option value="n">n</option>
                                            <%Else%>
                                                <option value="y">y</option>
                                                <option value="n" selected>n</option>
                                            <%End If%>
                                        </select>
                                    </td>
                                </tr>
                            <%Next%>
                            <tr>
                                <td colspan="6">
			                        <input class="form-control" type="hidden" name="submit_watched" id="submit_watched" value="submit_watched">
			                        <input class="form-control" type="submit" name="submit4" id="submit4" value="Check These">
                                </td>
                            </tr>
                        </table>
                         </form>
                     <%End If%>
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