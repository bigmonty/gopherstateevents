<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lEventID
Dim i, j, k
Dim sEventName, sViewWhat, sGalleryLink, sProofs, sEmbedLink
Dim iFirstPic, iImageID, iMainPic
Dim RacePix(), Events()
Dim dEventDate
Dim bHasPix, bHasProofs

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

iFirstPic = Request.QueryString("first_pic")
If CStr(iFirstPic) = vbNullString Then iFirstPic = 0

sViewWhat = Request.QueryString("view_what")
If sViewWhat = vbNullString Then sViewWhat = "selected"

iImageID = Request.QueryString("image_id")
If CStr(iImageID) = vbNullString Then iImageID = 0
If iImageID = 0 Then iMainPic = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
	
If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

i = 0
ReDim Events(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate, GalleryLink, EmbedLink FROM Events ORDER BY EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If HasPix(rs(0).Value) = True Then
		Events(0, i) = rs(0).Value
		Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
        Events(2, i) = rs(4).Value
		i = i + 1
		ReDim Preserve Events(2, i)
	ElseIf Not rs(3).Value & "" = "" Then
		Events(0, i) = rs(0).Value
		Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
        Events(2, i) = rs(4).Value
		i = i + 1
		ReDim Preserve Events(2, i)
	ElseIf Not rs(4).Value & "" = "" Then
		Events(0, i) = rs(0).Value
		Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
        Events(2, i) = rs(4).Value
		i = i + 1
		ReDim Preserve Events(2, i)
	End if
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not CLng(lEventID) = 0 Then
	bHasPix = False
    bHasProofs = False
	If HasProofs(lEventID) = True Then bHasPix = True
	
	'get event information
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT EventName, EventDate, GalleryLink, Proofs, EmbedLink FROM Events WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
	    sEventName = Replace(rs(0).Value, "''", "'")
	    dEventDate = rs(1).Value
	    sGalleryLink = rs(2).Value
        sProofs = rs(3).Value
        sEmbedLink = rs(4).Value
    End If
	rs.Close
	Set rs = Nothing
	
	i = 0
	ReDim RacePix(2, 0)
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT RacePixID, PixName, Caption FROM RacePix WHERE EventID = " & lEventID
	rs.Open sql, conn, 1, 2
	Do While NOt rs.EOF
		RacePix(0, i) = rs(0).Value
		RacePix(1, i) = "/gallery/" & lEventID & "/" & Replace(rs(1).Value, "''", "'")
		If Not rs(2).Value & "" = "" Then RacePix(2, i) = Replace(rs(2).Value, "''", "'")
		If Not CInt(iImageID) = 0 Then
			If CLng(rs(0).Value) = CLng(iImageID) Then iMainPic = i
		End If
		i = i + 1
		ReDim Preserve RacePix(2, i)
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	
	If CInt(iFirstPic) > UBound(RacePix, 2) - 5 Then iFirstPic = UBound(RacePix, 2) - 5
	If CInt(iFirstPic) < 0 Then iFirstPic = 0
End If

Private Function HasPix(lThisEvent)
	HasPix = False
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT EventID FROM RacePix WHERE EventID = " & lThisEvent
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then HasPix = True
	rs2.Close
	Set rs2 = Nothing
End Function

Private Function HasProofs(lThisEvent)
	HasProofs = False
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>

<!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events Image Gallery</title>
<meta name="description" content="Event Pictures by Gopher State Events, LLC, a conventional timing service in Minnetonka, MN.">
 </head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div class="row">
		<h3 style="text-align:center;color:#fff;background-color:#5679b9;margin-bottom:10px;">GSE Image Gallery</h3>
		
		<form name="which_event" method="post" action="gallery.asp?event_id=<%=lEventID%>">
		<span style="font-weight:bold;">Event:</span>
		<select name="events" id="events" onchange="this.form.get_event.click()">
			<option value="">&nbsp;</option>
			<%For i = 0 to UBound(Events, 2) - 1%>
				<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
					<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
				<%Else%>
					<option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
				<%End If%>
			<%Next%>
		</select>
		<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
		<input type="submit" name="get_event" id="get_event" value="Get This Event" style="font-size:0.8em;">
		</form>

		<%If Not CLng(lEventID) = 0 Then%>
			<div style="font-size:0.8em;background-color:#ececec;text-align:right;margin-bottom:10px;">
				<%If Not sGalleryLink = vbNullString Then%>
					<a href="javascript:pop('<%=sGalleryLink%>',1000,750)" rel="nofollow">External Gallery</a>
					<%If bHasPix = True Then%>
						&nbsp;|&nbsp;
					<%End If%>
				<%End If%>

                <%If Not sEmbedLink = vbNullString Then%>
                   <%=sEmbedLink%>
                <%End If%>

				<%If bHasPix = True Then%>
					<a href="gallery.asp?event_id=<%=lEventID%>&amp;view_what=selected" rel="nofollow">View Selected</a>
					&nbsp;|&nbsp;
					<a href="gallery.asp?event_id=<%=lEventID%>&amp;view_what=thumb" rel="nofollow">View Thumbnails</a>
				<%End If%>
			</div>
			
			<%If bHasPix = True Then%>
				<%If sViewWhat = "selected" Then%>
					<h4 class="h4">Select Image To View</h4>
					
					<div style="width:750px;margin:none;padding:none;float:left;">
						<%If Not RacePix(2, iMainPic) & "" = "" Then%>
							<p style="text-align:left;"><%=RacePix(2, iMainPic)%></p>
						<%End If%>
						<img src="<%=RacePix(1, iMainPic)%>" style="width:750px;">
					</div>
						
					<div style="margin-left:775px;text-align:center;background-color:#5679b9;text-align:center;padding:0 0 0 2px;">
						<table style="background:none;">
							<tr>
								<td style="background-color:#000405;text-align:center;">
									<ul style="background-color:#000405;display:block;list-type-style:none;margin:0 5px 0 8px;padding:0;">
										<li style="display:inline;padding:0;margin:0;">
											<a href="gallery.asp?event_id=<%=lEventID%>&amp;first_pic=0&amp;image_id=<%=iImageID%>" 
												style="width:100%;height:100%;padding:0;margin:0;" rel="nofollow">
												<img src="/graphics/first.jpg" alt="First" style="padding:0;margin:0;">
											</a>
										</li>
										<li style="display:inline;padding:0;margin:0;">
											<a href="gallery.asp?event_id=<%=lEventID%>&amp;first_pic=<%=CInt(iFirstPic) - 1%>&amp;image_id=<%=iImageID%>" 
												style="width:100%;height:100%;padding:0;margin:0;" rel="nofollow">
												<img src="/graphics/prev.jpg" alt="Prev" style="padding:0;margin:0;">
											</a>
										</li>
										<li style="display:inline;padding:0;margin:0;">
											<a href="gallery.asp?event_id=<%=lEventID%>&amp;first_pic=<%=CInt(iFirstPic) + 1%>&amp;image_id=<%=iImageID%>" 
												style="width:100%;height:100%;padding:0;margin:0;" rel="nofollow">
												<img src="/graphics/next.jpg" alt="Next" style="padding:0;margin:0;">
											</a>
										</li>
										<li style="display:inline;padding:0;margin:0;">
											<a href="gallery.asp?event_id=<%=lEventID%>&amp;first_pic=<%=CInt(UBound(RacePix, 2) + 5)%>&amp;image_id=<%=iImageID%>" 
												style="width:100%;height:100%;padding:0;margin:0;" rel="nofollow">
												<img src="/graphics/last.jpg" alt="Last" style="padding:0;margin:0;">
											</a>
										</li>
									</ul>
								</td>
							</tr>
							<%k = 0%>
							<%For i = 0 To UBound(RacePix, 2) - 1%>
								<%If CInt(i) >= CInt(iFirstPic) Then%>
									<tr>
										<td>
											<a href="gallery.asp?image_id=<%=RacePix(0, i)%>&amp;event_id=<%=lEventID%>&first_pic=<%=iFirstPic%>" rel="nofollow">
												<img src="<%=RacePix(1, i)%>" style="width:150px;">
											</a>
										</td>
									</tr>
									
									<%k = k + 1%>
									<%If k = 3 Then Exit For%>
								<%End If%>
							<%Next%>
						</table>
					</div>
				<%ElseIf sViewWhat = "thumb" Then%>
					<h4 class="h4">Thumbnail View (click to enlarge)</h4>
					
					<table style="background-color:#5679b9;">
						<%For i = 0 To UBound(RacePix, 2) - 1 Step 6%>
							<tr>
								<%For j = 0 To 5%>
									<td valign="top">
										<%If UBound(RacePix, 2) - 1 >= i + j Then%>
											<a href="javascript:pop('this_image.asp?image_id=<%=RacePix(0, i + j)%>&amp;event_id=<%=lEventID%>',800,800)" rel="nofollow">
												<img src="<%=RacePix(1, i + j)%>" style="width:160px;">
											</a>
										<%Else%>
											&nbsp;
										<%End If%>
									</td>
								<%Next%>
							</tr>
						<%Next%>
					</table>
				<%End If%>
			<%End If%>
		</div>
	<%End If%>
	<!--#include file = "../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>