<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lThisMeet
Dim i
Dim sMeetName, dMeetDate, sMeetSite, sMeetHost, sWebSite, sComments, sEntryFee, sSport
Dim iMonth, iYear, iDay
Dim MeetArr()
Dim sErrMsg

Dim sMapLink, sMeetInfoSheet, sCourseMap
Dim dWhenShutdown

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim MeetArr(1, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDirID = " & Session("my_id") & " ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetArr(0, i) = rs(0).Value
	MeetArr(1, i) = rs(1).Value & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve MeetArr(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If UBound(MeetArr, 2) = 1 Then lThisMeet = MeetArr(0, 0)

If Request.Form.Item("submit_meet") = "submit_meet" Then 
    lThisMeet = Request.Form.Item("meets")
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	sMeetName = Replace(Request.Form.Item("meet_name"), "'", "''")
	dMeetDate = Request.Form.Item("month") & "/" & Request.Form.Item("day") & "/" & Request.Form.Item("year") 
	sMeetSite = Replace(Request.Form.Item("meet_site"), "'", "''")
	sMeetHost = Replace(Request.Form.Item("meet_host"), "'", "''")
	sSport = Request.Form.Item("sport")
	sWebSite = Request.Form.Item("web_site")
	sComments = Replace(Request.Form.Item("comments"), "'", "''")
	sEntryFee = Request.Form.Item("entry_fee")
	
	If IsDate(dMeetDate) = True Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT MeetName, MeetDate, MeetSite, MeetHost, WebSite, Comments, EntryFee, Sport FROM Meets "
		sql = sql & "WHERE MeetsID = " & lThisMeet
		rs.Open sql, conn, 1, 2
		rs(0).Value = sMeetName
		rs(1).Value = dMeetDate
		rs(2).Value = sMeetSite
		rs(3).Value = sMeetHost
		rs(4).Value = sWebSite
		rs(5).Value = sComments
		rs(6).Value = sEntryFee
		rs(7).Value = sSport
		rs.Update
		rs.Close
		Set rs = Nothing
	Else
		sErrMsg = "This is not a valid date.  Please select a valid date and then re-submit these changes."
	End If
End If

If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

If Not CLng(lThisMeet) = 0 Then
    sql = "SELECT MeetName, MeetDate, MeetSite, MeetHost, WebSite, Comments, EntryFee, Sport, WhenShutdown FROM Meets WHERE MeetsID = " & lThisMeet
    Set rs = conn.Execute(sql)
    sMeetName = Replace(rs(0).Value, "''", "'")
    dMeetDate = rs(1).Value
    If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
    If Not rs(3).Value & "" = "" Then sMeetHost = Replace(rs(3).Value, "''", "'")
    sWebSite = rs(4).Value
    If Not rs(5).Value & "" = "" Then sComments = Replace(rs(5).Value, "''", "'")
    sEntryFee = rs(6).Value
    sSport = rs(7).Value
    dWhenShutdown = rs(8).Value
    Set rs = Nothing

    iMonth = Month(CDate(dMeetDate))
    iDay = Day(CDate(dMeetDate))
    iYear = Year(CDate(dMeetDate))
	
	'get maplink
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMapLink = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT InfoSheet FROM MeetInfo WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMeetInfoSheet = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get course map
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Map FROM CourseMap WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sCourseMap = rs(0).Value
	rs.Close
	Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>GSE  Edit CC Meet Information</title>
<!--#include file = "../../../../includes/js.asp" -->

<script>
function chkFields(){
	if (document.edit_meet.meet_name.value==''){
		alert('You must supply a meet name!');
		return false;
	}
	else 
		if (document.edit_meet.meet_host.value==''){
			alert('You must supply a meet host!');
			return false;
		}
	else
		if (document.edit_meet.meet_site.value==''){
			alert('You must supply a meet site!');
			return false;
		}
	else
		return true;
}
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
	<!--#include file = "../../../../includes/meet_dir_menu.asp" -->

	<h4 class="h4">Cross-Country/Nordic Ski Meet Director: Edit Meet</h4>

	<form class="form-inline" name="get_meets" method="post" action="edit_meet.asp?meet_id=<%=lThisMeet%>">
	<label for="meets">Select Meet:</label>
	<select class="form-control" name="meets" id="meets" onchange="this.form.submit1.click();">
		<option value="">&nbsp;</option>
		<%For i = 0 to UBound(MeetArr, 2) - 1%>
			<%If CLng(lThisMeet) = CLng(MeetArr(0, i)) Then%>
				<option value="<%=MeetArr(0, i)%>" selected><%=MeetArr(1, i)%></option>
			<%Else%>
				<option value="<%=MeetArr(0, i)%>"><%=MeetArr(1, i)%></option>
			<%End If%>
		<%Next%>
	</select>
	<input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
	<input type="submit" class="form-control" name="submit1" id="submit1" value="Get This">
	</form>
			
    <%If Not CLng(lThisMeet) = 0 Then%>
		<!--#include file = "../../meet_dir_nav.asp" -->
				
		<h4 style="background: none;border: none;">Edit Meet Information for <%=sMeetName%> on <%=dMeetDate%></h4>

		<%If Not sErrMsg = vbNullString Then%>
			<p><%=sErrMsg%></p>
		<%End If%>
			
		<form class="form" name="edit_meet" method="post" action="edit_meet.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkFields()">
		<table class="table">
			<tr>
				<th>Meet Name:</th>
				<td><input type="text" class="form-control" name="meet_name" id="meet_name" maxlength="25" value="<%=sMeetName%>"></td>
			</tr>
			<tr>
				<th>Meet Date:</th>
				<td>
					<select class="form-control" name="month" id="month" tabindex="2">
						<%For i = 1 to 12%>
							<%If CInt(iMonth) = CInt(i) Then%>
								<option value="<%=i%>" selected><%=i%></option>
							<%Else%>
								<option value="<%=i%>"><%=i%></option>
							<%End If%>
						<%Next%>
					</select> / <select class="form-control" name="day" id="day" tabindex="3">
						<%For i = 1 to 31%>
							<%If CInt(iDay) = CInt(i) Then%>
								<option value="<%=i%>" selected><%=i%></option>
							<%Else%>
								<option value="<%=i%>"><%=i%></option>
							<%End If%>
						<%Next%>
					</select> / <select class="form-control" name="year" id="year" tabindex="4">
						<%For i = 2005 to Year(Date) + 1%>
							<%If CInt(iYear) = CInt(i) Then%>
								<option value="<%=i%>" selected><%=i%></option>
							<%Else%>
								<option value="<%=i%>"><%=i%></option>
							<%End If%>
						<%Next%>
					</select>
				</td>
			</tr>
			<tr>
				<th>Meet Host:</th>
				<td><input type="text" class="form-control" name="meet_host" id="meet_host" maxlength="50" value="<%=sMeetHost%>"></td>
			</tr>
			<tr>
				<th>Web Site:</th>
				<td><input type="text" class="form-control" name="web_site" id="web_site" maxlength="50" value ="<%=sWebSite%>" onkeyup="chkStr(this)"></td>
			</tr>
			<tr>
				<th>Entry Fee:</th>
				<td><input type="text" class="form-control" name="entry_fee" id="entry_fee" maxlength="15" value="<%=sEntryFee%>" onkeyup="chkStr(this)"></td>
			</tr>
			<tr>
				<th>Sport:</th>
				<td>
					<select class="form-control" name="sport" id="sport">
						<%If sSport="Cross-Country" Then%>
							<option value="Cross-Country" selected>Cross-Country</option>
							<option value="Nordic Ski">Nordic Ski</option>
						<%Else%>
							<option value="Cross-Country">Cross-Country</option>
							<option value="Nordic Ski" selected>Nordic Ski</option>
						<%End If%>
					</select>
				</td>
			</tr>
			<tr>
				<th valign="top">Meet Site:</th>
				<td><textarea class="form-control" name="meet_site" id="meet_site" rows="2"><%=sMeetSite%></textarea></td>
			</tr>
			<tr>
				<th valign="top">Comments:</th>
				<td><textarea class="form-control" name="comments" id="comments" rows="2"><%=sComments%></textarea></td>
			</tr>
			<tr>
				<td style="text-align:center;" colspan="2">
					<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
					<input type="submit" class="form-control" name="submit" id="submit" tabindex="10" value="Save Changes">
				</td>
			</tr>
		</table>
		</form>
    <%End If%>
</div>
<%
conn.Close
Set Conn = Nothing
%>
</body>
</html>
