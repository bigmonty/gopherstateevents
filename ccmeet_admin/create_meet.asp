<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i
Dim lThisMeet
Dim MeetDir()
Dim sComments, sWhenShutdown

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	If Not Request.Form.Item("comments") & "" = "" Then sComments = Replace(Request.Form.Item("comments"), "'", "''")
	
    If Request.form.Item("when_shutdown") & "" = "" Then
        sWhenShutdown = Date - 1 & " 6:00:00 PM"
    Else
        sWhenShutdown = Request.Form.Item("when_shutdown")
    End If

	sql = "INSERT INTO Meets (MeetName, MeetDate, MeetHost, WebSite, MeetSite, Comments, Sport, ShowOnline, WhenShutdown, MeetDirID, DateReg) VALUES ('"
	sql = sql & Replace(Request.Form.Item("meet_name"), "'", "''") & "', '" & Request.Form.Item("meet_date") & "', '" 
	sql = sql & Replace(Request.Form.Item("meet_host"), "'", "''") & "', '" & Request.Form.Item("website") & "', '"
	sql = sql & Replace(Request.Form.Item("meet_site"), "'", "''") & "', '" & sComments & "', '"
	sql = sql & Request.Form.Item("sport") & "', '" & Request.Form.Item("show_online") & "', '" & sWhenShutdown & "', "
	sql = sql & Request.Form.Item("meet_dir_id") & ", '" & Date & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
	
	'get meet id
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MeetsID FROM Meets WHERE MeetName = '" & Replace(Request.Form.Item("meet_name"), "'", "''") & "' AND MeetDate = '"
	sql = sql & Request.Form.Item("meet_date") & "' ORDER BY MeetsID DESC"
    rs.Open sql, conn, 1, 2
	lThisMeet = rs(0).Value
    rs.Close
    Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "INSERT INTO MapLinks (MeetsID, MapLink) VALUES (" & lThisMeet & ", '" & Request.form.Item("map_link") & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "INSERT INTO BibRange (MeetsID, BibStart, BibEnd) VALUES (" & lThisMeet & ", 0, 0)"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If

i = 0
ReDim MeetDir(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetDirID, FirstName, LastName FROM MeetDir ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	MeetDir(0, i) = rs(0).Value
	MeetDir(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve MeetDir(1, i)
	rs.MoveNext
Loop
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE Create Cross-country/Nordic Ski Meet</title>

<script>
function chkFlds() {
 	if (document.new_meet.meet_name.value == '' || 
		document.new_meet.meet_host.value == '' || 
		document.new_meet.meet_date.value == '' || 
		document.new_meet.meet_dir_id.value == '' || 
		document.new_meet.sport.value == '' || 
	 	document.new_meet.meet_site.value == '')
		{
  		alert('Please fill in all required fields!');
  		return false
  		}
	else
   		return true
}
</script>
</head>
<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		
		<div class="col-md-10">
			<h4 class="h4">Create Cross-Country/Nordic Ski Meet</h4>
			
			<form name="new_meet" method="post" action="create_meet.asp" onsubmit="return chkFlds()">
			<table>
				<tr>	
					<th><span style="color:#d62002">*</span>Meet Name:</th>
					<td><input name="meet_name" id="meet_name" maxlength="25" size="45"></td>
					<th><span style="color:#d62002">*</span>Meet Date:</th>
					<td><input name="meet_date" id="meet_date" maxLength="10" size="10"></td>
				</tr>
				<tr>	
					<th><span style="color:#d62002">*</span>Meet Host:</th>
					<td><input type="text" name="meet_host" id="meet_host" size="45"></td>
					<th>Web Site:</th>
					<td><input type="text" name="website" id="website" size="45"></td>
				</tr>
				<tr>	
					<th valign="top" rowspan="2"><span style="color:#d62002">*</span>Meet Site:</th>
					<td rowspan="2"><textarea name="meet_site" id="meet_site" rows="2" cols="32"></textarea></td>
					<th valign="top">Map Link:</th>
					<td valign="top"><input type="text" name="map_link" id="map_link" size="45"></td>
				</tr>
				<tr>	
					<th valign="top"><span style="color:#d62002">*</span>Meet Director:</th>
					<td valign="top">
						<select name="meet_dir_id" id="meet_dir_id">
							<option value="">&nbsp;</option>
							<%For i = 0 To UBound(MeetDir, 2) - 1%>
								<option value="<%=MeetDir(0, i)%>"><%=MeetDir(1, i)%></option>
							<%Next%>
						</select>
					</td>
				</tr>
				<tr>	
					<th><span style="color:#d62002">*</span>Sport:</th>
					<td>
						<select name="sport" id="sport">
							<option value="">&nbsp;</option>
							<option value="Cross-Country">Cross-Country</option>
							<option value="Nordic Ski">Nordic Ski</option>
						</select>
					</td>
					<th>Show Online:</th>
					<td>
						<select name="show_online" id="show_online">
							<option value="n">n</option>
							<option value="y">y</option>
						</select>
					</td>
				</tr>
				<tr>	
					<th><span style="color:#d62002">*</span>Shutdown:</th>
					<td><input type="text" name="when_shutdown" id="when_shutdown" size="45"></td>
					<th valign="top">Comments:</th>
					<td><textarea name="comments" id="comments" rows="2" cols="32"></textarea></td>
				</tr>
				<tr>
					<td colspan="4">
						<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
						<input type="submit" name="submit1" id="submit1" value="Create Meet">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
