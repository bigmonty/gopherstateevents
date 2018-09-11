<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim EventDir()
Dim i, j
Dim sViewWho

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

sViewWho = "active"		'default criteria
If Request.Form.Item("submit_this") = ("submit_this") Then
	sViewWho = Request.Form.Item("view_who")
End If

i = 0
ReDim EventDir(8, 0)
If sViewWho = "active" Then
	sql = "SELECT EventDirID, FirstName, LastName, City, State, Zip, Phone, UserID, Password, Email FROM EventDir "
	sql = sql & "WHERE Active = 'y' ORDER BY LastName, FirstName"
Else
	sql = "SELECT EventDirID, FirstName, LastName, City, State, Zip, Phone, UserID, Password, Email FROM EventDir "
	sql = sql & "ORDER BY LastName, FirstName"
End If

Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventDir(0, i) = rs(0).Value
	EventDir(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then EventDir(2, i) = Replace(rs(3).Value, "''", "'")
	EventDir(3, i) = rs(4).Value
	EventDir(4, i) = rs(5).Value
	EventDir(5, i) = rs(6).Value
	EventDir(6, i) = rs(7).Value
	EventDir(7, i) = rs(8).Value
	EventDir(8, i) = rs(9).Value
	i = i + 1
	ReDim Preserve EventDir(8, i)
	rs.MoveNext
Loop
Set rs = Nothing

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; View Event Directors</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <!--#include file = "../../includes/event_dir_nav.asp" -->

			<h4 class="h4">Event Director Information</h4>
		
			<form class="form-inline" role="form" name="view_who" method="post" action="event_dirs.asp" style="font-size:0.85em;">
			<label for="view_who">View Who:&nbsp;&nbsp;</label>
			<input type="radio" name="view_who" id="view_who" value="active" 
			<%If sViewWho = "active" Then%>
				checked
			<%End If%> 
			onclick="this.form.submit.click();">Active Event Directors Only
			
			<input type="radio" name="view_who" id="view_who" value="all"				
			<%If sViewWho = "all" Then%>
				checked
			<%End If%> 
			onclick="this.form.submit.click();">All Event Directors
			<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
			<input class="form-control" type="submit" name="submit" id="submit" value="View These">
			</form>
		
			<div class="table-responsive">
				<table class="table table-striped">
					<tr>
						<th>No.</th>
						<th>Name (Email)</th>
						<th>City</th>
						<th>ST</th>
						<th>Zip</th>
						<th>Phone</th>
						<th>User ID</th>
						<th>Password</th>
					</tr>
					<%For i = 0 to UBound(EventDir, 2) - 1%>
						<tr>
							<td style="text-align:right;">
								<%=i + 1%>)
							</td>
							<%For j = 1 to 7%>
								<td style="white-space:nowrap;">
									<%If j = 1 Then%>
										<a href="mailto:<%=EventDir(8, i)%>"><%=EventDir(1, i)%></a>
									<%Else%>
										<%=EventDir(j, i)%>
									<%End If%>
								</td>
							<%Next%>
						</tr>
					<%Next%>
				</table>
			</div>
		</div>
	</div>
<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
