<%@ Language=VBScript %>
<%
Option Explicit

Dim sql, conn, rs

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	sql = "INSERT INTO SiteAccess (Status, DateSet) VALUES ('" & Request.Form.Item("status") & "', '" & Now() & "')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

sql = "SELECT Status FROM SiteAccess ORDER BY SiteAccessID DESC"
Set rs = conn.Execute(sql)
Session("site_access_status") = rs(0).Value
Set rs = Nothing

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE Site Access</title>
<!--#include file = "../includes/meta2.asp" -->



</head>
<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 style="margin-left:10px;">CCMeet Site Access</h4>
			
			<form name="set_access" method="Post" action="site_access.asp">
			<table class="display">
				<tr>
					<td class="sub_head">
						Access to the CCMeet portion of GSE is:
						<%Select Case Session("site_access_status")%>
							<%Case "locked"%>
								<input type="radio" name="status" id="status" value="open">Open
								<input type="radio" name="status" id="status" value="locked" checked>Locked
							<%Case Else%>
								<input type="radio" name="status" id="status" value="open" checked>Open
								<input type="radio" name="status" id="status" value="locked">Locked
						<%End Select%>
					</td>
				</tr>
				<tr>
					<td style="text-align:center;">
						<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
						<input type="submit" name="submit" id="submit" value="Change Access">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
</body>
</html>
