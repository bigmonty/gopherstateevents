<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim Visitors()
Dim i, j
Dim lMeetDirID, lCoachID
Dim dBegDate, dEndDate, sPage

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	dBegDate = Request.Form.Item("beg_month") & "/" & Request.Form.Item("beg_day") & "/" & Request.Form.Item("beg_year")
	dEndDate = Request.Form.Item("end_month") & "/" & Request.Form.Item("end_day") & "/" & Request.Form.Item("end_year")
	sPage = Request.Form.Item("page")
Else
	dBegDate = Date - 7
	dEndDate = Now()
	sPage = "all"
End If

i = 0
ReDim Visitors(5, 0)
sql = "SELECT When_Visit, Page, IPAddress, Browser, MeetDirID, CoachID FROM Visitors "
sql = sql & "WHERE IPAddress <> '64.8.133.230' AND (When_Visit >= '" & CDate(dBegDate) & "' AND When_Visit <= '" 
sql = sql & CDate(dEndDate) + 1 & "') ORDER BY When_Visit DESC"
If sPage <> "all" Then
	sql = sql & " AND Page = '" & sPage & "'"
End If

Set rs=conn.Execute(sql)
Do While Not rs.EOF
	For j = 0 to 5
		Visitors(j, i) = rs(j).Value
	Next
	
	i = i + 1
	ReDim Preserve Visitors(5, i)
	rs.MoveNext
Loop
Set rs=Nothing

Private Function GetMeetDir(lMeetDirID)
	If lMeetDirID <> "0" Then
		sql = "SELECT FirstName, LastName FROM MeetDir WHERE MeetDirID = " & lMeetDirID
		Set rs = conn.Execute(sql)
		GetMeetDir = rs(0).Value & " " & rs(1).Value
		Set rs = Nothing
	Else
		GetMeetDir = "None"
	End If
End Function		

Private Function GetCoach(lCoachID)
	If lCoachID <> "0" Then
		sql = "SELECT FirstName, LastName FROM Coaches WHERE CoachesID = " & lCoachID
		Set rs = conn.Execute(sql)
		GetCoach = rs(0).Value & " " & rs(1).Value
		Set rs = Nothing
	Else
		GetCoach = "None"
	End If
End Function		
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE Admin Visitors Log</title>

<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		
		<div class="col-md-10">
			<h4 class="h4">CCMeet Visitors Log</h4>
			
			<form name="get_log" method="post" action="visitors.asp">
			<div style="font-size:0.9em;">
				Date Range:&nbsp;&nbsp;
				<span style="font-weight:normal">From&nbsp;
				<select name="beg_month" id="beg_month">
					<%
					For i = 1 to 12
						If Month(CDate(dBegDate)) = i Then
							Response.Write("<option value='" & i & "' selected='selected'>" & i & "</option>")
						Else
							Response.Write("<option value='" & i & "'>" & i & "</option>")
						End If
					Next
					%>
				</select>&nbsp;/&nbsp;
				<select name="beg_day" id="beg_day">
					<%
					For i = 1 to 31
						If Day(CDate(dBegDate)) = i Then
							Response.Write("<option value='" & i & "' selected='selected'>" & i & "</option>")
						Else
							Response.Write("<option value='" & i & "'>" & i & "</option>")
						End If
					Next
					%>
				</select>&nbsp;/&nbsp;
				<select name="beg_year" id="beg_year">
					<%
					For i = 2004 to Year(Date)
						If Year(CDate(dBegDate)) = i Then
							Response.Write("<option value='" & i & "' selected='selected'>" & i & "</option>")
						Else
							Response.Write("<option value='" & i & "'>" & i & "</option>")
						End If
					Next
					%>
				</select>&nbsp;To&nbsp;
	
				<select name="end_month" id="end_month">
					<%
					For i = 1 to 12
						If Month(CDate(dEndDate)) = i Then
							Response.Write("<option value='" & i & "' selected='selected'>" & i & "</option>")
						Else
							Response.Write("<option value='" & i & "'>" & i & "</option>")
						End If
					Next
					%>
				</select>&nbsp;/&nbsp;
				<select name="end_day" id="end_day">
					<%
					For i = 1 to 31
						If Day(CDate(dEndDate)) = i Then
							Response.Write("<option value='" & i & "' selected='selected'>" & i & "</option>")
						Else
							Response.Write("<option value='" & i & "'>" & i & "</option>")
						End If
					Next
					%>
				</select>&nbsp;/&nbsp;
				<select name="end_year" id="end_year">
					<%
					For i = 2004 to Year(Date)
						If Year(CDate(dEndDate)) = i Then
							Response.Write("<option value='" & i & "' selected='selected'>" & i & "</option>")
						Else
							Response.Write("<option value='" & i & "'>" & i & "</option>")
						End If
					Next
					%>
				</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<select name="page" id="page">
					<option value="all">All</option>
					<option value="about_bob">About Bob</option>
					<option value="about_vira">About GSE</option>
					<option value="admin">Admin</option>
					<option value="contact_us">Contact Us</option>
					<option value="default">Default</option>
					<option value="event_dir_reg">New Event Dir</option>
					<option value="events">Events</option>
					<option value="part_data">Part Data</option>
					<option value="pricing">Pricing</option>
					<option value="race_dir">Race Dir</option>
					<option value="results">Results</option>
					<option value="services">Services</option>
				</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				</span>
				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input type="submit" name="submit" id="submit" value="Get Log">
			</div>
			</form>
			
			<table>	
				<tr>
					<td style="font-weight:bold;padding-right:0px;width:15px">No.</td>
					<td style="font-weight:bold;width:100px" nowrap="nowrap">Meet Dir</td>
					<td style="font-weight:bold;width:100px" nowrap="nowrap">Coach</td>
					<td style="font-weight:bold;width:130px">When</td>
					<td style="font-weight:bold;width:50px">Page</td>
					<td style="font-weight:bold;width:75px">IP Add</td>
					<td style="font-weight:bold;width:300px" nowrap="nowrap">Browser Information</td>
				</tr>
				<%For i = 0 to UBound(Visitors, 2) - 1%>
					<tr>
						<%If i mod 2 = 0 Then%>
							<td class="alt" style="text-align:center;padding-right:0px;width:15px"><%=i + 1%>)</td>
							<td class="alt" style="width:100px;white-space:nowrap;"><%=GetMeetDir(Visitors(4, i))%></td>
							<td class="alt" style="width:100px;white-space:nowrap;"><%=GetCoach(Visitors(5, i))%></td>
							<td class="alt" style="width:130px;white-space:nowrap;"><%=Visitors(0, i)%></td>
							<td class="alt" style="width:50px;white-space:nowrap;"><%=Visitors(1, i)%></td>
							<td class="alt" style="width:75px;white-space:nowrap;"><%=Visitors(2, i)%></td>
							<td class="alt" style="width:300px;white-space:nowrap;"><%=Left(Visitors(3, i), 55)%></td>
						<%Else%>
							<td style="text-align:center;padding-right:0px;width:15px"><%=i + 1%>)</td>
							<td style="width:100px;white-space:nowrap;"><%=GetMeetDir(Visitors(4, i))%></td>
							<td style="width:100px;white-space:nowrap;"><%=GetCoach(Visitors(5, i))%></td>
							<td style="width:130px;white-space:nowrap;"><%=Visitors(0, i)%></td>
							<td style="width:50px;white-space:nowrap;"><%=Visitors(1, i)%></td>
							<td style="width:75px;white-space:nowrap;"><%=Visitors(2, i)%></td>
							<td style="width:300px;white-space:nowrap;"><%=Left(Visitors(3, i), 55)%></td>
						<%End If%>
					</tr>
				<%Next%>
			</table>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
