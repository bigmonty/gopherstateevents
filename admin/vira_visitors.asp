<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim Visitors()
Dim i, j
Dim lEventDirID
Dim dBegDate, dEndDate, sPage

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	dBegDate = Request.Form.Item("beg_month") & "/" & Request.Form.Item("beg_day") & "/" & Request.Form.Item("beg_year")
	dEndDate = Request.Form.Item("end_month") & "/" & Request.Form.Item("end_day") & "/" & Request.Form.Item("end_year")
	sPage = Request.Form.Item("page")
End If

If CStr(dBegDate) = vbNullString Then dBegDate = Date - 1
If CStr(dEndDate) = vbNullString Then dEndDate = Now()
If CStr(sPage) = vbNullString Then sPage = "all"

ReDim Visitors(4, 0)
i = 0

sql = "SELECT When_Visit, Page, IPAddress, Browser, EventDirID FROM Visitors WHERE IPAddress <> '64.8.133.230' AND "
sql = sql & "(When_Visit >= '" & CDate(dBegDate) & "' AND When_Visit <= '" & CDate(dEndDate) + 1 & "') ORDER BY When_Visit DESC"

If sPage <> "all" Then
	sql = sql & " AND Page = '" & sPage & "'"
End If

Set rs=conn.Execute(sql)
Do While Not rs.EOF
	For j = 0 to 4
		Visitors(j, i) = rs(j).Value
	Next
	
	i = i + 1
	ReDim Preserve Visitors(4, i)
	rs.MoveNext
Loop
Set rs=Nothing

Private Function GetEventDir(lEventDirID)
	If lEventDirID <> "0" Then
		sql = "SELECT FirstName, LastName FROM EventDir WHERE EventDirID = " & lEventDirID
		Set rs = conn.Execute(sql)
		GetEventDir = rs(0).Value & " " & rs(1).Value
		Set rs = Nothing
	Else
		GetEventDir = "None"
	End If
End Function		
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE&copy; Visitors</title>

<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		
		<div class="col-md-10">
			<h4 class="h4">GSE&copy; Visitors Log</h4>

			<div style="font-weight:bold;">
				<form name="get_log" method="post" action="vira_visitors.asp">
				<span style="font-weight:normal">From</span>&nbsp;
				<select name="beg_month" id="beg_month">
					<%For i = 1 To 12%>
						<%If Month(CDate(dBegDate)) = i Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
				/
				<select name="beg_day" id="beg_day">
					<%For i = 1 To 31%>
						<%If Day(CDate(dBegDate)) = i Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
				/
				<select name="beg_year" id="beg_year">
					<%For i = 2005 To Year(Date)%>
						<%If Year(CDate(dBegDate)) = i Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
	
				<span style="font-weight:normal;">To</span>
				
				<select name="end_month" id="end_month">
					<%For i = 1 To 12%>
						<%If Month(CDate(dEndDate)) = i Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
				/
				<select name="end_day" id="end_day">
					<%For i = 1 To 31%>
						<%If Day(CDate(dEndDate)) = i Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
				/
				<select name="end_year" id="end_year">
					<%For i = 2005 To Year(Date)%>
						<%If Year(CDate(dEndDate)) = i Then%>
							<option value="<%=i%>" selected><%=i%></option>
						<%Else%>
							<option value="<%=i%>"><%=i%></option>
						<%End If%>
					<%Next%>
				</select>
				&nbsp;&nbsp;&nbsp;
				<select name="page" id="page">
					<option value="all">All</option>
					<option value="about_bob">About Bob</option>
					<option value="about_vira">About GSE</option>
					<option value="admin">Admin</option>
					<option value="contact_us">Contact Us</option>
					<option value="default">Default</option>
					<option value="event_dir_reg">New Event Dir</option>
					<option value="events">Events</option>
					<option value="honor_roll">Honor Roll</option>
					<option value="part_data">Part Data</option>
					<option value="perf_center">Performance Center</option>
					<option value="perf_list">Performance List</option>
					<option value="pricing">Pricing</option>
					<option value="race_dir">Race Dir</option>
					<option value="records">Records</option>
					<option value="results">Results</option>
					<option value="services">Services</option>
				</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input type="submit" name="submit" id="submit" value="Get Log">
				</form>
			</div>

			<h5 style="margin-left:10px;text-align:right;color:#555;background-color:#ececd8;">Num Visitors:&nbsp;<%=UBound(Visitors, 2)%></h5>
			
			<table>	
				<tr>
					<th style="width:15px;">No.</th>
					<th>When</th>
					<th>Page</th>
					<th>IP Add</th>
					<th style="white-space:nowrap;">Browser Information</th>
					<th style="white-space:nowrap;">Event Director</th>
				</tr>
				<%For i = 0 to UBound(Visitors, 2) - 1%>
					<%If i mod 2 = 0 Then%>
						<tr>
							<td class="alt" style="text-align:right;width:15px"><%=i + 1%></td>
							<td class="alt" style="white-space:nowrap;">
								<%=Visitors(0, i)%>
							</td>
							<td class="alt" style="white-space:nowrap;">
								<%=Visitors(1, i)%>
							</td>
							<td class="alt" style="white-space:nowrap;">
								<%=Visitors(2, i)%>
							</td>
							<td class="alt" style="width:300px;white-space:nowrap;">
								<%=Left(Visitors(3, i), 55)%>
							</td>
							<td class="alt" style="width:125px;white-space:nowrap;">
								<%=GetEventDir(Visitors(4, i))%>
							</td>
						</tr>
					<%Else%>
						<tr>
							<td style="text-align:right;width:15px"><%=i + 1%></td>
							<td style="white-space:nowrap;">
								<%=Visitors(0, i)%>
							</td>
							<td style="white-space:nowrap;">
								<%=Visitors(1, i)%>
							</td>
							<td style="white-space:nowrap;">
								<%=Visitors(2, i)%>
							</td>
							<td style="width:300px;white-space:nowrap;">
								<%=Left(Visitors(3, i), 55)%>
							</td>
							<td style="width:125px;white-space:nowrap;">
								<%=GetEventDir(Visitors(4, i))%>
							</td>
						</tr>
					<%End If%>
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
