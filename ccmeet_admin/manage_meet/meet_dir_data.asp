<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim MeetDirArr(), MeetsArr()
Dim lMeetDirID

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim MeetDirArr(5, 0)
sql = "SELECT MeetDirID,  FirstName, LastName, Email, Phone, UserID, Password FROM MeetDir ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetDirArr(0, i) = rs(0).Value
	MeetDirArr(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	MeetDirArr(2, i) = rs(3).Value
	MeetDirArr(3, i) = rs(4).Value
	MeetDirArr(4, i) = rs(5).Value
	MeetDirArr(5, i) = rs(6).Value
	i = i + 1
	ReDim Preserve MeetDirArr(5, i)
	rs.MoveNext
Loop
Set rs = Nothing

Function GetMeets(lMeetDirID)
	j = 0
	ReDim MeetsArr(1, 0)
	sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE MeetDirID = " & lMeetDirID 
	sql = sql & " ORDER BY MeetDate DESC"
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
		MeetsArr(0, j) = rs(0).Value
		MeetsArr(1, j) = rs(1).Value & " (" & Year(rs(2).Value) & ") " & rs(3).Value
		j = j + 1
		ReDim Preserve MeetsArr(1, j)
		rs.MoveNext
	Loop
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE  Meet Director Data</title>
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
	th{
		text-align:left;
	}
</style>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">CCMeet Meet Dir Data</h4>
			
			<table>
				<tr>
					<td style="text-align:right;" colspan="7">
                        <a href="javascript:pop('meet_dir_dwnld.asp',800,750)">Download</a>
                        |
						<a href="meet_dir_data.asp">Refresh</a>
					</td>
				</tr>
				<tr>
					<th>
						No.
					</th>
					<th>
						Name
					</th>
					<th>
						Email
					</th>
					<th>
						Phone
					</th>
					<th>
						User ID
					</th>
					<th>
						Password
					</th>
					<th>
						Meet(s)
					</th>
				</tr>
				<%For i = 0 to UBound(MeetDirArr, 2) - 1%>
					<%If i mod 2 = 0 Then%>
						<tr>
							<td style="background-color:#d0d0d0;text-align:right;white-space:nowrap;" valign="top">
								<%=i + 1%>)
							</td>
							<td style="background-color:#d0d0d0;white-space:nowrap;" valign="top">
								<a href="javascript:pop('this_meet_dir.asp?meet_dir_id=<%=MeetDirArr(0, i)%>',400,400)"><%=MeetDirArr(1, i)%></a>
							</td>
							<td style="background-color:#d0d0d0;white-space:nowrap;" valign="top">
								<a href="mailto:<%=MeetDirArr(2, i)%>">Send</a>
							</td>
							<td style="background-color:#d0d0d0;white-space:nowrap;" valign="top">
								<%=MeetDirArr(3, i)%>
							</td>
							<td style="background-color:#d0d0d0;white-space:nowrap;" valign="top">
								<%=MeetDirArr(4, i)%>
							</td>
							<td style="background-color:#d0d0d0;white-space:nowrap;" valign="top">
								<%=MeetDirArr(5, i)%>
							</td>
							<td style="background-color:#d0d0d0;white-space:nowrap;">
								<%Call GetMeets(MeetDirArr(0, i))%>
								<%For j = 0 to UBound(MeetsArr, 2) - 1%>
									<%If j <> UBound(MeetsArr, 2) - 1 Then%>
										<a href="javascript:pop('/events/cross_ctry/ccmeet_info.asp?meet_id=<%=MeetsArr(0, j)%>',650,500)"><%=MeetsArr(1, j)%></a><br>
									<%Else%>
										<a href="javascript:pop('/events/cross_ctry/ccmeet_info.asp?meet_id=<%=MeetsArr(0, j)%>',650,500)"><%=MeetsArr(1, j)%></a>
									<%End If%>
								<%Next%>
							</td>
						</tr>
					<%Else%>
						<tr>
							<td style="text-align:right;white-space:nowrap;" valign="top">
								<%=i + 1%>)
							</td>
							<td style="white-space:nowrap;" valign="top">
								<a href="javascript:pop('this_meet_dir.asp?meet_dir_id=<%=MeetDirArr(0, i)%>',400,400)"><%=MeetDirArr(1, i)%></a>
							</td>
							<td style="white-space:nowrap;" valign="top">
								<a href="mailto:<%=MeetDirArr(2, i)%>">Send</a>
							</td>
							<td style="white-space:nowrap;" valign="top">
								<%=MeetDirArr(3, i)%>
							</td>
							<td style="white-space:nowrap;" valign="top">
								<%=MeetDirArr(4, i)%>
							</td>
							<td style="white-space:nowrap;" valign="top">
								<%=MeetDirArr(5, i)%>
							</td>
							<td style="white-space:nowrap;">
								<%Call GetMeets(MeetDirArr(0, i))%>
								<%For j = 0 to UBound(MeetsArr, 2) - 1%>
									<%If j <> UBound(MeetsArr, 2) - 1 Then%>
										<a href="javascript:pop('/events/cross_ctry/ccmeet_info.asp?meet_id=<%=MeetsArr(0, j)%>',650,500)"><%=MeetsArr(1, j)%></a><br>
									<%Else%>
										<a href="javascript:pop('/events/cross_ctry/ccmeet_info.asp?meet_id=<%=MeetsArr(0, j)%>',650,500)"><%=MeetsArr(1, j)%></a>
									<%End If%>
								<%Next%>
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
