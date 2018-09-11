<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i
Dim sShowWhat
Dim Logins, Roles(1, 4)
Dim dBegDate, dEndDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

sShowWhat = Request.QueryString("show_what")
If sShowWhat = vbNullString Then sShowWhat = "Event Directors"

Roles(0, 0) = "0"
Roles(1, 0) = "Admin"
Roles(0, 1) = "1"
Roles(1, 1) = "Event Directors"
Roles(0, 2) = "2"
Roles(1, 2) = "Staff"
Roles(0, 3) = "3"
Roles(1, 3) = "Meet Directors"
Roles(0, 4) = "4"
Roles(1, 4) = "Coaches"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_this") = "submit_this" Then
	dBegDate = Request.Form.Item("beg_month") & "/" & Request.Form.Item("beg_day") & "/" & Request.Form.Item("beg_year")
	dEndDate = Request.Form.Item("end_month") & "/" & Request.Form.Item("end_day") & "/" & Request.Form.Item("end_year")
    dEndDate = dEndDate & " 11:59:59 PM"
    sShowWhat = Request.Form.Item("show_what")
End If

If CStr(dBegDate) = vbNullString Then dBegDate = Date - 30
If CStr(dEndDate) = vbNullString Then dEndDate = Now()

Set rs = Server.CreateObject("ADODB.Recordset")
Select Case sShowWhat
    Case "Staff"
        sql = "SELECT StaffLoginID, StaffID, WhenVisit, IPAddress, Browser FROM StaffLogin WHERE WhenVisit >= '" & dBegDate & "' AND WhenVisit <= '" & dEndDate
        sql = sql & "' ORDER BY WhenVisit DESC"
        rs.Open sql, conn, 1, 2
    Case "Event Directors"
        sql = "SELECT EventDirLoginID, EventDirID, WhenVisit, IPAddress, Browser FROM EventDirLogin WHERE WhenVisit >= '" & dBegDate & "' AND WhenVisit <= '" & dEndDate
        sql = sql & "' ORDER BY WhenVisit DESC"    
        rs.Open sql, conn, 1, 2
    Case "Coaches"
        sql = "SELECT CoachLoginID, CoachesID, WhenVisit, IPAddress, Browser FROM CoachLogin WHERE WhenVisit >= '" & dBegDate & "' AND WhenVisit <= '" & dEndDate
        sql = sql & "' ORDER BY WhenVisit DESC"
        rs.Open sql, conn2, 1, 2
    Case "Admin"
        sql = "SELECT AdminLoginID, AdminName, WhenVisit, IPAddress, Browser FROM AdminLogin WHERE WhenVisit >= '" & dBegDate & "' AND WhenVisit <= '" & dEndDate
        sql = sql & "' ORDER BY WhenVisit DESC"
        rs.Open sql, conn, 1, 2
    CAse "Meet Directors"
        sql = "SELECT MeetDirLoginID, MeetDirID, WhenVisit, IPAddress, Browser FROM MeetDirLogin WHERE WhenVisit >= '" & dBegDate & "' AND WhenVisit <= '" & dEndDate
        sql = sql & "' ORDER BY WhenVisit DESC"
        rs.Open sql, conn2, 1, 2
End Select
If rs.RecordCount > 0 Then
	Logins = rs.GetRows()
Else
    ReDim Logins(4, 0)
End If
rs.Close
Set rs = Nothing

If UBound(Logins, 2) > 0 Then
    For i = 0 To UBound(Logins, 2)
        Select Case sShowWhat
            Case "Staff"
                Logins(1, i) = StaffName(Logins(1, i))
            Case "Event Directors"
                Logins(1, i) = EventDirName(Logins(1, i))
            Case "Coaches"
                Logins(1, i) = CoachName(Logins(1, i))
            Case "Meet Directors"
                Logins(1, i) = MeetDirName(Logins(1, i))
        End Select
    Next
End If

Private Function StaffName(lStaffID)
    Set rs2 = SErver.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName FROM Staff WHERE StaffID = " & lStaffID
    rs2.Open sql2, conn, 1, 2
    StaffName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function EventDirName(lEventDirID)
    Set rs2 = SErver.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName FROM EventDir WHERE EventDirID = " & lEventDirID
    rs2.Open sql2, conn, 1, 2
    EventDirName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function MeetDirName(lMeetDirID)
    Set rs2 = SErver.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName FROM MeetDir WHERE MeetDirID = " & lMeetDirID
    rs2.Open sql2, conn2, 1, 2
    MeetDirName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function CoachName(lCoachID)
    Set rs2 = SErver.CreateObject("ADODB.Recordset")
    sql2 = "SELECT FirstName, LastName FROM Coaches WHERE CoachesID = " & lCoachID
    rs2.Open sql2, conn2, 1, 2
    CoachName = Replace(rs2(1).Value, "''", "'") & ", " & Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE&copy; Logins</title>

<!--#include file = "../includes/js.asp" -->

<style type="text/css">
    td,th{padding-right: 5px;}
</style>
</head>


<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE Logins</h4>

			<div style="font-weight:bold;font-size: 0.9em;padding:5px;">
				<form name="get_log" method="post" action="logins.asp">
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

                <span style="font-weight: bold;">Show What:</span>
				<select name="show_what" id="show_what" onchange="this.form.submit1.click()">
					<%For i = 0 To UBound(Roles, 2)%>
						<%If CStr(Roles(1, i)) = sShowWhat Then%>
							<option value="<%=Roles(1, i)%>" selected><%=Roles(1, i)%></option>
						<%Else%>
							<option value="<%=Roles(1, i)%>"><%=Roles(1, i)%></option>
						<%End If%>
					<%Next%>
				</select>

				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input type="submit" name="submit1" id="submit1" value="Set Date Range">
				</form>
			</div>

            <h4 class="h4"><%=sShowWhat%> Logins</h4>

            <table>
                <tr>
                    <th>No.</th>
                    <th>Name</th>
                    <th>When Visit</th>
                    <th>IP Address</th>
                    <th>Browser</th>
                </tr>
                <%For i = 0 To UBound(Logins, 2)%>
                    <%If i mod 2 = 0 Then%>
                        <tr>
                            <td class="alt"><%=i + 1%>)</td>
                            <td class="alt" style="white-space: nowrap;"><%=Logins(1, i)%></td>
                            <td class="alt" style="white-space: nowrap;"><%=Logins(2, i)%></td>
                            <td class="alt" style="white-space: nowrap;"><%=Logins(3, i)%></td>
                            <td class="alt"><%=Logins(4, i)%></td>
                        </tr>
                    <%Else%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td style="white-space: nowrap;"><%=Logins(1, i)%></td>
                            <td style="white-space: nowrap;"><%=Logins(2, i)%></td>
                            <td style="white-space: nowrap;"><%=Logins(3, i)%></td>
                            <td><%=Logins(4, i)%></td>
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

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
