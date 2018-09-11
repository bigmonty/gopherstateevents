<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lMeetDirID
Dim iNumRsrvd, iNumPndng, iYear
Dim sStatus, sShowOnline, sTimingMethod
Dim sngInvoiceTtl, sngInvoice
Dim sErrMsg, sSport
Dim  Meets(), MeetDir()
Dim fs, fname, sFileName

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
		
iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")
	
sSport = Request.QueryString("sport")
If sSport = vbNullString Then sSport = "Cross-Country"
'If sSport = vbNullString Then sSport = "Nordic Ski"
						
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

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
rs.Close
Set rs = Nothing

If Request.form.Item("submit_changes") = "submit_changes" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetsID, Status, ShowOnline, TimingMethod, MeetDirID FROM Meets "
    sql = sql & "WHERE (MeetDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear & "') AND Sport = '" & sSport & "' ORDER BY MeetDate"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
	    sStatus = Request.Form.Item("status_" & rs(0).Value)
	    sShowOnline = Left(Request.Form.Item("show_online_" & rs(0).Value), 1)
	    sTimingMethod = Request.Form.Item("timing_method_" & rs(0).Value)
	    lMeetDirID = Request.Form.Item("meet_dir_id_" & rs(0).Value)

		rs(1).Value = sStatus
		rs(2).Value = sShowOnline
		rs(3).Value = sTimingMethod
		rs(4).Value = lMeetDirID
  		rs.Update
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
End If

i = 0
iNumPndng = 0
iNumRsrvd = 0
sngInvoiceTtl = 0
ReDim Meets(11, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
If sSport = "Both" Then
    sql = "SELECT MeetsID, MeetName, MeetDate, MeetSite, Status, ShowOnline, TimingMethod, MeetDirID, Invoice, Sport FROM Meets "
    sql = sql & "WHERE MeetDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear & "' ORDER BY MeetDate"
Else
    sql = "SELECT MeetsID, MeetName, MeetDate, MeetSite, Status, ShowOnline, TimingMethod, MeetDirID, Invoice, Sport FROM Meets "
    sql = sql & "WHERE MeetDate BETWEEN '1/1/" & iYear & "' AND '12/31/" & iYear & "' AND Sport = '" & sSport & "' ORDER BY MeetDate"
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	For j = 0 To 9
		If Not rs(j).Value & "" = "" Then Meets(j, i) = Replace(rs(j).Value, "''", "'")
	Next

    If ChkCntrct(rs(0).Value, rs(2).Value) = True Then 
        Meets(10, i) = "View"
        Meets(11, i) = "/contracts/" & Year(rs(2).Value) & "/cross-country/" & rs(0).Value & ".pdf"
    End If

	i = i + 1
	ReDim Preserve Meets(11, i)

    If CDate(rs(2).Value) <= Date Then sngInvoiceTtl = CSng(sngInvoiceTtl) + CSng(rs(8).Value)

    If rs(4).Value = "reserved" Then
        iNumRsrvd = CInt(iNumRsrvd) + 1
    Else
        iNumPndng = CInt(iNumPndng) + 1
    End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Function ChkCntrct(lThisMeet, dMeetDate)
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
    sFileName = "C:\Inetpub\h51web\gopherstateevents\contracts\" & Year(dMeetDate) & "\cross-country\" & lThisMeet & ".pdf"
    ChkCntrct = fs.FileExists(sFileName)
    Set fs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Meets</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">GSE Cross-Country/Nordic Ski Meets</h4>

            <div class="col-md-8">
                <ul class="list-inline">
                    <li class="list-inline-item">Total Meets: <%=UBound(Meets, 2)%></li>
                    <li class="list-inline-item">Pending: <%=iNumPndng%></li>
                    <li class="list-inline-item">Reserved: <%=iNumRsrvd%></li>
                </ul>
            </div>
            <div class="col-md-4">
                <ul class="list-inline">
                    <li class="list-inline-item"><a href="meets.asp?sport=Cross-Country&amp;year=<%=iYear%>">Cross-Country</a></li>
                    <li class="list-inline-item"><a href="meets.asp?sport=Nordic Ski&amp;year=<%=iYear%>">Nordic Ski</a></li>
                    <li class="list-inline-item"><a href="meets.asp?sport=Both&amp;year=<%=iYear%>">Both</a></li>
                </ul>
            </div>

            <ul class="list-inline">
                <%For i = 2002 To Year(Date) + 1%>
                    <li class="list-inline-item"><a href="meets.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
            </ul>

            <%If Not sErrMsg = vbNullString Then%>
                <p><%=sErrMsg%></p>
            <%End If%>

			<form name="meet_status" method="Post" action="meets.asp?year=<%=iYear%>">
            <div class="table-responsive">
			    <table class="table table-striped">
			        <tr>
					    <td style="text-align:center;" colspan="9">
						    <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
						    <input type="submit" name="submit1" id="submit1" value="Submit Changes">
					    </td>
				    </tr>
				    <tr>
					    <th>No.</th>
					    <th>Event</th>
					    <th>Date</th>
                        <th>Sport</th>
					    <th>Status</th>
					    <th>Vsble</th>
                        <th>Event Dir</th>
                        <th>Contract</th>
				    </tr>
				    <%For i = 0 To UBound(Meets, 2) - 1%>
						<tr>
							<td><%=i + 1%>)</td>
							<td><a href="/ccmeet_admin/manage_meet/manage_meet.asp?meet_id=<%=Meets(0, i)%>"><%=Meets(1, i)%></a></td>
							<td><%=Meets(2, i)%></td>
							<td><%=Meets(9, i)%></td>
							<td>
                                <%If Meets(4, i) = "pending" Then%>
								    <select name="status_<%=Meets(0, i)%>" id="status_<%=Meets(0, i)%>" style="background-color: yellow">
									    <%If Meets(4, i) = "pending" Then%>
										    <option value="pending" selected>pending</option>
										    <option value="reserved">rsvrd</option>
									    <%Else%>
										    <option value="pending">pending</option>
										    <option value="reserved" selected>rsvrd</option>
									    <%End If%>
								    </select>
                                <%Else%>
								    <select name="status_<%=Meets(0, i)%>" id="status_<%=Meets(0, i)%>">
									    <%If Meets(4, i) = "pending" Then%>
										    <option value="pending" selected>pending</option>
										    <option value="reserved">rsvrd</option>
									    <%Else%>
										    <option value="pending">pending</option>
										    <option value="reserved" selected>rsvrd</option>
									    <%End If%>
								    </select>
                                <%End If%>
							</td>
							<td>
                                <%If Meets(5, i) = "n" Then%>
								    <select name="show_online_<%=Meets(0, i)%>" id="show_online_<%=Meets(0, i)%>" style="background-color: yellow;">
   								        <%If Meets(5, i) = "n" Then%>
										    <option value="n" selected>n</option>
										    <option value="y">y</option>
								        <%Else%>
										    <option value="n">n</option>
										    <option value="y" selected>y</option>
								        <%End If%>
							        </select>
                                <%Else%>
								    <select name="show_online_<%=Meets(0, i)%>" id="show_online_<%=Meets(0, i)%>">
   								        <%If Meets(5, i) = "n" Then%>
										    <option value="n" selected>n</option>
										    <option value="y">y</option>
								        <%Else%>
										    <option value="n">n</option>
										    <option value="y" selected>y</option>
								        <%End If%>
							        </select>
                                <%End If%>
							</td>
							<td>
								<select name="meet_dir_id_<%=Meets(0, i)%>" id="meet_dir_id_<%=Meets(0, i)%>">
									<%For j = 0 To UBound(MeetDir, 2) - 1%>
                                        <%If CLng(Meets(7, i)) = CLng(MeetDir(0, j)) Then%>
                                            <option value="<%=MeetDir(0, j)%>" selected><%=MeetDir(1, j)%></option>
                                        <%Else%>
                                            <option value="<%=MeetDir(0, j)%>"><%=MeetDir(1, j)%></option>
                                        <%End If%>
                                    <%Next%>
								</select>
							</td>
                            <td>
                                <%If Meets(10, i) ="View" Then%>
                                    <a href="javascript:pop('<%=Meets(11, i)%>',800,600)"><%=Meets(10, i)%></a>
                                <%Else%>
                                    &nbsp;
                                <%End If%>
                            </td>
						</tr>
				    <%Next%>
			    </table>
            </div>
			</form>
		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>