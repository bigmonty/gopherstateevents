<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim Staff(), AvailPref(11)
Dim i, j
Dim sSummCmnts, sSchlCmnts

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(3, 0)
sql = "SELECT StaffID, FirstName, LastName, Email, Phone FROM Staff ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then Staff(2, i) = Replace(rs(3).Value, "''", "'")
	If Not rs(4).Value & "" = "" Then Staff(3, i) = Replace(rs(4).Value, "''", "'")
	i = i + 1
	ReDim Preserve Staff(3, i)
	rs.MoveNext
Loop
Set rs = Nothing

Private Sub MyAvail(lStaffID)
    Dim x

    For x = 0 To 11
        AvailPref(x) = vbNullString
    Next

    sSummCmnts = vbNullString
    sSchlCmnts = vbNullString

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SummSatAM, SummSatPM, SummSunAM, SummSunPM, SummWkdyAM, SummWkdyPM, SchlSatAM, SchlSatPM, SchlSunAM, SchlSunPM, SchlWkdyAM, SchlWkdyPM, "
    sql = sql & "SummCmnts, SchlCmnts FROM StaffAvailPref WHERE StaffID = " & lStaffID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        For x = 0 To 11
            AvailPref(x) = rs(0).Value
        Next
        If Not rs(12).Value & "" = "" Then sSummCmnts = Replace(rs(12).Value, "''", "'")
        If Not rs(13).Value & "" = "" Then sSchlCmnts = Replace(rs(13).Value, "''", "'")
    End If
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Staff Availability Preferences</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td,th{padding-right: 5px;}
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">Staff Availability Preferences</h4>
		
		    <table style="font-size:0.85em;margin:10px 0 0 0;">
			    <tr>
				    <th style="text-align:right;" rowspan="2" valign="bottom">No.</th>
				    <th rowspan="2" valign="bottom">Name (Email)</th>
				    <th rowspan="2" valign="bottom">Phone</th>
				    <th style="text-align:center;background-color: #ececec;color: #039;" colspan="6">Summer</th>
				    <th style="text-align:center;background-color: #ececec;color: #093;" colspan="6">School Year</th>
			    </tr>
			    <tr>
				    <th style="text-align: center;color: #039;">Sat AM</th>
				    <th style="text-align: center;color: #039;">Sat PM</th>
				    <th style="text-align: center;color: #039;">Sun AM</th>
				    <th style="text-align: center;color: #039;">Sun PM</th>
				    <th style="text-align: center;color: #039;">Wkdy AM</th>
				    <th style="text-align: center;color: #039;">Wkdy PM</th>
				    <th style="text-align: center;color: #093;">Sat AM</th>
				    <th style="text-align: center;color: #093;">Sat PM</th>
				    <th style="text-align: center;color: #093;">Sun AM</th>
				    <th style="text-align: center;color: #093;">Sun PM</th>
				    <th style="text-align: center;color: #093;">Wkdy AM</th>
				    <th style="text-align: center;color: #093;">Wkdy PM</th>
			    </tr>
			    <%For i = 0 to UBound(Staff, 2) - 1%>
                    <%Call MyAvail(Staff(0, i))%>

					<%If i mod 2 = 0 Then%>
				        <tr>
						    <td class="alt" style="text-align:right"><%=i + 1%>)</td>
							<td class="alt" style="white-space:nowrap;"><a href="mailto:<%=Staff(2, i)%>"><%=Staff(1, i)%></a></td>
                            <td class="alt" style="white-space:nowrap;"><%=Staff(3, i)%></td>
                            <%For j = 0 To 11%>
                                <%If j < 6 Then%>
                                    <td class="alt" style="text-align: center;color: #039;"><%=AvailPref(j)%></td>
                                <%Else%>
                                    <td class="alt" style="text-align: center;color: #093;"><%=AvailPref(j)%></td>
                                <%End If%>
                            <%Next%>
                        </tr>
                        <tr>
                            <th class="alt" style="text-align:right;" colspan="3">Summer Comments:</th>
                            <td class="alt" colspan="12"><%=sSummCmnts%></td>
                        </tr>
                        <tr>
                            <th class="alt" style="text-align:right;" colspan="3">School Year Comments:</th>
                            <td class="alt" colspan="12"><%=sSchlCmnts%></td>
                        </tr>
					<%Else%>
				        <tr>
						    <td style="text-align:right"><%=i + 1%>)</td>
							<td style="white-space:nowrap;"><a href="mailto:<%=Staff(2, i)%>"><%=Staff(1, i)%></a></td>
                            <td style="white-space:nowrap;"><%=Staff(3, i)%></td>
                            <%For j = 0 To 11%>
                                <%If j < 6 Then%>
                                    <td style="text-align: center;color: #039;"><%=AvailPref(j)%></td>
                                <%Else%>
                                    <td style="text-align: center;color: #093;"><%=AvailPref(j)%></td>
                                <%End If%>
                            <%Next%>
                        </tr>
                        <tr>
                            <th style="text-align:right;" colspan="3">Summer Comments:</th>
                            <td colspan="12"><%=sSummCmnts%></td>
                        </tr>
                        <tr>
                            <th style="text-align:right;" colspan="3">School Year Comments:</th>
                            <td colspan="12"><%=sSchlCmnts%></td>
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
%></body>
</html>
