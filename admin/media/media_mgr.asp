<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, conn2, rs, sql, rs2, sql2
Dim lEventID, lMeetID
Dim i
Dim sEventRaces, sEventType
Dim iEvntTtl
Dim Events(), Meets(), MediaOrders(), Delete()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
		
lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
		
lMeetID = Request.QueryString("meet_id")
If CStr(lMeetID) = vbNullString Then lMeetID = 0
	
iEvntTtl = 0
			
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
							
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Not CLng(lEventID) = 0 Then
        sql = "SELECT MediaOrderID, BibNum, Email, AmtPd, MediaType, Processed FROM MediaOrder WHERE EventID = " & lEventID
        rs.Open sql, conn, 1, 2
    Else
        sql = "SELECT MediaOrderID, BibNum, Email, AmtPd, MediaType, Processed FROM MediaOrder WHERE MeetID = " & lMeetID
        rs.Open sql, conn2, 1, 2
    End If
    
    Do While Not rs.EOF
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
            rs(1).Value = Request.Form.Item("bib_num_" & rs(0).Value)
            rs(2).Value = Request.Form.Item("email_" & rs(0).Value)
            rs(3).Value = Request.Form.Item("amt_pd_" & rs(0).Value)
            rs(4).Value = Request.Form.Item("media_type_" & rs(0).Value)
            rs(5).Value = Request.Form.Item("processed_" & rs(0).Value)
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM MediaOrder WHERE MediaOrderID =  " & Delete(i)
        If Not CLng(lEventID) = 0 Then
            Set rs = conn.Execute(sql)
        Else
            Set rs = conn2.Execute(sql)
        End If
        Set rs = Nothing
    Next
ElseIf Request.Form.Item("submit_order") = "submit_order" Then
    Dim iBibNum, iAmtPd
    Dim sEmail, sMediaType, sGunTime

    iBibNum = Request.Form.Item("bib_num")
    iAmtPd = Request.Form.Item("amt_pd")
    If CStr(iAmtPd) = vbNullString Then iAmtPd = 0

    sEmail = Request.Form.Item("email")
    sMediaType = Request.Form.Item("media_type")
    sGunTime = Request.Form.Item("gun_time")

    'write to table
    If CLng(lMeetID) = 0 Then
        sql = "INSERT INTO MediaOrder(MeetID, BibNum, Email, AmtPd, MediaType, WhenOrdered) VALUES (" & lMeetID & ", " & iBibNum & ", '" & sEmail 
        sql = sql & "', " & iAmtPd & ", '" & sMediaType & "')"
        Set rs = conn.Execute(sql)
    Else
        sql = "INSERT INTO MediaOrder(MeetID, BibNum, Email, AmtPd, MediaType, WhenOrdered) VALUES (" & lMeetID & ", " & iBibNum & ", '" & sEmail 
        sql = sql & "', " & iAmtPd & ", '" & sMediaType & "')"
        Set rs = conn2.Execute(sql)
    End If
    Set rs = Nothing
ElseIf Request.form.Item("submit_event") = "submit_event" Then
    lEventID = Request.Form.Item("events")
    If CStr(lEventID) = vbNullString Then lEventID = 0
ElseIf Request.Form.Item("submit_meet") = "submit_meet" Then
    lMeetID = Request.Form.Item("meets")
    If CStr(lMeetID) = vbNullString Then lMeetID = 0
End If

'get fitness events
i = 0
ReDim Events(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate > '9/1/2013' AND EventDate < '" & Date & "' ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
    Events(2, i) = "fitness"
	i = i + 1
	ReDim Preserve Events(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get cc events
i = 0
ReDim Meets(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate > '9/1/2013' AND MeetDate < '" & Date & "' ORDER By MeetDate DESC"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
	Meets(0, i) = rs(0).Value
	Meets(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
    Meets(2, i) = "cc"
	i = i + 1
	ReDim Preserve Meets(2, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get order history
If Not (CLng(lEventID) = 0 AND CLng(lMeetID) = 0) Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Not CLng(lEventID) = 0 Then
        sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
        rs.Open sql, conn, 1, 2
    Else
        sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lMeetID
        rs.Open sql, conn2, 1, 2
    End If

    Do While Not rs.EOF
        sEventRaces = sEventRaces & rs(0).Value & ", "
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If Len(sEventRaces) > 0 Then
        sEventRaces = Trim(sEventRaces)
        sEventRaces = Left(sEventRaces, Len(sEventRaces) - 1)
    End If

    i = 0
    ReDim MediaOrders(8, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Not CLng(lEventID) = 0 Then
        sEventType = "fitness"
        sql = "SELECT MediaOrderID, BibNum, Email, MediaType, Processed, AmtPd FROM MediaOrder WHERE EventID = " & lEventID
        sql = sql & " ORDER BY Processed, MediaOrderID"
        rs.Open sql, conn, 1, 2
    Else
        sEventType = "ccmeet"
        sql = "SELECT MediaOrderID, BibNum, Email, MediaType, Processed, AmtPd FROM MediaOrder WHERE MeetID = " & lMeetID
        sql = sql & " ORDER BY Processed, MediaOrderID"
        rs.Open sql, conn2, 1, 2
    End If
    Do While Not rs.EOF
        iEvntTtl = iEvntTtl + rs(5).Value

        MediaOrders(0, i) = rs(0).Value 
        MediaOrders(1, i) = rs(1).Value 
        If rs(2).Value & "" = "" Then
            If sEventType = "fitness" Then MediaOrders(2, i) = GetEmail(rs(1).Value)
        Else
            MediaOrders(2, i) = rs(2).Value  
        End If
        MediaOrders(3, i) = rs(3).Value
        MediaOrders(4, i) = GetGunTime(sEventType, rs(1).Value)
        MediaOrders(5, i) = GetPartName(sEventType, rs(1).Value)
        MediaOrders(6, i) = GetRacePlaceTime(sEventType, rs(1).Value)
        MediaOrders(7, i) = rs(4).Value
        MediaOrders(8, i) = rs(5).Value
        i = i + 1
        ReDim Preserve MediaOrders(8, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function GetRacePlaceTime(sThisEventType, iThisBib)
    GetRacePlaceTime = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    If sThisEventType = "fitness" Then
        sql2 = "SELECT r.RaceName, ir.EventPl, ir.FnlTime, ir.ChipTime FROM RaceData r INNER JOIN IndResults ir ON r.RaceID = ir.RaceID "
        sql2 = sql2 & "INNER JOIN PartRace pr ON pr.ParticipantID = ir.ParticipantID WHERE r.RaceID IN ("& sEventRaces & ") AND pr.Bib = " & iThisBib
        sql = sql & " AND ir.RaceID IN ("& sEventRaces & ")" 
        rs2.Open sql2, conn, 1, 2
    Else
        sql2 = "SELECT r.RaceName, ir.Place, ir.RaceTime FROM Races r INNER JOIN IndRslts ir ON r.RacesID = ir.RacesID WHERE r.RacesID IN (" & sEventRaces 
        sql2 = sql2 & ") AND ir.RaceID IN ("& sEventRaces & ") AND ir.Bib = " & iThisBib
        rs2.Open sql2, conn2, 1, 2
    End If
    If rs2.RecordCount > 0 Then 
        If rs2(3).Value & "" = "" OR ConvertToSeconds(rs2(3).Value) = 0 Then
            GetRacePlaceTime = rs2(0).Value & "; Pl " & rs2(1).Value & "; Time: " & rs2(2).Value
        Else
            GetRacePlaceTime = rs2(0).Value & "; Pl " & rs2(1).Value & "; Time: " & rs2(3).Value
        End If
    End If
    rs2.Close
    Set rs2 = Nothing
End Function

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<%
    
Private Function GetPartName(sThisEventType, iThisBib)
    GetPartName = vbNullString

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    If sThisEventType = "fitness" Then
        sql2 = "SELECT p.FirstName, p.LastName FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID WHERE pr.RaceID IN (" 
        sql2 = sql2 & sEventRaces & ") AND pr.Bib = " & iThisBib
        rs2.Open sql2, conn, 1, 2
    Else
        sql2 = "SELECT r.FirstName, r.LastName FROM Roster r INNER JOIN IndRslts ir ON r.RosterID = ir.RosterID WHERE ir.RacesID IN (" & sEventRaces 
        sql2 = sql2 & ") AND ir.Bib = " & iThisBib
        rs2.Open sql2, conn2, 1, 2
    End If
    If rs2.RecordCount > 0 Then GetPartName = rs2(0).Value & " " & rs2(1).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetGunTime(sThisEventType, iThisBib)
    GetGunTime = "00:00.00"

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    If sThisEventType = "fitness" Then
        sql2 = "SELECT ir.FnlTime FROM IndResults ir INNER JOIN PartRace pr ON ir.ParticipantID = pr.ParticipantID WHERE ir.RaceID IN (" & sEventRaces 
        sql2 = sql2 & ") AND pr.RaceID IN ("& sEventRaces & ") AND pr.Bib = " & iThisBib
        rs2.Open sql2, conn, 1, 2
    Else
        sql2 = "SELECT ElpsdTime FROM IndRslts WHERE RacesID IN (" & sEventRaces & ") AND Bib = " & iThisBib
        rs2.Open sql2, conn2, 1, 2
    End If
    If rs2.RecordCount > 0 Then GetGunTime = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetEmail(iThisBib)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT p.Email FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID WHERE pr.RaceID IN (" & sEventRaces 
    sql2 = sql2 & ") AND pr.Bib = " & iThisBib
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then GetEmail = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function HasVids(lThisEvent, sThisEventType)
    HasVids = "n"

    Set rs2 = Server.CreateObject("ADODB.Recordset")

    If sThisEventType = "fitness" Then
        sql2 = "SELECT RaceVidsID FROM RaceVids WHERE EventID = " & lThisEvent
        rs2.Open sql2, conn, 1, 2
    Else
        sql2 = "SELECT RaceVidsID FROM RaceVids WHERE MeetsID = " & lThisEvent
        rs2.Open sql2, conn2, 1, 2
    End If
        
    If rs2.RecordCount > 0 Then HasVids = "y"
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Media Manager Utility</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    th, td{
        padding-left: 2px;
    }
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">Gopher State Events Media Manager</h4>
		
            <div style="float:left;width: 350px;font-size: 0.95em;">
                <h4 class="h4">Fitness Events</h4>

		        <form name="which_event" method="post" action="media_mgr.asp?event_id=<%=lEventID%>">
		        <span style="font-weight:bold;">Event:</span>
		        <select name="events" id="events" onchange="this.form.get_video.click()" style="font-size:0.9em;">
			        <option value="">&nbsp;</option>
			        <%For i = 0 to UBound(Events, 2) - 1%>
				        <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
					        <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
				        <%Else%>
					        <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
				        <%End If%>
			        <%Next%>
		        </select>
		        <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
		        <input type="submit" name="get_video" id="get_video" value="Get Orders" style="font-size:0.8em;">
		        </form>
		    </div>
            <div style="margin-left:360px;font-size: 0.9em;">
                <h4 class="h4">CC/Nordic Meets</h4>

		        <form name="which_meet" method="post" action="media_mgr.asp?meet_id=<%=lMeetID%>">
		        <span style="font-weight:bold;">Meet:</span>
		        <select name="meets" id="meets" onchange="this.form.get_video1.click()" style="font-size:0.9em;">
			        <option value="">&nbsp;</option>
			        <%For i = 0 to UBound(Meets, 2) - 1%>
				        <%If CLng(lMeetID) = CLng(Meets(0, i)) Then%>
					        <option value="<%=Meets(0, i)%>" selected><%=Meets(1, i)%></option>
				        <%Else%>
					        <option value="<%=Meets(0, i)%>"><%=Meets(1, i)%></option>
				        <%End If%>
			        <%Next%>
		        </select>
		        <input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
		        <input type="submit" name="get_video1" id="get_video1" value="Get Orders" style="font-size:0.8em;">
		        </form>
            </div>

            <%If Not (CLng(lMeetID) = 0 AND CLng(lEventID) = 0) Then%>
                <form name="order_video" method="post" action="media_mgr.asp?meet_if=<%=lMeetID%>&amp;event_id=<%=lEventID%>" 
                    style="background-color: #ececd8;padding: 5px;margin-top: 10px;">
                <h4 class="h4">New Order:</h4>

                <table>
                    <tr>
                        <th>Bib No:</th><td><input type="text" name="bib_num" id="bib_num" size="3"></td>
                        <th>Email:</th><td><input type="text" name="email" id="email" size="35"></td>
                        <th>Gun Time:</th><td><input type="text" name="gun_time" id="gun_time" size="6"></td>
                         <td style="padding-left:10px;">
                             <fieldset style="padding: 5px;">
                                <legend>Media Type:</legend>
                                <input type="radio" name="media_type" id="media_type" value="still">&nbsp;Still
                                <input type="radio" name="media_type" id="media_type" value="video">&nbsp;Video
                                <input type="radio" name="media_type" id="media_type" value="both" checked>&nbsp;Both
                            </fieldset>
                        </td>
                        <th>Amt Pd:&nbsp;$</th><td><input type="text" name="amt_pd" id="amt_pd" size="3"></td>
                        <td style="padding-left: 10px;">
                            <input type="hidden" name="submit_order" id="submit_order" value="submit_order">
                            <input type="submit" name="submit1" id="submit1" value="Order Media">
                        </td>
                    </tr>
                </table>
                </form>

                <form name="media_orders" method="post" action="media_mgr.asp?meet_id=<%=lMeetID%>&amp;event_id=<%=lEventID%>">
                <div style="margin: 10px 0 0 0;padding: 0;font-size: 0.85em;text-align: right;">
                    <a href="javascript:pop('print_orders.asp?event_id=<%=lEventID%>&amp;meet_id=<%=lMeetID%>',800,600)">Print</a>
                </div>
                <h4 class="h4">Existing Orders:</h4>
                <h5>Num Orders:&nbsp;<%=UBound(MediaOrders, 2)%> ($<%=iEvntTtl%>)</h5>

                <p style="font-weight: bold;margin: 10px 0 0 0;">Picture Specs:</p>
                <ul style="font-size:0.8em;margin-bottom: 10px;">
                    <li>Image Size: 400x550</li>
                    <li>Name: black, bold, italic, 26pt, emboss, dropshadow</li>
                    <li>Race Name: black, bold, 24pt, emboss</li>
                    <li>Race Name: white, 16pt</li>
                </ul>
                <table  style="font-size: 0.85em;">
                    <tr>
                        <th>Bib</th><th style="text-align: left;">Email</th><th>Type</th><th>GunTime</th><th>Name</th><th>Race Place Time</th>
                        <th>Procssd</th><th>Amt</th><th>Delete</th>
                    </tr>
                    <%For i = 0 To UBound(MediaOrders, 2) - 1%>
                        <tr>
                            <%If MediaOrders(7, i) = "y" Then%>
                                <td style="background-color: #000;color: #fff;"><input type="text" name="bib_num_<%=MediaOrders(0, i)%>" 
                                    id="bib_num_<%=MediaOrders(0, i)%>" value="<%=MediaOrders(1, i)%>" size="2" style="background-color: #000;color: #fff;"></td>
                                <td style="background-color: #000;color: #fff;"><input type="text" name="email_<%=MediaOrders(0, i)%>" 
                                    id="email_<%=MediaOrders(0, i)%>" value="<%=MediaOrders(2, i)%>" size="20" style="background-color: #000;color: #fff;"></td>
                                <td style="background-color: #000;color: #fff;">
                                    <select name="media_type_<%=MediaOrders(0, i)%>" id="media_type_<%=MediaOrders(0, i)%>"
                                     style="background-color: #000;color: #fff;">
                                        <option value="">&nbsp;</option>
                                        <%Select Case MediaOrders(3, i)%>
                                            <%Case "still"%>
                                                <option value="still" selected>Still</option>
                                                <option value="video">Video</option>
                                                <option value="both">Both</option>
                                            <%Case "video"%>
                                                <option value="still">Still</option>
                                                <option value="video" selected>Video</option>
                                                <option value="both">Both</option>
                                            <%Case "both"%>
                                                <option value="still">Still</option>
                                                <option value="video">Video</option>
                                                <option value="both" selected>Both</option>
                                            <%Case Else%>
                                                <option value="still">Still</option>
                                                <option value="video">Video</option>
                                                <option value="both">Both</option>
                                        <%End Select%>
                                    </select>
                                </td>
                                <td style="background-color: #000;color: #fff;"><%=MediaOrders(4, i)%></td>
                                <td style="background-color: #000;color: #fff;"><%=MediaOrders(5, i)%></td>
                                <td style="background-color: #000;color: #fff;"><%=MediaOrders(6, i)%></td>
                                <td style="background-color: #000;color: #fff;" style="background-color: #000;color: #fff;">
                                    <select name="processed_<%=MediaOrders(0, i)%>" id="processed_<%=MediaOrders(0, i)%>" style="background-color: #000;color: #fff;">
                                        <%If MediaOrders(7, i) = "y" Then%>
                                            <option value="y" selected>Yes</option>
                                            <option value="n">No</option>
                                        <%Else%>
                                            <option value="y">Yes</option>
                                            <option value="n" selected>No</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td style="background-color: #000;color: #fff;white-space: nowrap;">$<input type="text" name="amt_pd_<%=MediaOrders(0, i)%>" 
                                    id="amt_pd_<%=MediaOrders(0, i)%>" value="<%=MediaOrders(8, i)%>" size="1" style="background-color: #000;color: #fff;"></td>
                                <td style="background-color: #000;color: #fff;text-align: center;"><input type="checkbox" name="delete_<%=MediaOrders(0, i)%>" id="delete_<%=MediaOrders(0, i)%>"></td>
                            <%Else%>
                                <td><input type="text" name="bib_num_<%=MediaOrders(0, i)%>" id="bib_num_<%=MediaOrders(0, i)%>" value="<%=MediaOrders(1, i)%>" size="2"></td>
                                <td><input type="text" name="email_<%=MediaOrders(0, i)%>" id="email_<%=MediaOrders(0, i)%>" value="<%=MediaOrders(2, i)%>" size="20"></td>
                                <td>
                                    <select name="media_type_<%=MediaOrders(0, i)%>" id="media_type_<%=MediaOrders(0, i)%>">
                                        <option value="">&nbsp;</option>
                                        <%Select Case MediaOrders(3, i)%>
                                            <%Case "still"%>
                                                <option value="still" selected>Still</option>
                                                <option value="video">Video</option>
                                                <option value="both">Both</option>
                                            <%Case "video"%>
                                                <option value="still">Still</option>
                                                <option value="video" selected>Video</option>
                                                <option value="both">Both</option>
                                            <%Case "both"%>
                                                <option value="still">Still</option>
                                                <option value="video">Video</option>
                                                <option value="both" selected>Both</option>
                                            <%Case Else%>
                                                <option value="still">Still</option>
                                                <option value="video">Video</option>
                                                <option value="both">Both</option>
                                        <%End Select%>
                                    </select>
                                </td>
                                <td><%=MediaOrders(4, i)%></td>
                                <td><%=MediaOrders(5, i)%></td>
                                <td><%=MediaOrders(6, i)%></td>
                                <td>
                                    <select name="processed_<%=MediaOrders(0, i)%>" id="processed_<%=MediaOrders(0, i)%>">
                                        <%If MediaOrders(7, i) = "y" Then%>
                                            <option value="y" selected>Yes</option>
                                            <option value="n">No</option>
                                        <%Else%>
                                            <option value="y">Yes</option>
                                            <option value="n" selected>No</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td style="white-space: nowrap;">$<input type="text" name="amt_pd_<%=MediaOrders(0, i)%>" id="amt_pd_<%=MediaOrders(0, i)%>" value="<%=MediaOrders(8, i)%>" size="1"></td>
                                <td style="text-align: center;"><input type="checkbox" name="delete_<%=MediaOrders(0, i)%>" id="delete_<%=MediaOrders(0, i)%>"></td>
                            <%End If%>
                        </tr>
                    <%Next%>
                    <tr>
                        <td style="text-align: center;background-color: #ececec;" colspan="9">
                            <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                            <input type="submit" name="submit2" id="submit2" value="Save Changes">
                        </td>
                    </tr>
                </table>
                </form>
            <%End If%>
        </div>
	</div>
	<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>