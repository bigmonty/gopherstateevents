<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim lEventID
Dim iYear
Dim sEventName, sSport, sChargeClock, sChargeCamera, sClearCamera, sChargeGoPro, sClearGoPro, sChargeBatteries
Dim sChargeTablet, sChargeBoxes, sClearLogFile, sPullData, sGetSoftware, sContactEventDir, sGetBibs, sPCUpdates
Dim sChargeJetPack, sUploadRaceFile, sUploadRFIDData, sGetExtraBibs, sPixInDropbox, sMsg
Dim dEventDate
Dim Events(), ChklstPre(14), ChklstSite(2), ChklstPost(0), Options(2), SortArr(3)
Dim bFound
Dim cdoMessage, cdoConfig

If Not Session("role") = "staff" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0

sSport = Request.QueryString("sport")
If sSport = vbNullString Then sSport = "Fitness Event"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Options(0) = "n"
Options(1) = "y"
Options(2) = "na"

i = 0
ReDim Events(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '" & Date - 14 & "' AND EventDate <= '" & Date + 7 & "' ORDER BY EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'")
    Events(2, i) = rs(2).Value
    Events(3, i) = "Fitness Event"
    i = i + 1
    ReDim Preserve Events(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE MeetDate >= '" & Date - 14 & "' AND MeetDate <= '" & Date + 7 & "' ORDER BY MeetDate"
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    Events(0, i) = rs(0).Value
    Events(1, i) = Replace(rs(1).Value, "''", "'")
    Events(2, i) = rs(2).Value
    Events(3, i) = rs(3).Value
    i = i + 1
    ReDim Preserve Events(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'order by date
For i = 0 To UBound(Events, 2) - 2
    For j = i + 1 To UBound(Events, 2) - 1
        If CDate(Events(2, i)) > CDate(Events(2, j)) Then
            For k = 0 To 3
                SortArr(k) = Events(k, i)
                Events(k, i) = Events(k, j)
                Events(k, j) = SortArr(k)
            Next
        End If
    Next
Next

If Request.Form.Item("submit_event") = "submit_event" Then
    For i = 1 To Len(Request.Form.Item("events"))
        If Mid(Request.Form.Item("events"), i, 1) = "_" Then
            Exit For
        Else
            lEventID = lEventID & Mid(Request.Form.Item("events"), i, 1)
        End If
    Next

    If Right(Request.Form.Item("events"), 13) = "Fitness Event" Then 
        sSport = "Fitness Event"

        sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
        Set rs = conn.Execute(sql)
        sEventName = Replace(rs(0).Value, "''", "'")
        dEventDate = rs(1).Value
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Sport, MeetName, MeetDate FROM Meets WHERE MeetsID = " & lEventID
        rs.Open sql, conn2, 1, 2
        sSport = rs(0).Value
        sEventName = Replace(rs(1).Value, "''", "'")
        dEventDate = rs(2).Value
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_pre-race") = "submit_pre-race" Then
    sChargeClock = Request.Form.Item("charge_clock")
    sChargeGoPro = Request.Form.Item("charge_gopro")
    sClearGoPro = Request.Form.Item("clear_gopro")
    sChargeCamera = Request.Form.Item("charge_camera")
    sClearCamera = Request.Form.Item("clear_camera")
    sChargeBatteries = Request.Form.Item("charge_batteries")
    sChargeBoxes = Request.Form.Item("charge_boxes")
    sClearLogFile = Request.Form.Item("clear_log_file")
    sPullData = Request.Form.Item("pull_data")
    sGetSoftware = Request.Form.Item("get_software")
    sContactEventDir = Request.Form.Item("contact_event_dir")
    sGetBibs = Request.Form.Item("get_bibs")
    sPCUpdates = Request.Form.Item("pc_updates")
    sChargeTablet = Request.Form.Item("charge_tablet")
    sChargeJetPack = Request.Form.Item("charge_jetpack")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ChargeClock, ChargeGoPro, ClearGoPro, ChargeCamera, ClearCamera, ChargeBatteries, ChargeBoxes, "
    sql = sql & "ClearLogFile, PullData, GetSoftware, ContactEventDir, GetBibs, PCUpdates, WhenSubmit, ChargeTablet, "
    sql = sql & "ChargeJetPack FROM StaffChklstPre WHERE EventID = "  & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    rs(0).Value = sChargeClock
    rs(1).Value = sChargeGoPro
    rs(2).Value = sClearGoPro
    rs(3).Value = sChargeCamera
    rs(4).Value = sClearCamera
    rs(5).Value = sChargeBatteries
    rs(6).Value = sChargeBoxes
    rs(7).Value = sClearLogFile
    rs(8).Value = sPullData
    rs(9).Value = sGetSoftware
    rs(10).Value = sContactEventDir
    rs(11).Value = sGetBibs
    rs(12).Value = sPCUpdates
    rs(13).Value = Now 
    rs(14).Value = sChargeTablet
    rs(15).Value = sChargeJetPack
    rs.Update
    rs.Close
    Set rs = Nothing

    Call EventInfo
    Call SendEmail("pre")
ElseIf Request.Form.Item("submit_on-site") = "submit_on-site" Then
    sUploadRaceFile = Request.Form.Item("upload_race_file")
    sUploadRFIDData = Request.Form.Item("upload_rfid_data")
    sGetExtraBibs = Request.Form.Item("upload_race_file")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UploadRaceFile, UploadRFIDData, GetExtraBibs, WhenSubmit FROM StaffChklstSite WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    rs(0).Value = sUploadRaceFile
    rs(1).Value = sUploadRFIDData
    rs(2).Value = sGetExtraBibs
    rs(3).Value = Now 
    rs.Update
    rs.Close
    Set rs = Nothing

    Call EventInfo
    Call SendEmail("site")
ElseIf Request.Form.Item("submit_post-race") = "submit_post-race" Then
    sPixInDropBox = Request.Form.Item("pix_in_dropbox")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PixInDropbox, WhenSubmit FROM STaffChklstPost WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    rs(0).Value = sPixInDropBox
    rs(1).Value = Now 
    rs.Update
    rs.Close
    Set rs = Nothing

    Call EventInfo
    Call SendEmail("post")
End If

Private Sub EventInfo
    If sSport = "Fitness Event" Then 
        sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
        Set rs = conn.Execute(sql)
        sEventName = Replace(rs(0).Value, "''", "'")
        dEventDate = rs(1).Value
        Set rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Sport, MeetName, MeetDate FROM Meets WHERE MeetsID = " & lEventID
        rs.Open sql, conn2, 1, 2
        sSport = rs(0).Value
        sEventName = Replace(rs(1).Value, "''", "'")
        dEventDate = rs(2).Value
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub SendEmail(sPhase)
    sMsg = Session("my_name") & " has just submitted a checklist for the " & sEventName & " on " & dEventDate 
    sMsg = sMsg & ".  Below is a summary: " & vbCrLf & vbCrLf

    Select Case sPhase
        Case "pre"
            sMsg = sMsg & "Which Checklist: Pre-Race " & vbCrLf & vbCrLf

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT ChargeClock, ChargeGoPro, ClearGoPro, ChargeCamera, ClearCamera, ChargeBatteries, ChargeBoxes, "
            sql = sql & "ClearLogFile, PullData, GetSoftware, ContactEventDir, GetBibs, PCUpdates, ChargeTablet, ChargeJetPack "
            sql = sql & "FROM StaffChklstPre WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
            rs.Open sql, conn, 1, 2
            For i = 0 To 14
                sMsg = sMsg & rs(i).Name & ": " & rs(i).Value & vbCrLf
            Next
            rs.Close
            Set rs = Nothing
        Case "site"
            sMsg = sMsg & "Which Checklist: Post Race On Site " & vbCrLf & vbCrLf

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT UploadRaceFile, UploadRFIDData, GetExtraBibs FROM StaffChklstSite WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
            rs.Open sql, conn, 1, 2
            For i = 0 To 2
                sMsg = sMsg & rs(i).Name & ": " & rs(i).Value & vbCrLf
            Next
            rs.Close
            Set rs = Nothing
        Case "post"
            sMsg = sMsg & "Which Checklist: Post Race Home " & vbCrLf & vbCrLf

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT PixInDropbox FROM StaffChklstPost WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
            rs.Open sql, conn, 1, 2
            sMsg = sMsg & rs(0).Name & ": " & rs(0).Value & vbCrLf
            rs.Close
            Set rs = Nothing
    End Select
%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
			 
	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.com"
        .Bcc = Session("my_email")
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "Staff Checklist: " & sEventName
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
End Sub

If Not CLng(lEventID) = 0 Then
    bFound = True
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ChargeClock, ChargeGoPro, ClearGoPro, ChargeCamera, ClearCamera, ChargeBatteries, ChargeBoxes, "
    sql = sql & "ClearLogFile, PullData, GetSoftware, ContactEventDir, GetBibs, PCUpdates, ChargeTablet, ChargeJetPack "
    sql = sql & " FROM StaffChklstPre WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        For i = 0 To 14
            ChklstPre(i) = rs(i).Value
        Next
    Else
        bFound = False
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO StaffChklstPre(EventID, Sport, StaffID) VALUES (" & lEventID & ", '" & sSport & "', " & Session("staff_id") & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If

    bFound = True
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UploadRaceFile, UploadRFIDData, GetExtraBibs FROM StaffChklstSite WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        For i = 0 To 2
            ChklstSite(i) = rs(i).Value
        Next
    Else
        bFound = False
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO StaffChklstSite(EventID, Sport, StaffID) VALUES (" & lEventID & ", '" & sSport & "', " & Session("staff_id") & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If

    bFound = True
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PixInDropbox FROM StaffChklstPost WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        ChklstPost(0) = rs(0).Value
    Else
        bFound = False
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO StaffChklstPost(EventID, Sport, StaffID) VALUES (" & lEventID & ", '" & sSport & "', " & Session("staff_id") & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Staff Checklist</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->
	<div class="row">
		<!--#include file = "../../staff/staff_menu.asp" -->
		<div class="col-sm-10">
			<h1 class="h1">GSE Staff Checklist</h1>
			
			<div style="margin:10px;">
				<form class="form-inline" role="form" name="which_event" method="post" action="staff_chklst.asp">
                <div class="form-group">
				    <label for="events">Select Event:</label>
				    <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
					    <option value="">&nbsp;</option>
                        <%For i = 0 To UBound(Events, 2) - 1%>
                            <%If CLng(Events(0, i)) = CLng(lEventID) Then%>
                                <option value="<%=Events(0, i)%>_<%=Events(3, i)%>" selected><%=Events(1, i)%> &nbsp;<%=Events(2, i)%>&nbsp (<%=Events(3, i)%>)</option>
                            <%Else%>
                                <option value="<%=Events(0, i)%>_<%=Events(3, i)%>"><%=Events(1, i)%> &nbsp;<%=Events(2, i)%>&nbsp (<%=Events(3, i)%>)</option>
                            <%End If%>
                        <%Next%>
				    </select>
				    <input type="hidden" class="form-control" name="submit_event" id="submit_event" value="submit_event">
				    <input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
                </div>
				</form>
			</div>
			
			<%If Not Clng(lEventID) = 0 Then%>	
                <div class="row">
                    <div class="col-sm-6">
                        <div class="bg-warning">
                            <h3 class="h3">Phase 1: Pre-Race</h3>

                            <form class="form" role="form" name="update_pre-race" method="post" action="staff_chklst.asp?event_id=<%=lEventID%>&amp;sport=<%=sSport%>&amp;year=<%=iYear%>">
                            <table class="table">
                                <tr><th>Item</th><th>Status</th></tr>
                                <tr>
                                    <td>Charge Clock</td>
                                    <td>
                                        <select class="form-control" name="charge_clock" id="charge_clock">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(0) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Charge GoPro</td>
                                    <td>
                                        <select class="form-control" name="charge_gopro" id="charge_gopro">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(1) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td>Clear GoPro</td>
                                    <td>
                                        <select class="form-control" name="clear_gopro" id="clear_gopro">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(2) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Charge Camera</td>
                                    <td>
                                        <select class="form-control" name="charge_camera" id="charge_camera">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(3) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Clear Camera</td>
                                    <td>
                                        <select class="form-control" name="clear_camera" id="clear_camera">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(4) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Charge Batteries</td>
                                    <td>
                                        <select class="form-control" name="charge_batteries" id="charge_batteries">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(5) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Charge Timing Box(es)</td>
                                    <td>
                                        <select class="form-control" name="charge_boxes" id="charge_boxes">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(6) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Clear Log File(s)</td>
                                    <td>
                                        <select class="form-control" name="clear_log_file" id="clear_log_file">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(7) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Pull Data</td>
                                    <td>
                                        <select class="form-control" name="pull_data" id="pull_data">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(8) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Get Software</td>
                                    <td>
                                        <select class="form-control" name="get_software" id="get_software">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(9) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Contact Event Director</td>
                                    <td>
                                        <select class="form-control" name="contact_event_dir" id="contact_event_dir">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(10) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Get Bibs</td>
                                    <td>
                                        <select class="form-control" name="get_bibs" id="get_bibs">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(11) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Update PC</td>
                                    <td>
                                        <select class="form-control" name="pc_updates" id="pc_updates">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(12) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Charge Tablet</td>
                                    <td>
                                        <select class="form-control" name="charge_tablet" id="charge_tablet">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(13) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Charge Jet Pack</td>
                                    <td>
                                        <select class="form-control" name="charge_jetpack" id="charge_jetpach">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPre(14) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                           </table>
                            <div class="control-group">
                                <input type="hidden" class="form-control" name="submit_pre-race" id="submit_pre-race" value="submit_pre-race">
                                <input type="submit" class="form-control" name="submit1" id="submit1" value="Submit This">
                            </div>
                            </form>
                        </div>
                    </div>
                    <div class="col-sm-6">
                        <div class="bg-info">
                            <h3 class="h3">Phase 2: Post-Race On-Site</h3>

                            <form class="form" role="form" name="update_on-site" method="post" action="staff_chklst.asp?event_id=<%=lEventID%>&amp;sport=<%=sSport%>&amp;year=<%=iYear%>">
                            <table class="table">
                                <tr><th>Item</th><th>Status</th></tr>
                                 <tr>
                                    <td>Upload Race File To Dropbox</td>
                                    <td>
                                        <select class="form-control" name="upload_race_file" id="upload_race_file">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstSite(0) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Upload RFID Data</td>
                                    <td>
                                        <select class="form-control" name="upload_rfid_data" id="upload_rfid_data">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstSite(1) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                                 <tr>
                                    <td>Get Extra Bibs</td>
                                    <td>
                                        <select class="form-control" name="get_extra_bibs" id="get_extra_bibs">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstSite(2) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                            </table>
                            <div class="control-group">
                                <input type="hidden" class="form-control" name="submit_on-site" id="submit_on-site" value="submit_on-site">
                                <input type="submit" class="form-control" name="submit2" id="submit2" value="Submit This">
                            </div>
                            </form>
                        </div>

                        <div class="bg-success">
                            <h3 class="h3">Phase 3: Post-Race At Home</h3>

                            <form class="form" role="form" name="update_post-race" method="post" action="staff_chklst.asp?event_id=<%=lEventID%>&amp;sport=<%=sSport%>&amp;year=<%=iYear%>">
                            <table class="table">
                                <tr><th>Item</th><th>Status</th></tr>
                                 <tr>
                                    <td>Pix in Dropbox</td>
                                    <td>
                                        <select class="form-control" name="pix_in_dropbox" id="pix_in_dropbox">
                                            <%For i = 0 To UBound(Options)%>
                                                <%If ChklstPost(0) = Options(i) Then%>
                                                    <option value="<%=Options(i)%>" selected><%=Options(i)%></option>
                                                <%Else%>
                                                    <option value="<%=Options(i)%>"><%=Options(i)%></option>
                                                <%End If%>
                                            <%Next%>
                                        </select>
                                    </td>
                                </tr>
                            </table>
                            <div class="control-group">
                                <input type="hidden" class="form-control" name="submit_post-race" id="submit_post-race" value="submit_post-race">
                                <input type="submit" class="form-control" name="submit3" id="submit3" value="Submit This">
                            </div>
                            </form>
                        </div>
                    </div>
                </div>
			<%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>