<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim sEventName, sContractcmnts, sDepositcmnts, sDatacmnts, sBibcmnts, sEventEmailCmnts, sStaffEmailCmnts, sPartDataCmnts, sBibLabelCmnts
Dim sBibListCmnts, sPreRaceCmnts, sTimercmnts, sClockscmnts, sCameraCmnts, sGoProCmnts, sBatteriesCmnts, sErrorCmnts, sUploadPixCmnts
Dim sPixNotifCmnts, sInvoiceCmnts, sPromoCmnts, sFinanceCmnts, sPacketPrepCmnts, sUpdateSeriesCmnts
Dim dEventDate, dContractSent, dDepositReceived, dDataCollected, dBibsPrepped, dEventEmail, dStaffEmail, dPartData, dBibLabels, dBibList, dPreRaceParts
Dim dChargeTimers, dChargeClocks, dChargeCamera, dChargeGoPro, dChargeBatteries, dResolveErrors, dUploadPix, dPixNotif, dSendInvoice, dEventPromo
Dim dUpdateFinances, dPacketPrep, dUpdateSeries
Dim Events(), TimeLine(45)
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate > '" & Date - 7 & "' ORDER By EventDate"
rs.Open sql, conn, 1, 2
Do While Not rs.eOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
	If CStr(lEventID) = vbNullString Then lEventID = 0
ElseIf Request.Form.Item("submit_settings") = "submit_settings" Then
    If Not Request.Form.Item("contract_cmnts") = vbNullString Then sContractcmnts = Replace(Request.Form.Item("contract_cmnts"), "'", "''")
    If Not Request.Form.Item("deposit_cmnts") = vbNullString Then sDepositcmnts = Replace(Request.Form.Item("deposit_cmnts"), "'", "''")
    If Not Request.Form.Item("data_cmnts") = vbNullString Then sDatacmnts = Replace(Request.Form.Item("data_cmnts"), "'", "''")
    If Not Request.Form.Item("bib_cmnts") = vbNullString Then sBibcmnts = Replace(Request.Form.Item("bib_cmnts"), "'", "''")
    If Not Request.Form.Item("event_email_cmnts") = vbNullString Then sEventEmailCmnts = Replace(Request.Form.Item("event_email_cmnts"), "'", "''")
    If Not Request.Form.Item("staff_email_cmnts") = vbNullString Then sStaffEmailCmnts = Replace(Request.Form.Item("staff_email_cmnts"), "'", "''")
    If Not Request.Form.Item("part_data_cmnts") = vbNullString Then sPartDataCmnts = Replace(Request.Form.Item("part_data_cmnts"), "'", "''")
    If Not Request.Form.Item("bib_label_cmnts") = vbNullString Then sBibLabelCmnts = Replace(Request.Form.Item("bib_label_cmnts"), "'", "''")
    If Not Request.Form.Item("bib_list_cmnts") = vbNullString Then sBibListCmnts = Replace(Request.Form.Item("bib_list_cmnts"), "'", "''")
    If Not Request.Form.Item("pre_race_cmnts") = vbNullString Then sPreRaceCmnts = Replace(Request.Form.Item("pre_race_cmnts"), "'", "''")
    If Not Request.Form.Item("timer_cmnts") = vbNullString Then sTimercmnts = Replace(Request.Form.Item("timer_cmnts"), "'", "''")
    If Not Request.Form.Item("clocks_cmnts") = vbNullString Then sClockscmnts = Replace(Request.Form.Item("clocks_cmnts"), "'", "''")
    If Not Request.Form.Item("camera_cmnts") = vbNullString Then sCameraCmnts = Replace(Request.Form.Item("camera_cmnts"), "'", "''")
    If Not Request.Form.Item("go-pro_cmnts") = vbNullString Then sGoProCmnts = Replace(Request.Form.Item("go-pro_cmnts"), "'", "''")
    If Not Request.Form.Item("batteries_cmnts") = vbNullString Then sBatteriesCmnts = Replace(Request.Form.Item("batteries_cmnts"), "'", "''")
    If Not Request.Form.Item("error_cmnts") = vbNullString Then sErrorCmnts = Replace(Request.Form.Item("error_cmnts"), "'", "''")
    If Not Request.Form.Item("upload_pix_cmnts") = vbNullString Then sUploadPixCmnts = Replace(Request.Form.Item("upload_pix_cmnts"), "'", "''")
    If Not Request.Form.Item("pix_notif_cmnts") = vbNullString Then sPixNotifCmnts = Replace(Request.Form.Item("pix_notif_cmnts"), "'", "''")
    If Not Request.Form.Item("invoice_cmnts") = vbNullString Then sInvoiceCmnts = Replace(Request.Form.Item("invoice_cmnts"), "'", "''")
    If Not Request.Form.Item("promo_cmnts") = vbNullString Then sPromoCmnts = Replace(Request.Form.Item("promo_cmnts"), "'", "''")
    If Not Request.Form.Item("finance_cmnts") = vbNullString Then sFinanceCmnts = Replace(Request.Form.Item("finance_cmnts"), "'", "''")
    If Not Request.Form.Item("packet_prep_cmnts") = vbNullString Then sPacketPrepCmnts = Replace(Request.Form.Item("packet_prep_cmnts"), "'", "''")
    If Not Request.Form.Item("update_series_cmnts") = vbNullString Then sUpdateSeriesCmnts = Replace(Request.Form.Item("update_series_cmnts"), "'", "''")

    dContractSent = Request.Form.Item("contract_sent")
    dDepositReceived = Request.Form.Item("deposit_received")
    dDataCollected = Request.Form.Item("data_collected")
    dBibsPrepped = Request.Form.Item("bibs_prepped")
    dEventEmail = Request.Form.Item("event_email")
    dStaffEmail = Request.Form.Item("staff_email")
    dPartData = Request.Form.Item("part_data")
    dBibLabels = Request.Form.Item("bib_labels")
    dBibList = Request.Form.Item("bib_list")
    dPreRaceParts = Request.Form.Item("pre_race_parts")
    dChargeTimers = Request.Form.Item("charge_timers")
    dChargeClocks = Request.Form.Item("charge_clocks")
    dChargeCamera = Request.Form.Item("charge_camera")
    dChargeGoPro = Request.Form.Item("charge_go-pro")
    dChargeBatteries = Request.Form.Item("charge_batteries")
    dResolveErrors = Request.Form.Item("resolve_errors")
    dUploadPix = Request.Form.Item("upload_pix")
    dPixNotif = Request.Form.Item("pix_notif")
    dSendInvoice = Request.Form.Item("send_invoice")
    dEventPromo = Request.Form.Item("event_promo")
    dUpdateFinances = Request.Form.Item("update_finances")
    dPacketPrep = Request.Form.Item("packet_prep")
    dUpdateSeries = Request.Form.Item("update_series")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ContractSent, ContractCmnts, DepositReceived, DepositCmnts, DataCollected, DataCmnts, BibsPrepped, BibCmnts, EventEmail, "
    sql = sql & "EventEmailCmnts, StaffEmail, StaffEmailCmnts, PartData, PartDataCmnts, BibLabels, BibLabelCmnts, PreRaceParts, PreRaceCmnts, "
    sql = sql & "ChargeTimers, TimersCmnts, ChargeClocks, ClocksCmnts, ChargeBatteries, BatteriesCmnts, ChargeGoPro, GoProCmnts, ChargeCamera, "
    sql = sql & "CameraCmnts, ResolveErrors, ErrorCmnts, UploadPix, UploadPixCmnts, PixNotif, PixNotifCmnts, SendInvoice, InvoiceCmnts, "
    sql = sql & "EventPromo, PromoCmnts, UpdateFinances, FinanceCmnts, BibList, BibListCmnts, PacketPrep, PacketPrepCmnts, UpdateSeries, UpdateSeriesCmnts "
    sql = sql & "FROM EventTimeline WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If Not dContractSent & "" = "" Then rs(0).Value = dContractSent
    rs(1).Value = sContractCmnts
    If Not dDepositReceived & "" = "" Then rs(2).Value = dDepositReceived
    rs(3).Value = sDepositCmnts
    If Not dDataCollected & "" = "" Then rs(4).Value = dDataCollected
    rs(5).Value = sDataCmnts
    If Not dBibsPrepped & "" = "" Then rs(6).Value = dBibsPrepped
    rs(7).Value = sBibCmnts
    If Not dEventEmail & "" = "" Then rs(8).Value = dEventEmail
    rs(9).Value = sEventEmailCmnts
    If Not dStaffEmail & "" = "" Then rs(10).Value = dStaffEmail
    rs(11).Value = sStaffEmailCmnts
    If Not dPartData & "" = "" Then rs(12).Value = dPartData
    rs(13).Value = sPartDataCmnts
    If Not dBibLabels & "" = "" Then rs(14).Value = dBibLabels
    rs(15).Value = sBibLabelCmnts
    If Not dPreRaceParts & "" = "" Then rs(16).Value = dPreRaceParts
    rs(17).Value = sPreRaceCmnts
    If Not dChargeTimers & "" = "" Then rs(18).Value = dChargeTimers
    rs(19).Value = sTimerCmnts
    If Not dChargeClocks & "" = "" Then rs(20).Value = dChargeClocks
    rs(21).Value = sClockscmnts
    If Not dChargeBatteries & "" = "" Then rs(22).Value = dChargeBatteries
    rs(23).Value = sBatteriesCMnts
    If Not dChargeGoPro & "" = "" Then rs(24).Value = dChargeGoPro
    rs(25).Value = sGoProCmnts
    If Not dChargeCamera & "" = "" Then rs(26).Value = dChargeCamera
    rs(27).Value = sCameraCmnts
    If Not dResolveErrors & "" = "" Then rs(28).Value = dResolveErrors
    rs(29).Value = sErrorCmnts
    If Not dUploadPix & "" = "" Then rs(30).Value = dUploadPix
    rs(31).Value = sUploadPixCmnts
    If Not dPixNotif & "" = "" Then rs(32).Value = dPixNotif
    rs(33).Value = sPixNotifCmnts
    If Not dSendInvoice & "" = "" Then rs(34).Value = dSendInvoice
    rs(35).Value = sInvoiceCmnts
    If Not dEventPromo & "" = "" Then rs(36).Value = dEventPromo
    rs(37).Value = sPromoCmnts
    If Not dUpdateFinances & "" = "" Then rs(38).Value = dUpdateFinances
    rs(39).Value = sFinanceCmnts
    If Not dBibList & "" = "" Then rs(40).Value = dBibList
    rs(41).Value = sBibListCmnts
    If Not dPacketPrep & "" = "" Then rs(42).Value = dPacketPrep
    rs(43).Value = sPacketPrepCmnts
    If Not dUpdateSeries & "" = "" Then rs(44).Value = dUpdateSeries
    rs(45).Value = sUpdateSeriesCmnts
    rs.Update
    rs.Close
    Set rs = Nothing
End If

If Not CLng(lEventID) = 0 Then
	sql = "SELECT EventName, EventDate, EventDirID FROM Events WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	sEventName = Replace(rs(0).Value, "''", "'")
    dEventDate = rs(1).Value
	Set rs = Nothing

    bFound = True
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ContractSent, ContractCmnts, DepositReceived, DepositCmnts, DataCollected, DataCmnts, BibsPrepped, BibCmnts, EventEmail, "
    sql = sql & "EventEmailCmnts, StaffEmail, StaffEmailCmnts, PartData, PartDataCmnts, BibLabels, BibLabelCmnts, PreRaceParts, PreRaceCmnts, "
    sql = sql & "ChargeTimers, TimersCmnts, ChargeClocks, ClocksCmnts, ChargeBatteries, BatteriesCmnts, ChargeGoPro, GoProCmnts, ChargeCamera, "
    sql = sql & "CameraCmnts, ResolveErrors, ErrorCmnts, UploadPix, UploadPixCmnts, PixNotif, PixNotifCmnts, SendInvoice, InvoiceCmnts, "
    sql = sql & "EventPromo, PromoCmnts, UpdateFinances, FinanceCmnts, BibList, BibListCmnts, PacketPrep, PacketPrepCmnts, UpdateSeries, UpdateSeriesCmnts "
    sql = sql & "FROM EventTimeline WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        For i = 0 To 45
            If Not rs(i).Value & "" = "" Then Timeline(i) = Replace(rs(i).Value, "''", "'")
        Next
    Else
        bFound = False
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO EventTimeline(EventID) VALUES (" & lEventID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Event TimeLine</title>
<script>
$(function() {
    $( "#contract_sent" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#deposit_received" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#data_collected" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#bibs_prepped" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#event_email" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#staff_email" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#part_data" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#bib_labels" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#bib_list" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#pre_race_parts" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#charge_timers" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#charge_clocks" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#charge_go-pro" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#charge_camera" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#charge_batteries" ).datepicker({
      autoclose: true
    });
}); 

$(function() {
    $( "#resolve_errors" ).datepicker({
      autoclose: true
    });
}); 
    
$(function() {
    $( "#upload_pix" ).datepicker({
      autoclose: true
    });
}); 
    
$(function() {
    $( "#pix_notif" ).datepicker({
      autoclose: true
    });
}); 
    
$(function() {
    $( "#send_invoice" ).datepicker({
      autoclose: true
    });
}); 
    
$(function() {
    $( "#event_promo" ).datepicker({
      autoclose: true
    });
}); 
    
$(function() {
    $( "#update_finances" ).datepicker({
      autoclose: true
    });
}); 
    
$(function() {
    $( "#packet_prep" ).datepicker({
      autoclose: true
    });
}); 
    
$(function() {
    $( "#update_series" ).datepicker({
      autoclose: true
    });
}); 
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h1 class="h1">Event TimeLine</h1>
			
			<div style="margin:10px;">
				<form class="form-inline" role="form" name="which_event" method="post" action="event_timeline.asp?event_id=<%=lEventID%>">
                <div class="form-group">
				    <label for="events">Select Event:</label>
				    <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
					    <option value="">&nbsp;</option>
					    <%For i = 0 to UBound(Events, 2) - 1%>
						    <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
							    <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
						    <%Else%>
							    <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
						    <%End If%>
					    <%Next%>
				    </select>
				    <input type="hidden" class="form-control" name="submit_event" id="submit_event" value="submit_event">
				    <input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
                </div>
				</form>
			</div>
			
			<%If Not Clng(lEventID) = 0 Then%>	
                <form class="form" role="form" name="update_settings" method="post" action="event_timeline.asp?event_id=<%=lEventID%>">

                <div class="control-group">
                    <input type="hidden" class="form-control" name="submit_settings" id="submit_settings" value="submit_settings">
                    <input type="submit" class="form-control" name="submit1" id="submit1" value="Save Settings">
                </div>

                <h3 class="h3">Phase 1: Advance</h3>

                <table class="table table-striped">
                    <tr><th>Item</th><th>Target</th><th>Completed</th><th>Comments</th></tr>
                    <tr>
                        <td>Send Contract</td>
                        <td>ASAP</td>
                        <td><input type="text" class="form-control" name="contract_sent" id="contract_sent" value="<%=Timeline(0)%>"></td>
                        <td><input type="text" class="form-control" name="contract_cmnts" id="contract_cmnts" value="<%=Timeline(1)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Receive Deposit</td>
                        <td>ASAP</td>
                        <td><input type="text" class="form-control" name="deposit_received" id="deposit_received" value="<%=Timeline(2)%>"></td>
                        <td><input type="text" class="form-control" name="deposit_cmnts" id="deposit_cmnts" value="<%=Timeline(3)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Collect Data</td>
                        <td>ASAP</td>
                        <td><input type="text" class="form-control" name="data_collected" id="data_collected" value="<%=Timeline(4)%>"></td>
                        <td><input type="text" class="form-control" name="data_cmnts" id="data_cmnts" value="<%=Timeline(5)%>" size="50"></td>
                    </tr>
                </table>

                <h3 class="h3">Phase 2: Pre-Race</h3>

                <table class="table table-striped">
                    <tr><th>Item</th><th>Target</th><th>Completed</th><th>Comments</th></tr>
                    <tr>
                        <td>Prep Bibs</td>
                        <td><%=CDate(dEventDate) - 14%></td>
                        <td><input type="text" class="form-control" name="bibs_prepped" id="bibs_prepped" value="<%=Timeline(6)%>"></td>
                        <td><input type="text" class="form-control" name="bib_cmnts" id="bib_cmnts" value="<%=Timeline(7)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Send Pre-Race Email-Event</td>
                        <td><%=CDate(dEventDate) - 14%></td>
                        <td><input type="text" class="form-control" name="event_email" id="event_email" value="<%=Timeline(8)%>"></td>
                        <td><input type="text" class="form-control" name="event_email_cmnts" id="event_email_cmnts" value="<%=Timeline(9)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Send Pre-Race Email-Staff</td>
                        <td><%=CDate(dEventDate) - 14%></td>
                        <td><input type="text" class="form-control" name="staff_email" id="staff_email" value="<%=Timeline(10)%>"></td>
                        <td><input type="text" class="form-control" name="staff_email_cmnts" id="staff_email_cmnts" value="<%=Timeline(11)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Event Promo</td>
                        <td><%=CDate(dEventDate) - 10%></td>
                        <td><input type="text" class="form-control" name="event_promo" id="event_promo" value="<%=Timeline(36)%>"></td>
                        <td><input type="text" class="form-control" name="promo_cmnts" id="promo_cmnts" value="<%=Timeline(37)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Get Part Data</td>
                        <td><%=CDate(dEventDate) - 7%></td>
                        <td><input type="text" class="form-control" name="part_data" id="part_data" value="<%=Timeline(12)%>"></td>
                        <td><input type="text" class="form-control" name="part_data_cmnts" id="part_data_cmnts" value="<%=Timeline(13)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Packet Prep</td>
                        <td><%=CDate(dEventDate) - 2%></td>
                        <td><input type="text" class="form-control" name="packet_prep" id="packet_prep" value="<%=Timeline(42)%>"></td>
                        <td><input type="text" class="form-control" name="packet_prep_cmnts" id="packet_prep_cmnts" value="<%=Timeline(43)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Prep Bib Labels</td>
                        <td><%=CDate(dEventDate) - 2%></td>
                        <td><input type="text" class="form-control" name="bib_labels" id="bib_labels" value="<%=Timeline(14)%>"></td>
                        <td><input type="text" class="form-control" name="bib_label_cmnts" id="bib_label_cmnts" value="<%=Timeline(15)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Send Bib List</td>
                        <td><%=CDate(dEventDate) - 2%></td>
                        <td><input type="text" class="form-control" name="bib_list" id="bib_list" value="<%=Timeline(40)%>"></td>
                        <td><input type="text" class="form-control" name="bib_list_cmnts" id="bib_list_cmnts" value="<%=Timeline(41)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Send Pre-Race-Parts</td>
                        <td><%=CDate(dEventDate) - 1%></td>
                        <td><input type="text" class="form-control" name="pre_race_parts" id="pre_race_parts" value="<%=Timeline(16)%>"></td>
                        <td><input type="text" class="form-control" name="pre_race_cmnts" id="pre_race_cmnts" value="<%=Timeline(17)%>" size="50"></td>
                    </tr>
                </table>

                <h3 class="h3">Phase 3: Equipment</h3>

                <table class="table table-striped">
                    <tr><th>Item</th><th>Target</th><th>Completed</th><th>Comments</th></tr>
                    <tr>
                        <td>Charge Timers</td>
                        <td><%=CDate(dEventDate) - 1%></td>
                        <td><input type="text" class="form-control" name="charge_timers" id="charge_timers" value="<%=Timeline(18)%>"></td>
                        <td><input type="text" class="form-control" name="timer_cmnts" id="timer_cmnts" value="<%=Timeline(19)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Charge Clocks</td>
                        <td><%=CDate(dEventDate) - 1%></td>
                        <td><input type="text" class="form-control" name="charge_clocks" id="charge_clocks" value="<%=Timeline(20)%>"></td>
                        <td><input type="text" class="form-control" name="clocks_cmnts" id="clocks_cmnts" value="<%=Timeline(21)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Charge Batteries</td>
                        <td><%=CDate(dEventDate) - 1%></td>
                        <td><input type="text" class="form-control" name="charge_batteries" id="charge_batteries" value="<%=Timeline(22)%>"></td>
                        <td><input type="text" class="form-control" name="batteries_cmnts" id="batteries_cmnts" value="<%=Timeline(23)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Charge Go-Pro</td>
                        <td><%=CDate(dEventDate) - 1%></td>
                        <td><input type="text" class="form-control" name="charge_go-pro" id="charge_go-pro" value="<%=Timeline(24)%>"></td>
                        <td><input type="text" class="form-control" name="go-pro_cmnts" id="go-pro_cmnts" value="<%=Timeline(25)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Charge Camera</td>
                        <td><%=CDate(dEventDate) - 1%></td>
                        <td><input type="text" class="form-control" name="charge_camera" id="charge_camera" value="<%=Timeline(26)%>"></td>
                        <td><input type="text" class="form-control" name="camera_cmnts" id="camera_cmnts" value="<%=Timeline(27)%>" size="50"></td>
                    </tr>
                </table>

                <h3 class="h3">Phase 4: Post-Race</h3>

                <table class="table table-striped">
                    <tr><th>Item</th><th>Target</th><th>Completed</th><th>Comments</th></tr>
                    <tr>
                        <td>Resolve Errors</td>
                        <td><%=dEventDate%></td>
                        <td><input type="text" class="form-control" name="resolve_errors" id="resolve_errors" value="<%=Timeline(28)%>"></td>
                        <td><input type="text" class="form-control" name="error_cmnts" id="error_cmnts" value="<%=Timeline(29)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Prep/Upload Pix</td>
                        <td><%=dEventDate%></td>
                        <td><input type="text" class="form-control" name="upload_pix" id="upload_pix" value="<%=Timeline(30)%>"></td>
                        <td><input type="text" class="form-control" name="upload_pix_cmnts" id="upload_pix_cmnts" value="<%=Timeline(31)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Send Pix Notification</td>
                        <td><%=dEventDate%></td>
                        <td><input type="text" class="form-control" name="pix_notif" id="pix_notif" value="<%=Timeline(32)%>"></td>
                        <td><input type="text" class="form-control" name="pix_notif_cmnts" id="pix_notif_cmnts" value="<%=Timeline(33)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Update Series</td>
                        <td><%=dEventDate%></td>
                        <td><input type="text" class="form-control" name="update_series" id="update_series" value="<%=Timeline(44)%>"></td>
                        <td><input type="text" class="form-control" name="update_series_cmnts" id="update_series_cmnts" value="<%=Timeline(45)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Send Invoice</td>
                        <td><%=dEventDate%></td>
                        <td><input type="text" class="form-control" name="send_invoice" id="send_invoice" value="<%=Timeline(34)%>"></td>
                        <td><input type="text" class="form-control" name="invoice_cmnts" id="invoice_cmnts" value="<%=Timeline(35)%>" size="50"></td>
                    </tr>
                    <tr>
                        <td>Update Finances</td>
                        <td><%=dEventDate%></td>
                        <td><input type="text" class="form-control" name="update_finances" id="update_finances" value="<%=Timeline(38)%>"></td>
                        <td><input type="text" class="form-control" name="finance_cmnts" id="finance_cmnts" value="<%=Timeline(39)%>" size="50"></td>
                    </tr>
                </table>
                </form>
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
%>