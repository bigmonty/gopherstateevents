<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim lEventID, lRaceID, lFinanceEventsID, lFinanceRacesID
Dim iYear, iNumParts
Dim sSport, sEventName, sComments
Dim sngMileage, sngInvoice, sngPacketPickup, sngAnnouncer, sngExtraBoxFee, sngDigitalDisplay, sngFeaturedEvnt
Dim sngMiscCost, sngStaffing, sngPartCost, sngLaborCost, sngEventCost, sngEventItemCost, sngEventItemTotal
Dim sngScalar
Dim Events(), SortArr(3), Races(), Items()
Dim dEventDate
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
sSport = Request.QueryString("sport")

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

lRaceID = Request.QueryString("race_id")
lFinanceEventsID = Request.QueryString("finance_events_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Items(4, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FinanceItemsID, ItemName, ItemType, UnitCost, Comments FROM FinanceItems WHERE Active = 'y' ORDER BY ItemType, ItemName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Items(0, i) = rs(0).Value
    Items(1, i) = Replace(rs(1).Value, "''", "'")
    Items(2, i) = rs(2).Value
    Items(3, i) = rs(3).Value
    Items(4, i) = rs(4).Value
    i = i + 1 
    ReDim Preserve Items(4, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

i = 0
ReDim Events(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '1/1/" & iYear & "' AND EventDate <= '12/31/" & iYear & "' ORDER BY EventDate"
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
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '12/31/" & iYear & "' ORDER BY MeetDate"
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

If Request.Form.Item("submit_race_items") = "submit_race_items" Then
    'get finance race id
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceRacesID FROM FinanceRaces WHERE FinanceEventsID = " & lFinanceEventsID & " AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    lFinanceRacesID = rs(0).Value
    rs.Close
    Set rs = Nothing

    'delete all
    sql = "DELETE FROM FinanceRaceItems WHERE FinanceRacesID = " & lFinanceRacesID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    're-insert new values
    For i = 0 To UBound(Items, 2) - 1
        sngScalar = Request.Form.Item("items_" & Items(0, i))

        If sngScalar = vbNullString Then sngScalar = 0
        
        If CSng(sngScalar) > 0 Then
            sql = "INSERT INTO FinanceRaceItems (FinanceRacesID, ItemID, Scalar) VALUES (" & lFinanceRacesID & ", " & Items(0, i) & ", " & sngScalar & ")"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    Next

    'reset finance events labor and part costs
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT LaborCost, PartCost FROM FinanceEvents WHERE FinanceEventsID = " & lFinanceEventsID
    rs.Open sql, conn, 1, 2
    rs(0).Value = 0
    rs(1).Value = 0
    rs.Update
    rs.Close
    Set rs = Nothing

    Call GetRaces(lEventID, sSport)

    For i = 0 To UBound(Items, 2) - 1
        For j = 0 To UBound(Races, 2) - 1
            'update event table
            Call GetItemData(Items(0, i), Races(0, j))

            If Items(2, i) = "Labor Costs" Then
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT LaborCost FROM FinanceEvents WHERE FinanceEventsID = " & lFinanceEventsID
                rs.Open sql, conn, 1, 2
                rs(0).Value = CSng(rs(0).Value) + CSng(sngEventItemCost)
                rs.Update
                rs.Close
                Set rs = Nothing
            Else
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT PartCost FROM FinanceEvents WHERE FinanceEventsID = " & lFinanceEventsID
                rs.Open sql, conn, 1, 2
                rs(0).Value = CSng(rs(0).Value) + CSng(sngEventItemCost)
                rs.Update
                rs.Close
                Set rs = Nothing
            End If
        Next
    Next
ElseIf Request.Form.Item("submit_num_parts") = "submit_num_parts" Then
    iNumParts = Request.Form.Item("num_parts")
    If CStr(iNumParts) = vbNullString Then iNumParts = 0

    'get financeeventsid
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceEventsID FROM FinanceEvents WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    lFinanceEventsID = rs(0).Value
    rs.Close
    Set rs = Nothing

    'enter num parts
    bFound = False
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT NumParts FROM FinanceRaces WHERE FinanceEventsID = " & lFinanceEventsID & " AND RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs(0).Value = iNumParts
        rs.Update
        bFound = True
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO FinanceRaces(FinanceEventsID, RaceID, NumParts) VALUES (" & lFinanceEventsID & ", " & lRaceID & ", " & iNumParts & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
    lRaceID = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_event_data") = "submit_event_data" Then
    sngMileage = Request.Form.Item("mileage")
    sngInvoice = Request.Form.Item("invoice")
    sngStaffing = Request.Form.Item("staffing")
    sngMiscCost = Request.Form.Item("misc_cost")
    sngAnnouncer = Request.Form.Item("announcer")
    sngExtraBoxFee = Request.Form.Item("extra_box_fee")
    sngDigitalDisplay = Request.Form.Item("digital_display")
    sngFeaturedEvnt = Request.Form.Item("featured_evnt")

    If CStr(sngMileage) = vbNullString Then sngMileage = 0
    If CStr(sngInvoice) = vbNullString Then sngInvoice = 0
    If CStr(sngStaffing) = vbNullString Then sngStaffing = 0
    If CStr(sngMiscCost) = vbNullString Then sngMiscCost = 0
    If CStr(sngAnnouncer) = vbNullString Then sngAnnouncer = 0
    If CStr(sngExtraBoxFee) = vbNullString Then sngExtraBoxFee = 0
    If CStr(sngDigitalDisplay) = vbNullString Then sngDigitalDisplay = 0
    If CStr(sngFeaturedEvnt) = vbNullString Then sngFeaturedEvnt = 0

    sComments = Request.Form.Item("comments")
    If Not sComments = vbNullString Then sComments = Replace(sComments, "'", "''")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice, Mileage, Staffing, MiscCost, Announcer, ExtraBoxFee, DigitalDisplay, FeaturedEvnt, Comments FROM FinanceEvents WHERE EventID = " 
    sql = sql & lEventID & " AND Sport = '" & sSport & "'"
    rs.Open sql, conn, 1, 2
    rs(0).Value = sngInvoice
    rs(1).Value = sngMileage
    rs(2).Value = sngStaffing
    rs(3).Value = sngMiscCost
    rs(4).Value = sngAnnouncer
    rs(5).Value = sngExtraBoxFee
    rs(6).Value = sngDigitalDisplay
    rs(7).Value = sngFeaturedEvnt
    rs(8).Value = sComments
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
    For i = 1 To Len(Request.Form.Item("events"))
        If Mid(Request.Form.Item("events"), i, 1) = "_" Then
            Exit For
        Else
            lEventID = lEventID & Mid(Request.Form.Item("events"), i, 1)
        End If
    Next

    If Right(Request.Form.Item("events"), 13) = "Fitness Event" Then 
        sSport = "Fitness Event"
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Sport FROM Meets WHERE MeetsID = " & lEventID
        rs.Open sql, conn2, 1, 2
        sSport = rs(0).Value
        rs.Close
        Set rs = Nothing
    End If

    If Not CStr(lEventID) = vbNullString Then
        'get financeeventsid
        lFinanceEventsID = 0

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT FinanceEventsID FROM FinanceEvents WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then lFinanceEventsID = rs(0).Value
        rs.Close
        Set rs = Nothing

        If CLng(lFinanceEventsID) = 0 Then
            sql = "INSERT INTO FinanceEvents (EventID, Sport) VALUES (" & lEventID & ", '" & sSport & "')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT FinanceEventsID FROM FinanceEvents WHERE EventID = " & lEventID & " AND Sport = '" & sSport & "'"
            rs.Open sql, conn, 1, 2
            lFinanceEventsID = rs(0).Value
            rs.Close
            Set rs = Nothing
        End If
    End If
End If

If CStr(lEventID) = vbNullString Then lEventID = "0"
If CStr(lRaceID) = vbNullString Then lRaceID = "0"

ReDim Races(1, 0)

sngEventCost = 0

If Not CLng(lEventID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Invoice, Mileage, Staffing, Announcer, ExtraBoxFee, PacketPickup, DigitalDisplay, MiscCost, Comments, PartCost, LaborCost, FeaturedEvnt "
    sql = sql & "FROM FinanceEvents WHERE FinanceEventsID = " & lFinanceEventsID
    rs.Open sql, conn, 1, 2
    sngInvoice = rs(0).Value
    sngMileage = rs(1).Value
    sngStaffing = rs(2).Value
    sngAnnouncer = rs(3).Value
    sngExtraBoxFee = rs(4).Value
    sngPacketPickup = rs(5).Value
    sngDigitalDisplay = rs(6).Value
    sngMiscCost = rs(7).Value
    sngPartCost = rs(9).Value
    sngLaborCost = rs(10).Value
    sComments = rs(8).Value
    sngFeaturedEvnt = rs(11).Value
    rs.Close
    Set rs = Nothing

    If sSport = "Fitness Event" Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
        rs.Open sql, conn, 1, 2
        sEventName = Replace(rs(0).Value, "''", "'")
        dEventDate = rs(1).Value
        rs.Close
        Set rs = Nothing

        Call GetRaces(lEventID, "Fitness Event")
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lEventID
        rs.Open sql, conn2, 1, 2
        sEventName = Replace(rs(0).Value, "''", "'")
        dEventDate = rs(1).Value
        rs.Close
        Set rs = Nothing

        Call GetRaces(lEventID, sSport)
    End If

    sngEventCost = CSng(sngEventCost) + CSng(sngMiscCost) + CSng(sngPartCost) + CSng(sngLaborCost) + CSng(sngStaffing) + CSng(sngMileage)

    If Not CLng(lRaceID) = 0 Then
        'enter num parts
        iNumParts = 0
        bFound = False
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT NumParts FROM FinanceRaces WHERE FinanceEventsID = " & lFinanceEventsID & " AND RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            iNumParts = rs(0).Value
            bFound = True
        End If
        rs.Close
        Set rs = Nothing

        If iNumParts = 0 Then
            If sSport = "Fitness Event" Then
                'get from part reg
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT PartRegID FROM PartReg WHERE RaceID = " & lRaceID
                rs.Open sql, conn, 1, 2
                If rs.RecordCount > 0 Then iNumParts = rs.RecordCount
                rs.Close
                Set rs = Nothing
            Else
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT IndRsltsID FROM IndRslts WHERE RacesID = " & lRaceID
                rs.Open sql, conn2, 1, 2
                If rs.RecordCount > 0 Then iNumParts = rs.RecordCount
                rs.Close
                Set rs = Nothing
            End If
        End If

        If bFound = False Then
            sql = "INSERT INTO FinanceRaces(FinanceEventsID, RaceID, NumParts) VALUES (" & lFinanceEventsID & ", " & lRaceID & ", " & iNumParts & ")"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    End If
End If

Private Sub GetRaces(lThisEvent, sThisSport)
    Dim x

    x = 0

    ReDim Races(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    If sThisSport = "Fitness Event" Then
        sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lThisEvent
        rs.Open sql, conn, 1, 2
    Else
        sql = "SELECT RacesID, RaceName FROM Races WHERE MeetsID = " & lThisEvent
        rs.Open sql, conn2, 1, 2
    End If
    Do While Not rs.EOF
        Races(0, x) = rs(0).Value
        Races(1, x) = Replace(rs(1).Value, "''", "'")
        x = x + 1
        ReDim Preserve Races(1, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If UBound(Races, 2) = 1 Then lRaceID = Races(0, 0)
End Sub

sngEventItemTotal = 0

Private Sub GetItemData(lThisItem, lThisRaceID)
    Dim sngUnitCost
    Dim iThisNumParts
   
    sngScalar = 0
    sngEventItemCost = 0

    'get num parts and finance races id
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceRacesID, NumParts FROM FinanceRaces WHERE RaceID = " & lThisRaceID & " AND FinanceEventsID = " & lFinanceEventsID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then  
        lFinanceRacesID = rs(0).Value
        iThisNumParts = rs(1).Value
    End If
    rs.Close
    Set rs = Nothing

    'get item scalar
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Scalar FROM FinanceRaceItems WHERE FinanceRacesID = " & lFinanceRacesID & " AND ItemID = " & lThisItem
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sngScalar = rs(0).Value
    rs.Close
    Set rs = Nothing

    'get unit cost
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UnitCost FROM FinanceItems WHERE FinanceItemsID = " & lThisItem
    rs.Open sql, conn, 1, 2
    sngUnitCost = rs(0).Value
    rs.Close
    Set rs = Nothing

    'calculate event cost
    sngEventItemCost = Round(CSng(sngScalar) * CSng(sngUnitCost)*CInt(iThisNumParts), 2)
    sngEventItemTotal = CSng(sngEventItemTotal) + CSng(sngEventItemCost)
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Manage Events</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <!--#include file = "events_nav.asp" -->

		    <h3 class="h3">GSE Finances: Manage Events</h3>

            <ul class="nav">
                <%For i = 2015 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="events.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
           </ul>

            <form class="form-inline" name="get_event" method="post" action="events.asp?year=<%=iYear%>">
            <label for="events">Select Event To Manage</label>
            <select class="form-control" name="events" id="events" onchange="this.form.submit1.click();">
                <option value=""></option>
                <%For i = 0 To UBound(Events, 2) - 1%>
                    <%If CLng(Events(0, i)) = CLng(lEventID) Then%>
                        <option value="<%=Events(0, i)%>_<%=Events(3, i)%>" selected><%=Events(1, i)%> &nbsp;<%=Events(2, i)%>&nbsp (<%=Events(3, i)%>)</option>
                    <%Else%>
                        <option value="<%=Events(0, i)%>_<%=Events(3, i)%>"><%=Events(1, i)%> &nbsp;<%=Events(2, i)%>&nbsp (<%=Events(3, i)%>)</option>
                    <%End If%>
                <%Next%>
            </select>
            <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
            <input type="submit" class="form-control" name="submit1" id="submit1" value="Manage This Event">
            </form>

            <%If Not CLng(lEventID) = 0 Then%>
                <br>
                <ul class="nav">
                    <li class="nav-item"><span  class="nav-link">Invoice:&nbsp;$<%=sngInvoice%></span></li>
                    <li class="nav-item"><span  class="nav-link">Staffing:&nbsp;$<%=sngStaffing%></span></li>
                    <li class="nav-item"><span  class="nav-link">Misc:&nbsp;$<%=sngMiscCost%></span></li>
                    <li class="nav-item"><span  class="nav-link">Part:&nbsp;$<%=sngPartCost%></span></li>
                    <li class="nav-item"><span  class="nav-link">Labor:&nbsp;$<%=sngLaborCost%></span></li>
                    <li class="nav-item"><span  class="nav-link">Total Cost:&nbsp;$<%=sngEventCost%></span></li>
                    <li class="nav-item"><span  class="nav-link">Profit/Loss:&nbsp;$<%=Round(CSng(sngInvoice) - CSng(sngEventCost), 2)%></span></li>
                </ul>

                <div class=" row bg-success">
                    <div>
                        <h4 class="h4">Event Data:</h4>
                        <form role="form" class="form-horizontal" name="event_data" method="post" action="events.asp?event_id=<%=lEventID%>&amp;sport=<%=sSport%>&amp;year=<%=iYear%>&amp;finance_events_id=<%=lFinanceEventsID%>">
                        <div class="form-group row">
                            <label class=" col-form-label col-xs-1" for="mileage">Mileage:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="mileage" id="mileage" value="<%=sngMileage%>">
                            </div>
                            <label class=" col-form-label col-xs-1" for="invoice">Invoice:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="invoice" id="invoice" value="<%=sngInvoice%>">
                            </div>
                            <label class=" col-form-label col-xs-1" for="staffing">Staff:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="staffing" id="staffing" value="<%=sngStaffing%>">
                            </div>
                        </div>
                        <div class="form-group row">
                        <label class=" col-form-label col-xs-1" for="misc_cost">Misc:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="misc_cost" id="misc_cost" value="<%=sngMiscCost%>">
                            </div>
                            <label class=" col-form-label col-xs-1" for="announcer">Announcer:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="announcer" id="announcer" value="<%=sngAnnouncer%>">
                            </div>
                            <label class=" col-form-label col-xs-1" for="digital_display">Digital Display:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="digital_display" id="digital_display" value="<%=sngDigitalDisplay%>">
                            </div>
                        </div>
                        <div class="form-group row">
                            <label class=" col-form-label col-xs-1" for="extra_box_fee">Extra Box Fee:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="extra_box_fee" id="extra_box_fee" value="<%=sngExtraBoxFee%>">
                            </div>
                            <label class="col-form-label col-xs-1" for="featured_evnt">Featured Event:</label>
                            <div class="col-xs-3">
                                <input type="text" class="form-control form-control-sm" name="featured_evnt" id="featured_evnt" value="<%=sngFeaturedEvnt%>">
                            </div>
                            <label class=" col-form-label col-xs-1" for="comments">Comments:</label>
                            <div class="col-xs-3">
                                <textarea class="form-control form-control-sm" name="comments" id="comments" rows="2"><%=sComments%></textarea>
                            </div>
                        </div>
                        <div class="form-group">
                            <input type="hidden" name="submit_event_data" id="submit_event_data" value="submit_event_data">
                            <input type="submit" class="form-control form-control-sm" name="submit2" id="submit2" value="Submit Event Data">
                        </div>
                        </form>
                    </div>
                </div>

                <hr>

                <h4 class="h4">Race Data:</h4>

                <form class="form-inline" name="get_race" method="post" action="events.asp?event_id=<%=lEventID%>&amp;sport=<%=sSport%>&amp;year=<%=iYear%>&amp;finance_events_id=<%=lFinanceEventsID%>">
                <label for="races">Select Race:</label>
                <select class="form-control" name="races" id="races" onchange="this.form.submit3.click();">
                    <option value=""></option>
                    <%For i = 0 To UBound(Races, 2) - 1%>
                        <%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
                            <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
                        <%Else%>
                            <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
                <input type="hidden" name="submit_race" id="submit_race" value="submit_race">
                <input type="submit" class="form-control" name="submit3" id="submit3" value="Submit Race">
                </form>

                <br>

                <%If Not CLng(lRaceID) = 0 Then%>
                    <form class="form-inline" name="get_num_parts" method="post" 
                        action="events.asp?event_id=<%=lEventID%>&amp;sport=<%=sSport%>&amp;year=<%=iYear%>&amp;race_id=<%=lRaceID%>&amp;finance_events_id=<%=lFinanceEventsID%>">
                    <label for="num_parts">Num Parts:</label>
                    <input type="text" class="form-control" name="num_parts" id="num_parts" value="<%=iNumParts%>">
                    <input type="hidden" name="submit_num_parts" id="submit_num_parts" value="submit_num_parts">
                    <input type="submit" class="form-control" name="submit4" id="submit4" value="Submit Num Parts">
                    </form>

                    <%If iNumParts > 0 Then%>
                        <br>
                        <span style="font-weight: bold;">Enter Race Costs:</span>
                        <form name="get_item_cost" method="post" 
                        action="events.asp?event_id=<%=lEventID%>&amp;sport=<%=sSport%>&amp;year=<%=iYear%>&amp;race_id=<%=lRaceID%>&amp;finance_events_id=<%=lFinanceEventsID%>">
                        <table class="table table-striped">
                            <tr>
                                <th>Item</th>
                                <th>Category</th>
                                <th>Unit Cost</th>
                                <th>Comments</th>
                                <th>Scalar</th>
                                <th>Event Cost</th>
                            </tr>
                            <%For j = 0 To UBound(Items, 2) - 1%>
                                <%Call GetItemData(Items(0, j), lRaceID)%>
                                <tr>
                                    <th><%=Items(1, j)%></th>
                                    <td><%=Items(2, j)%></td>
                                    <td><%=Items(3, j)%></td>
                                    <td><%=Items(4, j)%></td>
                                    <td>
                                        <input type="text" class="form-control" name="items_<%=Items(0, j)%>" id="items_<%=Items(0, j)%>" 
                                            value="<%=sngScalar%>" style="text-align: right;">
                                    </td>
                                    <td>$<%=sngEventItemCost%></td>
                                </tr>
                            <%Next%>
                            <tr>
                                <th style="text-align: right;" colspan="5">Total Cost:</th>
                                <th style="text-align: right;">$<%=Round(sngEventItemTotal)%></th>
                            </tr>
                            <tr>
                                <td colspan="6">
                                    <input type="hidden" name="submit_race_items" id="submit_race_items" value="submit_race_items">
                                    <input type="submit" class="form-control" name="submit5" id="submit5" value="Submit Race Items">
                                </td>
                            </tr>
                        </table>
                        </form>
                    <%End If%>
                <%End If%>
            <%End If%>
        </div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%	
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
