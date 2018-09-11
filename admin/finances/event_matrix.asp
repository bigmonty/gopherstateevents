<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i, j, k
Dim iYear, iNumParts
Dim lngTotalParts, lFinanceRcd
Dim sngMileage, sngInvoice, sngMileageTotal, sngInvoiceTotal, sngTotalCost, sngStaffing, sngPacketPickup, sngAnnouncer, sngExtraBoxFee, sngDigitalDisplay
Dim sngEventIncome, sngEventCost, sngStaffingTotal, sngAnnouncerTotal, sngPacketPickupTotal, sngExtraBoxFeeTotal, sngDigitalDisplayTotal, sngIncomeTotal
Dim sngCostTotal, sngMiscCost, sngMiscCostTotal, sngLaborCost, sngPartCost, sngLaborCostTotal, sngPartCostTotal, sngEventCostTotal, sngFeaturedEvnt
Dim sngPixSponsor, sngFeaturedEvntTotal, sngPixSponsorTotal
Dim Events(), SortArr(3), Items()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = REquest.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(Date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Items(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FinanceItemsID, ItemName, ItemType, UnitCost FROM FinanceItems WHERE Active = 'y' ORDER BY ItemType, ItemName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Items(0, i) = rs(0).Value
    Items(1, i) = Replace(rs(1).Value, "''", "'")
    Items(2, i) = rs(2).Value
    Items(3, i) = rs(3).Value
    i = i + 1 
    ReDim Preserve Items(3, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

i = 0
ReDim Events(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '1/1/" & iYear & "' AND EventDate <= '" & Date & "' ORDER BY EventDate"
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
sql = "SELECT MeetsID, MeetName, MeetDate, Sport FROM Meets WHERE MeetDate >= '1/1/" & iYear & "' AND MeetDate <= '" & Date & "' ORDER BY MeetDate"
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

sngMileageTotal = 0
sngInvoiceTotal = 0

sngStaffingTotal = 0
sngAnnouncerTotal = 0
sngPixSponsorTotal = 0
sngFeaturedEvntTotal = 0
sngPacketPickupTotal = 0
sngDigitalDisplayTotal = 0
sngExtraBoxFeeTotal = 0
sngMiscCostTotal = 0
sngPartCostTotal = 0
sngLaborCostTotal = 0
sngEventCostTotal = 0

lngTotalParts = 0

Private Sub EventData(lThisEvent, sThisSport)
    sngMileage = 0
    sngInvoice = 0
    sngEventIncome = 0
    sngEventCost = 0
    sngStaffing = 0
    sngAnnouncer = 0
    sngPacketPickup = 0
    sngDigitalDisplay = 0
    sngExtraBoxFee = 0
    sngPixSponsor = 0
    sngFeaturedEvnt = 0
    sngMiscCost = 0
    sngPartCost = 0
    sngLaborCost = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Mileage, Invoice, Staffing, Announcer, PacketPickup, DigitalDisplay, ExtraBoxFee, MiscCost, PartCost, LaborCost, PixSponsor, FeaturedEvnt "
    sql = sql & "FROM FinanceEvents WHERE EventID = " & lThisEvent & " AND Sport = '" & sThisSport & "'"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sngMileage = rs(0).Value
        sngInvoice = rs(1).Value
        sngStaffing = rs(2).Value
        sngAnnouncer = rs(3).Value
        sngPacketPickup = rs(4).Value
        sngDigitalDisplay = rs(5).Value
        sngExtraBoxFee = rs(6).Value
        sngPixSponsor = rs(10).Value
        sngFeaturedEvnt = rs(11).Value
        sngMiscCost = rs(7).Value
        sngPartCost = rs(8).Value
        sngLaborCost = rs(9).Value
    End If
    rs.Close
    Set rs = Nothing

    sngStaffingTotal = CSng(sngStaffingTotal) + CSng(sngStaffing)
    sngMiscCostTotal = CSng(sngMiscCostTotal) + CSng(sngMiscCost)
    sngPartCostTotal = CSng(sngPartCostTotal) + CSng(sngPartCost)
    sngLaborCostTotal = CSng(sngLaborCostTotal) + CSng(sngLaborCost)

    sngEventIncome = CSng(sngMileage) + CSng(sngInvoice) + CSng(sngAnnouncer) + CSng(sngPacketPickup) + CSng(sngDigitalDisplay) + CSng(sngExtraBoxFee) + CSng(sngFeaturedEvnt) + CSng(sngPixSponsor)
    sngEventCost = CSng(sngStaffing) + CSNg(sngMiscCost) + CSNg(sngPartCost) + CSNg(sngLaborCost) + CSng(sngMileage)

    sngPixSponsorTotal = CSng(sngPixSponsorTotal) + CSng(sngPixSponsor)
    sngFeaturedEvntTotal = CSng(sngFeaturedEvntTotal) + CSng(sngFeaturedEvnt)
    sngAnnouncerTotal = CSng(sngAnnouncerTotal) + CSng(sngAnnouncer)
    sngPacketPickupTotal = CSng(sngPacketPickupTotal) + CSng(sngPacketPickup)
    sngDigitalDisplayTotal = CSng(sngDigitalDisplayTotal) + CSng(sngDigitalDisplay)
    sngExtraBoxFeeTotal = CSng(sngExtraBoxFeeTotal) + CSng(sngExtraBoxFee)
    sngMileageTotal = CSng(sngMileageTotal) + CSng(sngMileage)
    sngInvoiceTotal = CSng(sngInvoiceTotal) + CSng(sngInvoice)

    sngCostTotal = CSng(sngStaffingTotal) + CSng(sngMiscCostTotal) + CSng(sngPartCostTotal) + CSng(sngLaborCostTotal)

    'get num parts
    iNumParts = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT fr.NumParts FROM FinanceRaces fr INNER JOIN FinanceEvents fe ON fr.FinanceEventsID = fe.FinanceEventsID WHERE " 
    sql = sql & " fe.EventID = " & lThisEvent & " AND fe.Sport = '" & sThisSport & "'"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        iNumParts = CInt(iNumParts) + CInt(rs(0).Value)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    lngTotalParts = CLng(lngTotalParts) + CInt(iNumParts)
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Finances: Events Matrix</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
            <!--#include file = "events_nav.asp" -->

		    <h3 class="h3">GSE Finances: Events Matrix</h3>

            <ul class="nav">
                <%For i = 2015 To Year(Date) + 1%>
                    <li class="nav-item"><a class="nav-link" href="event_matrix.asp?year=<%=i%>"><%=i%></a></li>
                <%Next%>
           </ul>

            <div class="table-responsive">
                <table class="table table-condensed table-sm">
                    <tr>
                        <th rowspan="2">No.</th>
                        <th rowspan="2">Event/Meet (Date)</th>
                        <th rowspan="2">Num Parts</th>
                        <th class="bg-success" colspan="7">Income</th>
                        <th class="bg-info" colspan="6">Cost</th>
                        <th class="bg-danger" rowspan="2">Profit</th>
                        <th class="bg-danger" rowspan="2">Margin</th>
                    </tr>
                    <tr>
                        <th class="bg-success" >Invce</th>
                        <th class="bg-success" >Anncr</th>
                        <th class="bg-success" >Pkt PckUp</th>
                        <th class="bg-success" >Dig Disp</th>
                        <th class="bg-success" >Xtra Box</th>
                        <th class="bg-success" >Pix Spnsr</th>
                        <th class="bg-success" >Ftrd Evnt</th>
                        <th class="bg-info">Mlg</th>
                        <th class="bg-info">Staff</th>
                        <th class="bg-info">Partic</th>
                        <th class="bg-info">Labor</th>
                        <th class="bg-info">Misc</th>
                        <th class="bg-info">Total</th>
                    </tr>
                    <%For j = 0 To UBound(Events, 2) - 1%>
                        <%Call EventData(Events(0, j), Events(3, j))%>
                        <tr>
                            <td><%=j + 1%></td>
                            <td><%=Events(1, j)%> (<%=Events(2, j)%>)</td>
                            <td><%=iNumParts%></td>
                            <td class="bg-success">$<%=sngInvoice%></td>
                            <td class="bg-success">$<%=sngAnnouncer%></td>
                            <td class="bg-success">$<%=sngPacketPickup%></td>
                            <td class="bg-success">$<%=sngDigitalDisplay%></td>
                            <td class="bg-success">$<%=sngExtraBoxFee%></td>
                            <td class="bg-success">$<%=sngPixSponsor%></td>
                            <td class="bg-success">$<%=sngFeaturedEvnt%></td>
                            <td class="bg-info">$<%=sngMileage%></td>
                            <td class="bg-info">$<%=sngStaffing%></td>
                            <td class="bg-info">$<%=sngPartCost%></td>
                            <td class="bg-info">$<%=sngLaborCost%></td>
                            <td class="bg-info">$<%=sngMiscCost%></td>
                            <th class="bg-info">$<%=sngEventCost%></th>
                            <th class="bg-danger">$<%=Round(CSng(sngInvoice) - CSng(sngEventCost), 2)%></th>
                            <%If CSng(sngInvoice) > 0 Then%>
                                <th class="bg-danger"><%=Round(((CSng(sngInvoice) - CSng(sngEventCost))/CSng(sngInvoice))*100, 2)%>%</th>
                            <%Else%>
                                <th class="bg-danger">&nbsp;</th>
                            <%End If%>
                        </tr>
                    <%Next%>
                    <tr>
                        <th colspan="2">Column Totals</th>
                        <th><%=lngTotalParts%></th>
                        <th class="bg-success">$<%=sngInvoiceTotal%></th>
                        <th class="bg-success">$<%=sngAnnouncerTotal%></th>
                        <th class="bg-success">$<%=sngPacketPickupTotal%></th>
                        <th class="bg-success">$<%=sngDigitalDisplayTotal%></th>
                        <th class="bg-success">$<%=sngExtraBoxFeeTotal%></th>
                        <th class="bg-success">$<%=sngPixSponsorTotal%></th>
                        <th class="bg-success">$<%=sngFeaturedEvntTotal%></th>
                        <th class="bg-info">$<%=sngMileageTotal%></th>
                        <th class="bg-info">$<%=sngStaffingTotal%></th>
                        <th class="bg-info">$<%=sngPartCostTotal%></th>
                        <th class="bg-info">$<%=sngLaborCostTotal%></th>
                        <th class="bg-info">$<%=sngMiscCostTotal%></th>
                        <th class="bg-info">$<%=sngCostTotal%></th>
                        <th class="bg-danger">$<%=Round(CSng(sngInvoiceTotal) - CSng(sngCostTotal), 2)%></th>
                        <%If CSng(sngInvoiceTotal) > 0 Then%>
                            <th class="bg-danger"><%=Round(((CSng(sngInvoiceTotal) - CSng(sngCostTotal))/CSng(sngInvoiceTotal))*100, 2)%>%</th>
                        <%Else%>
                            <th class="bg-danger">&nbsp;</th>
                        <%End If%>
                    </tr>
                </table>
            </div>
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
