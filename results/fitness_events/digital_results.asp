<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i, j, k
Dim lEventID, lMyPartID, lMyRaceID, lSuppLegID
Dim iBibToFind, iMinPlace, iMaxPlace
Dim sEventName, sEventRaces, sErrMsg, sMF, sSuppTime, sOtherTime, sLegName, sOtherName, sOrderBy, sSortRsltsBy
Dim sShowAge
Dim BibRslts(7), EventList, IndRslts, RsltsScroll

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

'Response.Redirect "/misc/taking_break.htm"

iBibToFind = 0
lSuppLegID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SuppLegID, LegName, OtherName FROM SuppLeg WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then 
    lSuppLegID = rs(0).Value
    sLegName = Replace(rs(1).Value, "''", "'")
    If Not rs(2).Value & "" = "" Then sOtherName = Replace(rs(2).Value, "''", "'")
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate <= '" & Date & "' ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    EventList = rs.GetRows()
Else
    ReDim EventList(2, 0)
End If
rs.Close
Set rs = Nothing

If CLng(lEventID) = 0 Then 
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID FROM Events WHERE EventDate <= '" & Date & "' ORDER BY EventDate DESC"
    rs.Open sql, conn, 1, 2
    lEventID = rs(0).Value
    rs.Close
    Set rs = Nothing
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sEventRaces = sEventRaces & rs(0).Value & ", "
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not sEventRaces = vbNullString Then sEventRaces = Left(sEventRaces, Len(sEventRaces) - 2)

sql = "SELECT ShowAge, SortRsltsBy, EventID FROM RaceData WHERE RaceID IN (" & sEventRaces & ")"
Set rs = conn.Execute(sql)
sShowAge = rs(0).Value
sSortRsltsBy = rs(1).Value
lEventID = rs(2).Value
Set rs = Nothing

If sSortRsltsBy = "EventPl" Then
    sOrderBy = "IR.EventPl"
Else
    sOrderBy = "IR.FnlScnds"
End If
							
sql = "SELECT P.Country, PR.Bib, P.FirstName, P.LastName, P.Gender, PR.Age, IR.ChipTime, IR.FnlTime, IR.ChipStart "
sql = sql & "FROM dbo.RaceData AS R INNER JOIN dbo.PartRace AS PR ON R.RaceID = PR.RaceID INNER JOIN "
sql = sql & "dbo.Participant AS P ON PR.ParticipantID = P.ParticipantID INNER JOIN "
sql = sql & "(SELECT DISTINCT RaceID, ParticipantID, ChipTime, FnlTime, ChipStart, FnlScnds, EventPl FROM dbo.IndResults "
sql = sql & "WHERE (FnlTime IS NOT NULL) AND (FnlTime > '00:00:00.000')) AS IR ON R.RaceID = IR.RaceID AND P.ParticipantID = IR.ParticipantID "
sql = sql & "WHERE (R.RaceID IN (" & sEventRaces & ")) ORDER BY " & sOrderBy   
Set rs = conn.Execute(sql)
If True = rs.BOF Then
    ReDim RsltsScroll(8, 0)
Else
    RsltsScroll = rs.GetRows()
End If
Set rs = Nothing

For i = 0 To UBound(RsltsScroll, 2)
    RsltsScroll(0, i) = i + 1
    RsltsScroll(2, i) = Replace(RsltsScroll(2, i), """", "")
    RsltsScroll(3, i) = Replace(RsltsScroll(3, i), """","")
    If RsltsScroll(5, i) = "99" Or sShowAge = "n" Then RsltsScroll(5, i) = "--"
Next

If Request.form.Item("submit_bib") = "submit_bib" Then
    iBibToFind = Request.Form.Item("bib_to_find")
    iBibToFind = Replace(CStr(iBibToFind), ".", "")

    If CStr(iBibToFind) = vbNullString Then 
        iBibToFind = 0
    ElseIf Not IsNumeric(iBibToFind) Then
        iBibToFind = 0
        sErrMsg = "Bib numbers must be numeric."
    ElseIf Len(CStr(iBibToFind)) > 4 Then
        iBibToFind = 0
        sErrMsg = "Bib numbers must be 4 characters or less."
    End If
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")

    If CStr(lEventID) = vbNullString Then lEventID = 0
    If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
    If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")
End If

If Not CInt(iBibToFind) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, RaceID, Age FROM PartRace WHERE Bib = " & iBibToFind & " AND RaceID IN (" & sEventRaces & ")"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        lMyPartID = rs(0).Value
        lMyRaceID = rs(1).Value
        If rs(2).Value="99" Then
            BibRslts(2) = "na"
        Else
            BibRslts(2) = rs(2).Value
        End If
    Else
        sErrMsg = "I'm sorry.  That bib number was not found in the results for this event."
    End If
    rs.Close
    Set rs = Nothing

    If sErrMsg = vbNullString Then
	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT LastName, FirstName, Gender FROM Participant WHERE ParticipantID = " & lMyPartID
        rs.Open sql, conn, 1, 2
        BibRslts(1) = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
        sMF = rs(2).Value
        rs.Close
        Set rs = Nothing

        k = 1
	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ParticipantID FROM IndResults WHERE RaceID = " & lMyRaceID & " AND FnlScnds > 0 ORDER BY FnlScnds"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If ThisGender(CStr(rs(0).Value)) = CStr(sMF) Then
                If CLng(rs(0).Value) = CLng(lMyPartID) Then
                    BibRslts(0) = k
                    Exit Do
                Else
                    k = k + 1
                End If
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing

	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lMyRaceID
        rs.Open sql, conn, 1, 2
        BibRslts(3) = Replace(rs(0).Value, "''", "'") 
        rs.Close
        Set rs = Nothing

        iMinPlace = 0
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ChipTime, FnlTime, ChipStart, EventPl FROM IndResults WHERE RaceID = " & lMyRaceID & " AND ParticipantID = " & lMyPartID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            BibRslts(4) = rs(0).Value
            BibRslts(5) = rs(1).Value
            BibRslts(6) = rs(2).Value
            BibRslts(7) = rs(3).Value
            iMinPlace = CInt(rs(3).Value) - 3
            iMaxPlace = CInt(rs(3).Value) + 3
        Else
            sErrMsg = "I'm sorry.  That bib number was not found."
        End If
        rs.Close
        Set rs = Nothing

        If CLng(lSuppLegID) > 0 Then
            sSuppTime = "00:00.000"
            sOtherTime = "00:00:00.000"

            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT SuppTime, OtherTime FROM SuppLegRslts WHERE SuppLegID = " & lSuppLegID & " AND Bib = " & iBib
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then 
                sSuppTime = rs(0).Value
                sOtherTime = rs(1).Value
            End If
            rs.Close
            Set rs = Nothing
        End If

        If CInt(iMinPlace) <= 0 Then 
            iMinPlace = 1
            iMaxPlace = 20
        End If
    End If

    If sErrMsg = vbNullString Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT pr.Bib, p.LastName, p.FirstName, p.Gender, pr.Age, ir.ChipTime, ir.FnlTime, ir.ChipStart, p.City, p.St "
        sql = sql & "FROM Participant p JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
        sql = sql & "JOIN PartRace pr ON pr.RaceID = ir.RaceID AND pr.ParticipantID = p.ParticipantID "
        sql = sql & "WHERE ir.RaceID IN (" & sEventRaces & ") AND ir.FnlTime IS NOT NULL AND ir.FnlTime > '00:00:00.000' AND ir.EventPl >= " & iMinPlace
        sql = sql & " AND ir.EventPl <= " & iMaxPlace & " ORDER BY ir.FnlScnds"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            IndRslts = rs.GetRows()
        Else
            ReDim IndRslts(9, 0)
        End If
        rs.Close
        Set rs = Nothing

        For i = 0 To UBound(IndRslts, 2)
            If IndRslts(4, i) ="99" Then
                IndRslts(4, i) = "na"
'            Else
'                IndRslts(4, i) = rs(4).Value
            End If
        Next
    End If
End If

Private Function ThisGender(lThisPart)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Gender FROM Participant WHERE ParticipantID = " & lThisPart
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then ThisGender = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing
End Function

Private Sub GetSplits(iBib)
    sSuppTime = "00:00.000"
    sOtherTime = "00:00:00.000"

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SuppTime, OtherTime FROM SuppLegRslts WHERE SuppLegID = " & lSuppLegID & " AND Bib = " & iBib
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then 
        sSuppTime = rs(0).Value
        sOtherTime = rs(1).Value
    End If
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE Results Kiosk For <%=sEventName%></title>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
<meta name="description" content="Scrolling Results from Gopher State Events">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.min.css">
<link rel="alternate" href="https://gopherstateevents.com" hreflang="en-us" />
<link rel="stylesheet" href="https://code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap-theme.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-submenu/3.0.1/css/bootstrap-submenu.min.css">

<script>
    function chkFlds() {
        if (document.find_bib.bib_to_find.value == '')
            {
            alert('You must submit a bib number to look for.');
            return false
            }
        else
            if (isNaN(document.find_bib.bib_to_find.value))
                {
                alert('The bib number can not contain non-numeric values');
                return false
                }
        else
            return true
    }
</script>

<style>
body {
    margin: 10px;
    font-family:'Lato', sans-serif;
    }
    
    .marquee {
        width: 600px;
        overflow: hidden;
        border:1px solid #ccc;
        background: black;
        font-size: 14px;
        height: 40px;
        color: rgb(202, 255, 195);
    }
</style>

<script>
        var data = [
        {
            "Pl": "1","Bib": "106","First Name":"Chase","Last Name": "Cayo","MFX": "M","Age": "22","Chip Time": "14:46.3","Gun Time": "14:46.9","Start Time":"0:00.6","City": "Otsego","St": "MN"
            },
            {
            "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
            },
            {
            "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
            },
            {
            "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
            }
        ];
        
        var x = 0;

        $(document).ready(function() { 
            $(".marquee").html(JSON.stringify(data[x]));
            
            $('.marquee').marquee({
              duration: 2500,
              direction: 'up'
            }).bind('finished', function(){
            		//Change text to something else after first loop finishes
            		$(this).marquee('destroy');
                if(x == data.length-1) {
                  x = 0;
                } else {
                x++;
                }
            		//Load new content using Ajax and update the marquee container
            		$(this).html(JSON.stringify(data[x]))
            			//Apply marquee plugin again
                  
            			.marquee({
                    duration: 2500,
                    direction: 'up'
                  })              
                });
        });
    </script>
</head>

<body onload="javascript:find_bib.bib_to_find.focus()">
<div class="container">
    <div class="row">
        <div class="col-sm-6">
            <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
        </div>
        <div class="col-sm-6">
            <h1 class="h1">Race Day Results <br> <%=sEventName%></h1>
        </div>
    </div>
    
    <div class="row">
        <div class="col-sm-6">
            <form role="form" class="form-inline" name="which_event" method="post" action="digital_results.asp">
            <label>Event:</label>
            <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()" style="font-size:0.9em;">
                <%For i = 0 to UBound(EventList, 2)%>
                    <%If CLng(lEventID) = CLng(EventList(0, i)) Then%>
                        <option value="<%=EventList(0, i)%>" selected><%=Replace(EventList(1, i), "''", "'")%>&nbsp;(<%=EventList(2, i)%>)</option>
                    <%Else%>
                        <option value="<%=EventList(0, i)%>"><%=Replace(EventList(1, i), "''", "'")%>&nbsp;(<%=EventList(2, i)%>)</option>
                    <%End If%>
                <%Next%>
            </select>
            <input class="form-control" type="hidden" name="submit_event" id="submit_event" value="submit_event">
            <input class="form-control" type="submit" name="get_event" id="get_event" value="Get These">
            </form>
            <br>
            <form role="form" class="form-inline" name="find_bib" method="post" action="digital_results.asp?event_id=<%=lEventID%>" onsubmit="return chkFlds;">
            <label>Bib To Find:</label>
            <input class="form-control" type="text" name="bib_to_find" id="bib_to_find" size="3" value="<%=iBibToFind%>" onfocus="this.select()">
            <input class="form-control" type="hidden" name="submit_bib" id="submit_bib" value="submit_bib">
            <input class="form-control" type="submit" name="submit_lookup" id="submit_lookup" value="Find Bib">
            </form>
            <br>
        </div>
        <div class="col-sm-6">
            <h3 class="h3" style="color:red;">Enter Your Bib Number on Keypad To View Time</h3>
        </div>
    </div>

    <%If Not CInt(iBibToFind) = 0 Then%>
        <%If sErrMsg = vbNullString Then%>
            <table class="table bg-success" style="color:#fff;">
                <tr>
                    <!--<th>Race Pl</th> not sure why this is rendering incorrectly-->
                    <th>Name</th>
                    <th>Age</th>
                    <th>Race</th>
                    <th>Chip Time</th>
                    <th>Gun Time</th>
                    <th>Chip Start</th>
                </tr>
                <tr>
                <!--<td><%=BibRslts(0)%></td>-->
                    <td><%=BibRslts(1)%></td>
                    <td><%=BibRslts(2)%></td>
                    <td><%=BibRslts(3)%></td>
                    <td><%=BibRslts(4)%></td>
                    <td><%=BibRslts(5)%></td>
                    <td><%=BibRslts(6)%></td>
                </tr>
                <%If CLng(lSuppLegID) > 0 Then%>
                    <tr>
                        <th style="text-align: right;" colspan="3"><%=sLegName%></th>
                        <td><%=sSuppTime%></td>
                        <th style="text-align: right;" colspan="3"><%=sOtherName%></th>
                        <td><%=sOtherTime%></td>
                    </tr>
                <%End If%>
            </table>
        <%Else%>
            <p style="border: none;"><%=sErrMsg%></p>
        <%End If%>
    <%Else%>
        <p style="border: none;"><%=sErrMsg%></p>
    <%End If%>

    <!--
        <%If sErrMsg = vbNullString Then%>
            <%If CInt(iBibToFind) > 0 Then%>

            <h4 class="h4">Finishers Near Bib <%=iBibToFind%></h4>
                <table class="table table-striped">
                    <tr>
                        <th>Bib-Name</th>
                        <th>M/F</th>
                        <th>Age</th>
                        <th>Chip Time</th>
                        <th>Gun Time</th>
                        <th>Start Time</th>
                        <th>From</th>
                    </tr>
                    <%For i = 0 To UBound(IndRslts, 2)%>
                        <%If CInt(IndRslts(0, i)) = CInt(iBibToFind) Then%>
                            <tr>
                                <td class="bg-success"><%=IndRslts(0, i)%> - <%=IndRslts(2, i)%>&nbsp;<%=IndRslts(1, i)%></td>
                                <td class="bg-success"><%=IndRslts(3, i)%></td>
                                <td class="bg-success"><%=IndRslts(4, i)%></td>
                                <td class="bg-success"><%=IndRslts(5, i)%></td>
                                <td class="bg-success"><%=IndRslts(6, i)%></td>
                                <td class="bg-success"><%=IndRslts(7, i)%></td>
                                <td class="bg-success"><%=IndRslts(8, i)%>, <%=IndRslts(9, i)%></td>
                            </tr>
                        <%Else%>
                            <tr>
                                <td><%=IndRslts(0, i)%> - <%=IndRslts(2, i)%>&nbsp;<%=IndRslts(1, i)%></td>
                                <td><%=IndRslts(3, i)%></td>
                                <td><%=IndRslts(4, i)%></td>
                                <td><%=IndRslts(5, i)%></td>
                                <td><%=IndRslts(6, i)%></td>
                                <td><%=IndRslts(7, i)%></td>
                                <td><%=IndRslts(8, i)%>, <%=IndRslts(9, i)%></td>
                            </tr>
                        <%End If%>
                    <%Next%>
                </table>
            <%End If%>
        <%End If%>
    -->
    <div class="marquee">
        Marquee text here...
    </div>
</div>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.4/jquery.min.js"></script> <!-- you only need ONE link to jQuery-->
<script src="https://cdn.jsdelivr.net/jquery.marquee/1.3.1/jquery.marquee.min.js"></script>
<script src="https://code.jquery.com/ui/1.11.4/jquery-ui.js"></script> <!-- not really needed as you aren't using it -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>