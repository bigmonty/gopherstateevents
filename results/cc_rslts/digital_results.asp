<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lMeetsID
Dim iBibToFind
Dim sMeetName, sMeetRaces, sErrMsg, sMF, sOrderBy, sSortRsltsBy
Dim BibRslts(5), MeetList
Dim dMeetDate
Dim bRsltsOfficial

lMeetsID = Request.QueryString("Meet_id")
If CStr(lMeetsID) = vbNullString Then lMeetsID = 0
If Not IsNumeric(lMeetsID) Then Response.Redirect("http://www.google.com")
If CLng(lMeetsID) < 0 Then Response.Redirect("http://www.google.com")

'Response.Redirect "/misc/taking_break.htm"

iBibToFind = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If CLng(lMeetsID) = 0 Then 
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetsID FROM Meets WHERE MeetDate <= '" & Date & "' ORDER BY MeetDate DESC"
    rs.Open sql, conn, 1, 2
    lMeetsID = rs(0).Value
    rs.Close
    Set rs = Nothing
End If

bRsltsOfficial = False
sql = "SELECT MeetsID FROM OfficialRslts WHERE MeetsID = " & lMeetsID
Set rs = conn.Execute(sql)
If rs.BOF and rs.EOF Then
    bRsltsOfficial = False
Else
    bRsltsOfficial = True
End If
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDate <= '" & Date & "' ORDER By MeetDate DESC"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    MeetList = rs.GetRows()
Else
    ReDim MeetList(2, 0)
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetsID
rs.Open sql, conn, 1, 2
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RacesID FROM Races WHERE MeetsID = " & lMeetsID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sMeetRaces = sMeetRaces & rs(0).Value & ", "
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not sMeetRaces = vbNullString Then sMeetRaces = Left(sMeetRaces, Len(sMeetRaces) - 2)

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
ElseIf Request.Form.Item("submit_meet") = "submit_meet" Then
	lMeetsID = Request.Form.Item("meets")

    If CStr(lMeetsID) = vbNullString Then lMeetsID = 0
    If Not IsNumeric(lMeetsID) Then Response.Redirect("http://www.google.com")
    If CLng(lMeetsID) < 0 Then Response.Redirect("http://www.google.com")
End If

If Not CInt(iBibToFind) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.Bib, r.FirstName, r.LastName, r.TeamsID, r.Gender, ir.RacesID, ir.FnlScnds "
    sql = sql & "FROM IndRslts ir INNER JOIN Roster r ON ir.RosterID = r.RosterID WHERE ir.Bib = " & iBibToFind 
    sql = sql & " AND ir.RacesID IN (" & sMeetRaces & ") AND ir.Place > 0 AND ir.FnlScnds > 0"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        BibRslts(0) = GetRacePlace(rs(0).Value, rs(5).Value)
        BibRslts(1) = rs(0).Value & "-" & Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
        BibRslts(2) = GetTeamName(rs(3).Value)
        BibRslts(3) = rs(4).Value
        BibRslts(4) = GetRaceName(rs(5).Value)
        BibRslts(5) = ConvertToMinutes(rs(6).Value)
    Else
        sErrMsg = "I'm sorry.  That bib number was not found in the results for this Meet."
    End If
    rs.Close
    Set rs = Nothing
End If

%>
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<%

Private Function GetRacePlace(iBib, lRaceID)
    GetRacePlace = 0

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Bib FROM IndRslts WHERE RacesID = " & lRaceID & " AND FnlScnds > 0 ORDER BY FnlScnds"
    rs2.OPen sql2, conn, 1, 2
    Do While Not rs2.EOF 
        GetRacePlace = CInt(GetRacePlace) + 1
        If CInt(rs2(0).Value) = CInt(iBib) Then Exit Do
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = NOthing
End Function

Private Function GetRaceName(lRaceID)
    sql2 = "SELECT RaceName FROM Races WHERE RacesID = " & lRaceID
    Set rs2 = conn.Execute(sql2)
    GetRaceName = Replace(rs2(0).Value, "''", "'") 
    Set rs2 = Nothing
End Function

Private Function GetTeamName(lTeamID)
    sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
    Set rs2 = conn.Execute(sql2)
    GetTeamName = Replace(rs2(0).Value, "''", "'") 
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE Results Kiosk For <%=sMeetName%></title>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
<meta name="description" content="Scrolling Results from Gopher State Meets">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.min.css">
<link rel="alternate" href="https://gopherstateMeets.com" hreflang="en-us" />
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
</head>

<body onload="javascript:find_bib.bib_to_find.focus()">
<div class="container">
    <div class="row">
        <div class="col-sm-6">
            <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
        </div>
        <div class="col-sm-6">
            <h1 class="h1">GSE Results By Bib <br> <%=sMeetName%></h1>
        </div>
    </div>
    
    <div class="row">
        <div class="col-sm-6">
            <form role="form" class="form-inline" name="which_Meet" method="post" action="digital_results.asp">
            <label>Meet:</label>
            <select class="form-control" name="meets" id="meets" onchange="this.form.get_meet.click()">
                <%For i = 0 to UBound(MeetList, 2)%>
                    <%If CLng(lMeetsID) = CLng(MeetList(0, i)) Then%>
                        <option value="<%=MeetList(0, i)%>" selected><%=Replace(MeetList(1, i), "''", "'")%>&nbsp;(<%=MeetList(2, i)%>)</option>
                    <%Else%>
                        <option value="<%=MeetList(0, i)%>"><%=Replace(MeetList(1, i), "''", "'")%>&nbsp;(<%=MeetList(2, i)%>)</option>
                    <%End If%>
                <%Next%>
            </select>
            <input class="form-control" type="hidden" name="submit_Meet" id="submit_Meet" value="submit_Meet">
            <input class="form-control" type="submit" name="get_Meet" id="get_Meet" value="Get These">
            </form>
            <br>
            <form role="form" class="form-inline" name="find_bib" method="post" action="digital_results.asp?Meet_id=<%=lMeetsID%>" onsubmit="return chkFlds;">
            <label>Bib To Find:</label>
            <input class="form-control" type="text" name="bib_to_find" id="bib_to_find" size="3" value="<%=iBibToFind%>" onfocus="this.select()">
            <input class="form-control" type="hidden" name="submit_bib" id="submit_bib" value="submit_bib">
            <input class="form-control" type="submit" name="submit_lookup" id="submit_lookup" value="Find Bib">
            </form>
            <br>
        </div>
        <div class="col-sm-6">
            <h3 class="h3" style="color:red;">Enter Your Bib Number To View Results</h3>
        </div>
    </div>
    
    <%If CDate(Date) < CDate(dMeetDate) + 7 Then%>
        <%If bRsltsOfficial = False Then%>
            <div class="text-danger">PRELIMINARY RESULTS...UNOFFICIAL...SUBJECT TO CHANGE.  Please report any issues to 
                bob.schneider@gopherstateevents.com.</div>
        <%Else%>
            <div class="text-success">These results are now official.  If you notice any errors please contact us 
            via <a href="mailto:bob.schneider@gopherstateevents.com">email</a> or by telephone (612-720-8427).</div>
        <%End If%>
    <%End If%>

    <%If Not CInt(iBibToFind) = 0 Then%>
        <%If sErrMsg = vbNullString Then%>
            <table class="table bg-success" style="color:#fff;">
                <tr>
                    <th>Pl</th>
                    <th>Name</th>
                    <th>School</th>
                    <th>MF</th>
                    <th>Race</th>
                    <th>Time</th>
                </tr>
                <tr>
                    <td><%=BibRslts(0)%></td>
                    <td><%=BibRslts(1)%></td>
                    <td><%=BibRslts(2)%></td>
                    <td><%=BibRslts(3)%></td>
                    <td><%=BibRslts(4)%></td>
                    <td><%=BibRslts(5)%></td>
                </tr>
            </table>
        <%Else%>
            <p style="border: none;"><%=sErrMsg%></p>
        <%End If%>
    <%Else%>
        <p style="border: none;"><%=sErrMsg%></p>
    <%End If%>

    <div class="row">
        <div class="col-sm-8">
            <h5 class="h5">Would you like to...</h5>
            <ul class="list-group">
                <li class="list-group-item list-group-item-success">
                    ...get your results via email and/or text message within minutes of finishing any GSE-timed event?
                </li>
                <li class="list-group-item list-group-item-success">
                    ...send your results via email and/or text message parents, friends, and others within minutes of 
                    finishing, even if they are not at the meet?
                </li>
                <li class="list-group-item list-group-item-success">
                    ...get access to a free online training log.
                </li>
                <li class="list-group-item list-group-item-success">
                    ...track your performances as well as the performances of your teammates and competitors.
                </li>
            </ul>

            <p>
                Click the image to the right or use your phone's QR code reader to create your own Performance Tracker account.  There is a one-time
                set-up fee of $5...good for as long as you are running or skiing GSE-timed events.
            </p>
        </div>
        <div class="col-sm-4">
            <a href="http://www.gopherstateevents.com/cc_meet/perf_trkr/create_accnt.asp">
                <img src="/graphics/perf_trkr_qr.png" alt="Performance Tracker QR">
            </a>
        </div>
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