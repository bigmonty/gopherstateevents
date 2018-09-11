<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql, rs2, sql2, rs3, sql3
Dim i
Dim lEventID
Dim sGender, sFirstName, sLastName, sQueryFirst, sQueryLast, sErrMsg, sNoRslts, sFnlTime, sRaceName, sEventName, sMyPix, sMyTime
Dim iEventPl
Dim dDOB, dEventDate
Dim FltrRslts(), MyRslts()

If CStr(Session("my_hist_id")) = vbNullString Then Response.Redirect "my_hist_login.asp"

sQueryFirst = Request.QueryString("query_first")
sQueryLast = Request.QueryString("query_last")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Gender, Birthdate FROM PartData WHERE PartID = " & Session("etraxc_id")
rs.Open sql, conn2, 1, 2
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName = Replace(rs(1).Value, "''", "'")
sGender = rs(2).Value
dDOB = rs(3).Value
rs.Close
Set rs = Nothing

Dim sRandPic
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, PixName FROM RacePix ORDER BY NEWID()"
rs.Open sql, conn, 1, 2
sRandPic = "/gallery/" & rs(0).Value & "/" & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

'get picture
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PixURL FROM IndPix WHERE PartID = " & Session("etraxc_id")
rs.Open sql, conn2, 1, 2
If rs.RecordCount > 0 Then 	
	If rs(0).Value & "" = "" Then
		sMyPix = sRandPic
	Else
		If FileExists(rs(0).Value) = True Then 
			sMyPix = "http://www.etraxc.com/graphics/ind_pix/" & rs(0).value
		Else
			sMyPix = sRandPic
		End If
	End If
Else
	sMyPix = sRandPic
End If
rs.Close
Set rs = Nothing

ReDim FltrRslts(5, 0)
ReDim MyRslts(4, 0)

If Request.Form.Item("submit_claim") = "submit_claim" Then
    Dim lRaceID, lParticipantID

    Call MyResults()
    Call GetResults()

    For i = 0 To UBound(FltrRslts, 2) - 1
        If Request.Form.Item("avail_" & FltrRslts(0, i)) = "y" Then  'make sure it's available
            If Request.Form.Item("perf_" & FltrRslts(0, i)) = "on" Then
                'add it to MyHistRaces
                sql = "INSERT INTO MyHistRaces(MyHistID, IndRsltsID) VALUES (" & Session("my_hist_id") & ", " & FltrRslts(0, i) & ")"
                Set rs = conn.Execute(sql)
                Set rs = Nothing

                'change ParticipantID to Session("part_id") in IndResults, PartRace, and PartReg
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT ParticipantID, RaceID FROM IndResults WHERE IndRsltsID = " & FltrRslts(0, i)
                rs.Open sql, conn, 1, 2
                lParticipantID = rs(0).Value
                lRaceID = rs(1).Value
                rs(0).Value = Session("part_id")
                rs.Update
                rs.Close
                Set rs = Nothing

                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT ParticipantID FROM PartRace WHERE ParticipantID = " & lParticipantID & " AND RaceID = " & lRaceID
                rs.Open sql, conn, 1, 2
                rs(0).Value = Session("part_id")
                rs.Update
                rs.Close
                Set rs = Nothing

                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT ParticipantID FROM PartReg WHERE ParticipantID = " & lParticipantID & " AND RaceID = " & lRaceID
                rs.Open sql, conn, 1, 2
                rs(0).Value = Session("part_id")
                rs.Update
                rs.Close
                Set rs = Nothing

                'add it my-etraxc
            End If
        End If
    Next
ElseIf Request.Form.Item("submit_filter") = "submit_filter" Then
    sQueryFirst = Request.Form.Item("first_name")
    sQueryLast = Request.Form.Item("last_name")
End If

Call MyResults()
Call GetResults()

Private Sub MyResults()
    Dim x, y, z
    Dim SortArr(4)

    x = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT IndRsltsID FROM MyHistRaces WHERE MyHistID = " & Session("my_hist_id")
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Call GetRsltsInfo(rs(0).Value)

        MyRslts(0, x) = sEventName
        MyRslts(1, x) = dEventDate
        MyRslts(2, x) = sRaceName
        MyRslts(3, x) = sMyTime
        MyRslts(4, x) = rs(0).Value
        x = x + 1 
        ReDim Preserve MyRslts(4, x)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'sort by event date
    For x = 0 To UBound(MyRslts, 2) - 2
        For y = x + 1 To UBound(MyRslts, 2) - 1
            If CDate(MyRslts(1, x)) < CDate(MyRslts(1, y)) Then
                For z = 0 To 4
                    SortArr(z) = MyRslts(z, x)
                    MyRslts(z, x) = MyRslts(z, y)
                    MyRslts(z, y) = SortArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Sub GetResults()
    Dim x, y, z
    Dim SortArr(5)
    Dim bFound

    If sQueryFirst = vbNullString Then sQueryFirst = sFirstName
    If sQueryLast = vbNullString Then sQueryLast = sLastName

    x = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.IndRsltsID, ir.RaceID, ir.FnlScnds FROM Participant p INNER JOIN IndResults ir "
    sql = sql & "ON p.ParticipantID = ir.ParticipantID WHERE p.FirstName = '" & sQueryFirst & "' AND p.LastName = '" 
    sql = sql & sQueryLast & "' AND p.Gender = '" & sGender & "' AND ir.EventPl > 0 AND ir.FnlScnds > 0 AND (p.DOB = '" & dDOB 
    sql = sql & "' OR p.DOB = '1/1/1900')"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            bFound = False
            For y = 0 To UBound(MyRslts, 2) - 1
                If CLng(rs(0).Value) = CLng(MyRslts(4, y)) Then
                    bFound = True
                    Exit For
                End If
            Next

            If bFound = False Then
                Call GetEventInfo(rs(1).Value)

                FltrRslts(0, x) = rs(0).Value
                FltrRslts(1, x) = sEventName
                FltrRslts(2, x) = dEventDate
                FltrRslts(3, x) = sRaceName
                FltrRslts(4, x) = ConvertToMinutes(rs(2).Value)
                FltrRslts(5, x) = CheckStatus(rs(0).Value)
                x = x + 1
                ReDim Preserve FltrRslts(5, x)
            End If
            rs.MoveNext
        Loop
    Else
        sNoRslts = "Your search returned no results.  Please modify and try again or contact us for assistance."
    End If
    rs.Close
    Set rs = Nothing

    'sort by event date
    For x = 0 To UBound(FltrRslts, 2) - 2
        For y = x + 1 To UBound(FltrRslts, 2) - 1
            If CDate(FltrRslts(2, x)) < CDate(FltrRslts(2, y)) Then
                For z = 0 To 5
                    SortArr(z) = FltrRslts(z, x)
                    FltrRslts(z, x) = FltrRslts(z, y)
                    FltrRslts(z, y) = SortArr(z)
                Next
            End If
        Next
    Next
End Sub

Private Sub GetRsltsInfo(lIndRsltsID)
    Dim lRaceID

    sql2 = "SELECT RaceID, FnlScnds FROM IndResults WHERE IndRsltsID = " & lIndRsltsID
    Set rs2 = conn.Execute(sql2)
    lRaceID = rs2(0).Value
    sMyTime = ConvertToMinutes(rs2(1).Value)
    Set rs2 = Nothing

    sql2 = "SELECT RaceName, EventID FROM RaceData WHERE RaceID = " & lRaceID
    Set rs2 = conn.Execute(sql2)
    sRaceName = rs2(0).Value
    lEventID = rs2(1).Value
    Set rs2 = Nothing

    sql2 = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
    Set rs2 = conn.Execute(sql2)
    sEventName = rs2(0).Value
    dEventDate = rs2(1).Value
    Set rs2 = Nothing
End Sub

Private Sub GetEventInfo(lRaceID)
    sql2 = "SELECT RaceName, EventID, Type FROM RaceData WHERE RaceID = " & lRaceID
    Set rs2 = conn.Execute(sql2)
    sRaceName = rs2(0).Value
    lEventID = rs2(1).Value
    Set rs2 = Nothing

    sql2 = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
    Set rs2 = conn.Execute(sql2)
    sEventName = rs2(0).Value
    dEventDate = rs2(1).Value
    Set rs2 = Nothing
End Sub

Private Function CheckStatus(lMyHistID)
    CheckStatus = "Available"
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT IndRsltsID FROM MyHistRaces WHERE MyHistID = " & lMyHistID
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then 
        If Not CLng(lMyHistID) = CLng(Session("my_hist_id")) Then CheckStatus = "Unavailable"
    End If
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function FileExists(lThisPic)
	FileExists = False
	
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists("C:\Inetpub\h51web\eTRaXC\graphics\ind_pix\" & lThisPic) = True Then FileExists = True
	Set fso = Nothing
End Function

Private Function ConvertToMinutes(sglScnds)
    Dim sHourPart, sMinutePart, sSecondPart
    
    'accomodate a '0' value
    If CSng(sglScnds) <= 0 Then
        ConvertToMinutes = "00:00"
        Exit Function
    End If
    
    'break the string apart
    sMinutePart = CStr(CSng(sglScnds) \ 60)
    sSecondPart = CStr(((CSng(sglScnds) / 60) - (CSng(sglScnds) \ 60)) * 60)
    
    'add leading zero to seconds if necessary
    If CSng(sSecondPart) < 10 Then
        sSecondPart = "0" & sSecondPart
    End If
    
    'make sure there are exactly two decimal places
    If Len(sSecondPart) < 5 Then
        If Len(sSecondPart) = 2 Then
            sSecondPart = sSecondPart & ".00"
        ElseIf Len(sSecondPart) = 4 Then
            sSecondPart = sSecondPart & "0"
        End If
    Else
        sSecondPart = Left(sSecondPart, 5)
    End If
    
    'do the conversion
    If CInt(sMinutePart) <= 60 Then
        ConvertToMinutes = sMinutePart & ":" & sSecondPart
    Else
        sHourPart = CStr(CSng(sMinutePart) \ 60)
        sMinutePart = CStr(CSng(sMinutePart) Mod 60)

        If Len(sMinutePart) = 1 Then
            sMinutePart = "0" & sMinutePart
        End If

        ConvertToMinutes = sHourPart & ":" & sMinutePart & ":" & sSecondPart
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Claim My GSE&copy; Results</title>
<meta name="description" content="Claim my results for a Gopher State Events (GSE) timed event.">
<!--#include file = "../includes/js.asp" --> 

<style type="text/css">
	th,td{
		padding-right:5px;
	}
</style>

<script>
function chkFlds() {
if (document.find_perf.first_name.value == '' || 
    document.find_perf.first_name.value == '')

{
 	alert('You must include a first name and a last name in your search!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="My GSE History Portal">
    <h3 class="h3">My GSE History</h3>

    <!--#include file = "my_hist_nav.asp" -->

    <h4 class="h4">Claim My GSE History Results</h4>

    <div class="bg-info">
        <p>On this page you can claim your results of performances in GSE-managed events.  Please only claim performances from
        events that you are sure are your own.  You may contact us regarding questions, results of yours that have been claimed
        by others, and/or missing results.</p>
    </div>

    <div class="col-sm-5 bg-primary">
        <%If Not sErrMsg = vbNullString Then%>
            <p style="background-color: #fff;"><%=sErrMsg%></p>
        <%End If%>
        <h4 class="h4">Claim My GSE Race Results</h4>
        <h5 class="h5">Search For:</h5>

        <form role="form" class="form-inline" name="find_perf" method="post" action="claim_results.asp" onsubmit="return chkFlds();">
        <div class="form-group">
            <label for="first_name">First Name:</label>
            <input type="text" class="form-control" name="first_name" id="first_name" maxlength="25" value="<%=sQueryFirst%>">
        </div>
        <div class="form-group">
            <label for="last_name">Last Name:</label>
            <input type="text" class="form-control" name="last_name" id="last_name" maxlength="25" value="<%=sQueryLast%>">
        </div>
 	    <div class="form-group">
		    <input type="hidden" name="submit_filter" id="submit_filter" value="submit_filter">
		    <input type="submit" class="form-control" name="submit1" id="submit1" value="Search Now">
	    </div>
        </form>

        <hr>

        <h4 class="h4">Search Results</h4>

        <%If Not sNoRslts = vbNullString Then%>
            <p style="background-color: #fff;"><%=sNoRslts%></p>
        <%End If%>

        <form role="form" class="form" name="claim_perf" method="post" action="claim_results.asp?query_first=<%=sQueryFirst%>&amp;query_last=<%=sQueryLast%>">
        <%For i = 0 To UBound(FltrRslts, 2) - 1%>
            <%If Not FltrRslts(5, i) = "Mine" Then%>
                <ul class="list-unstyled">
                    <li>Event: <%=FltrRslts(1, i)%></li>
                    <li>Date: <%=FltrRslts(2, i)%></li>
                    <li>Race: <%=FltrRslts(3, i)%></li>
                    <li>Time: <%=FltrRslts(4, i)%></li>
                </ul>

                <%If FltrRslts(5, i) = "Available" Then%>
                    <div class="checkbox">
                        <label>
                            <input type="hidden" name="avail_<%=FltrRslts(0, i)%>" id="avail_<%=FltrRslts(0, i)%>" value="y">
                            <input type="checkbox" name="perf_<%=FltrRslts(0, i)%>"  id="perf_<%=FltrRslts(0, i)%>"> 
                            <span style="font-weight: bold;">Claim It!</span>
                        </label>
                    </div>
                <%ElseIf FltrRslts(5, i) = "Unavailable" Then%>
                    <input type="hidden" name="avail_<%=FltrRslts(0, i)%>" id="avail_<%=FltrRslts(0, i)%>" value="n">
                    This race has already been claimed.  If you think this is your race please 
                    <a href="mailto:bob.schneider@gopherstateevents.com" style="font-weight: bold;">Contact Bob Schneider at Gopher State Events</a> to investigate.
                <%End If%>

                <hr>
            <%End If%>
        <%Next%>
        <div class="form-group">
 			<input type="hidden" name="submit_claim" id="submit_claim" value="submit_claim">
			<input type="submit" class="form-control" name="submit2" id="submit2" value="Claim Selected Performances">
        </div>
        </form>
    </div>
    <div class="col-sm-5 bg-success">
        <h4 class="h4">My GSE Race Results</h4>
        <%For i = 0 To UBound(MyRslts, 2) - 1%>
            <ul class="list-unstyled">
                <li>Event: <%=MyRslts(0, i)%></li>
                <li>Date: <%=MyRslts(1, i)%></li>
                <li>Race: <%=MyRslts(2, i)%></li>
                <li>Time: <%=MyRslts(3, i)%></li>
            </ul>
            <hr>
        <%Next%>
    </div>
    <div class="col-sm-2">
		<img class="img-responsive center-block" src="<%=sMyPix%>" alt="My Profile">
    </div>
</div>
<%
conn2.CLose
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>
</body>
</html>
