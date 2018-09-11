<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, k, x, m, n
Dim lEventID, lMyPartID, lMyRaceID, lSuppLegID, lRaceID, lParticipantID
Dim iDeletedPlace, iRaceType, iBibToFind, iTtlRcds, iMinPlace, iMaxPlace, iInsertPlace, iInsertBib, iStartBib, iPlace1, iPlace2, iMyBib, iPlace, iDeletedBib
Dim sEventName, sErrMsg, sEventRaces, sSuppTime, sOtherTime, sChipStart, sActualStart, sNewStart, sElpsdTime, sRaceDelay, sIndDelay, sInsertTime, sNewTime
Dim sngMyTime, sngNewStart, sngFnlTime
Dim dEventDate
Dim BibRslts(6), Races(), IndRslts, RaceParts
Dim bFound

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

iStartBib = Request.QueryString("start_bib")

iBibToFind = 0
iTtlRcds = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim Events
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
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

If Request.Form.item("submit_refresh_times") = "submit_refresh_times" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, RaceID, IndDelay FROM PartRace WHERE RaceID IN (" & sEventRaces & ") ORDER BY ParticipantID"
    rs.Open sql, conn, 1, 2
    RaceParts = rs.GetRows
    rs.Close
    Set rs = Nothing
    
    For i = 0 To UBound(RaceParts, 2)
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RaceDelay FROM RaceData WHERE RaceID = " & RaceParts(1, i)
        rs.Open sql, conn, 1, 2
        Call UpdateTimes(CLng(RaceParts(0, i)), CLng(RaceParts(1, i)), CStr(RaceParts(2, i)), CStr(rs(0).Value))
        rs.Close
        Set rs = Nothing
    Next
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
ElseIf Request.form.Item("submit_new_time") = "submit_new_time" Then 
    iMyBib = Request.Form.Item("bib_to_change")
    sNewTime = Request.Form.Item("new_time")
        
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, RaceID, IndDelay FROM PartRace WHERE Bib = " & iMyBib & " AND RaceID IN (" & sEventRaces & ")"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        lParticipantID = rs(0).Value
        lRaceID = rs(1).Value
        sIndDelay = rs(2).Value
    Else
        sErrMsg = "I'm sorry.  That bib number is not in the listed results."
    End If
    rs.Close
    Set rs = Nothing
        
    If sErrMsg = vbNullString Then
        'change the time
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ElpsdTime FROM IndResults WHERE ParticipantID = " & lParticipantID & " AND RaceID IN (" & sEventRaces & ")"
        rs.Open sql, conn, 1, 2
        rs(0).Value = sNewTime
        rs.Update
        rs.Close
        Set rs = Nothing
        
        'get race delay
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RaceDelay FROM RaceData WHERE RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then sRaceDelay = rs(0).Value
        rs.Close
        Set rs = Nothing
        
        Call UpdateTimes(lParticipantID, lRaceID, sIndDelay, sRaceDelay)
    End If
ElseIf Request.form.Item("submit_bib") = "submit_bib" Then 
    iBibToFind = Request.Form.Item("bib_to_find")
ElseIf Request.Form.Item("submit_insert") = "submit_insert" Then
    iInsertPlace = Request.Form.Item("insert_place")
    iInsertBib = Request.Form.Item("insert_bib")
    sInsertTime = Request.Form.Item("insert_time")

    'first see if this bib is already in the results
    bFound = False
    sql = "SELECT pr.Bib FROM PartRace pr INNER JOIN IndResults ir ON ir.ParticipantID = pr.ParticipantID "
    sql = sql & "INNER JOIN RaceData rd ON (ir.RaceID = rd.RaceID AND pr.RaceID = rd.RaceID) WHERE (rd.EventID = " & lEventID
    sql = sql & " AND rd.RaceID IN (" & sEventRaces & "))"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
        If rs(0).Value = iInsertBib Then
            bFound = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
        
    If bFound = True Then sErrMsg = "That bib already exists in the results."
    
    If sErrMsg = vbNullString Then    
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT EventPl FROM IndResults WHERE RaceID IN (" & sEventRaces & ") AND EventPl >= " & iInsertPlace
        sql = sql & " ORDER BY EventPl"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            rs(0).Value = rs(0).Value + 1
            rs.Update
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        'get race id
        bFound = False
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RaceID, ParticipantID FROM PartRace WHERE Bib = " & iInsertBib & " AND RaceID IN (" & sEventRaces & ")"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            lRaceID = rs(0).Value
            lParticipantID = rs(1).Value
            bFound = True
        End If
        rs.Close
        Set rs = Nothing
        
        If bFound = True Then
            sql = "INSERT INTO IndResults(ParticipantID, RaceID, EventPl, ElpsdTime, Source) VALUES (" & lParticipantID & ", "
            sql = sql & lRaceID & ", " & iInsertPlace & ", '" & sInsertTime & "', 'manual')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
            
            'need ind delay
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT IndDelay FROM PartRace WHERE RaceID = " & lRaceID & " AND ParticipantID = " & lParticipantID
            rs.Open sql, conn, 1, 2
            sIndDelay = rs(0).Value
            rs.Close
            Set rs = Nothing
            
            'get race delay
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT RaceDelay FROM RaceData WHERE RaceID = " & lRaceID
            rs.Open sql, conn, 1, 2
            sRaceDelay = rs(0).Value
            rs.Close
            Set rs = Nothing
            
            Call UpdateTimes(lParticipantID, lRaceID, sIndDelay, sRaceDelay)
        End If
    End If

    Call RcrdChanges(lParticipantID, lRaceID, iInsertPlace, sInsertTime)
ElseIf Request.Form.Item("submit_delete") = "submit_delete" Then
    iDeletedBib = Request.Form.Item("delete_bib")
    
    'make sure this place has been assigned
    lParticipantID = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, RaceID FROM PartRace WHERE RaceID IN (" & sEventRaces & ") AND Bib = " & iDeletedBib
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        lParticipantID = rs(0).Value
        lRaceID = rs(1).Value
    End If
    rs.Close
    Set rs = Nothing
    
    If CLng(lParticipantID) > 0 Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ElpsdTime FROM IndResults WHERE RaceID IN (" & sEventRaces & ") AND ParticipantID >= " & lParticipantID
        rs.Open sql, conn, 1, 2
        sElpsdTime = rs(0).Value
        rs.Close
        Set rs = Nothing

        sql = "DELETE FROM IndResults WHERE RaceID IN (" & sEventRaces & ") AND ParticipantID >= " & lParticipantID
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        Call RcrdChanges(lParticipantID, lRaceID, iDeletedBib, sElpsdTime)
    End If
ElseIf Request.Form.Item("submit_reset") = "submit_reset" Then
    i = 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventPl FROM IndResults WHERE RaceID IN (" & sEventRaces & ") ORDER BY EventPl"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        rs(0).Value = i
        rs.Update
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_swap") = "submit_swap" Then
    iPlace1 = Request.Form.Item("place_1")
    iPlace2 = Request.Form.Item("place_2")

    i = 1
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventPl FROM IndResults WHERE RaceID IN (" & sEventRaces & ") AND EventPl IN (" & iPlace1 & ", " & iPlace2 & ") ORDER BY EventPl"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        If i = 1 Then
            rs(0).Value = iPlace2
            i = i + 1
        Else
            rs(0).Value = iPlace1
        End If
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_start_bib") = "submit_start_bib" Then
    iStartBib = Request.Form.Item("start_bib")
    
    sChipStart = "00:00"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ir.ChipStart FROM IndResults ir INNER JOIN PartRace pr  ON ir.ParticipantID = pr.ParticipantID WHERE ir.RaceID IN (" & sEventRaces 
    sql = sql & ") AND pr.Bib = " & iStartBib
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sChipStart = rs(0).Value
    rs.Close
    Set rs = Nothing
    
    sActualStart = "00:00"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ActualStart FROM StartData WHERE EventID = " & lEventID & " AND Bib = " & iStartBib
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then sActualStart = rs(0).Value
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_start") = "submit_start" Then
    sNewStart = Request.Form.Item("chip_start")
    sngNewStart = Round(ConvertToSeconds(sNewStart), 3)

    'then update indresults if record exists yet
    sql = "SELECT ParticipantID, RaceID, IndDelay FROM PartRace WHERE RaceID IN (" & sEventRaces & ") AND Bib = " & iStartBib
    Set rs = conn.Execute(sql)
    lParticipantID = rs(0).Value
    lMyRaceID = rs(1).Value
    sIndDelay = rs(2).Value
    Set rs = Nothing

    sql = "SELECT RaceDelay FROM RaceData WHERE RaceID = " & lMyRaceID
    Set rs = conn.Execute(sql)
    sRaceDelay = rs(0).Value
    Set rs = Nothing

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FnlTime, ChipTime, ElpsdTime, ChipStart, FnlScnds FROM IndResults WHERE RaceID = " & lMyRaceID & " AND ParticipantID = " 
    sql = sql & lParticipantID
    rs.Open sql, conn, 1, 2
    sngFnlTime = Round(ConvertToSeconds(rs(2).Value) - ConvertToSeconds(sRaceDelay) - ConvertToSeconds(sIndDelay), 3)
    rs(0).Value = ConvertToMinutes(sngFnlTime)
    rs(1).Value = ConvertToMinutes(sngFnlTime - sngNewStart)
    rs(3).Value = sNewStart
    rs(4).Value = sngFnlTime - sngNewStart
    rs.Update
    rs.Close
    Set rs = Nothing
End If

sql = "SELECT IndRsltsID FROM IndResults WHERE RaceID IN (" & sEventRaces & ")"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn, 1, 2
If rs.RecordCount  > 0 Then iTtlRcds = rs.RecordCount
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")	
sql = "SELECT rd.RaceName, ir.EventPl, pr.Bib, p.FirstName, p.LastName, "
sql = sql & "pr.AgeGrp, p.Gender, ir.ElpsdTime, rd.RaceDelay, pr.IndDelay, ir.ChipStart, ir.FnlTime, ir.ChipTime "
sql = sql & "FROM Participant p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID JOIN PartRace pr ON "
sql = sql & "pr.ParticipantID = p.ParticipantID JOIN RaceData rd ON ir.RaceID = rd.RaceID  AND pr.RaceID = rd.RaceID "
sql = sql & "WHERE rd.EventID = " & lEventID & " ORDER BY ir.EventPl"
'sql = "SELECT rd.RaceName, ir.EventPl, pr.Bib, p.FirstName, p.LastName, "
'sql = sql & "p.Gender, pr.Age, ir.ElpsdTime, rd.RaceDelay, pr.IndDelay, ir.ChipStart, ir.FnlTime, ir.ChipTime "
'sql = sql & "FROM Participant p INNER JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID JOIN PartRace pr ON "
'sql = sql & "pr.ParticipantID = p.ParticipantID JOIN RaceData rd ON ir.RaceID = rd.RaceID  AND pr.RaceID = rd.RaceID "
'sql = sql & "WHERE rd.EventID = " & lEventID & " ORDER BY ir.EventPl"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    IndRslts = rs.GetRows()
Else
    ReDim IndRslts(12, 0)
End If
rs.Close
Set rs = Nothing

If Not CInt(iBibToFind) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID, RaceID, Age FROM PartRace WHERE Bib = " & iBibToFind & " AND RaceID IN (" & sEventRaces & ")"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        lMyPartID = rs(0).Value
        lMyRaceID = rs(1).Value
        BibRslts(2) = rs(2).Value
    Else
        sErrMsg = "I'm sorry.  That bib number was not found."
    End If
    rs.Close
    Set rs = Nothing

    If sErrMsg = vbNullString Then
	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT LastName, FirstName, Gender FROM Participant WHERE ParticipantID = " & lMyPartID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then BibRslts(1) = Replace(rs(0).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
        rs.Close
        Set rs = Nothing

        k = 1
	    Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ParticipantID, EventPl, FnlScnds, FnlTime FROM IndResults WHERE RaceID = " & lMyRaceID & " AND FnlScnds > 0 ORDER BY FnlScnds"
        rs.Open sql, conn, 1, 2
        Do While Not rs.EOF
            If CLng(rs(0).Value) = CLng(lMyPartID) Then
                BibRslts(0) = k
                Exit Do
            Else
                k = k + 1
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

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ChipTime, FnlTime, ChipStart FROM IndResults WHERE RaceID = " & lMyRaceID & " AND ParticipantID = " & lMyPartID
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            BibRslts(4) = rs(0).Value
            BibRslts(5) = rs(1).Value
            BibRslts(6) = rs(2).Value
        Else
            sErrMsg = "I'm sorry.  That bib number was not found."
        End If
        rs.Close
        Set rs = Nothing
    End If

    If CLng(lSuppLegID) > 0 Then
        sSuppTime = "00:00.000"
        sOtherTime = "00:00:00.000"

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT SuppTime, OtherTime FROM SuppLegRslts WHERE SuppLegID = " & lSuppLegID & " AND Bib = " & iBibToFind
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then
            sSuppTime = rs(0).Value
            sOtherTime = ConvertToMinutes(ConvertToSeconds(BibRslts(4)) - ConvertToSeconds(rs(0).Value))
        End If
        rs.Close
        Set rs = Nothing
    End If

    If CInt(iMinPlace) <= 0 Then 
        iMinPlace = 1
        iMaxPlace = 20
    End If
End If

Private Sub RcrdChanges(lThisPart, lThisRace, iThisPlace, sThisTime)
    Dim sParticipant, sRace
    Dim iBib
    
    sRace = "None Avail"
    iBib = 0
    sParticipant = "Not Available"
    
    If Not CLng(lThisRace) = 0 Then
        If Not CLng(lParticipantID) = 0 Then
            'get bib
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT pr.Bib FROM RaceData rd INNER JOIN PartRace pr on rd.RaceID = pr.RaceID WHERE pr.RaceID = " & lThisRace
            sql = sql & " AND pr.ParticipantID = " & lThisPart
            rs.Open sql, conn, 1, 2
            iBib = rs(0).Value
            rs.Close
            Set rs = Nothing
    
            'get participant
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT LastName, FirstName FROM Participant WHERE ParticipantID = " & lThisPart
            rs.Open sql, conn, 1, 2
            sParticipant = rs(0).Value & ", " & rs(1).Value
            rs.Close
            Set rs = Nothing
        End If
        
        'get race name
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lThisRace
        rs.Open sql, conn, 1, 2
        sRace = rs(0).Value
        rs.Close
        Set rs = Nothing
    End If
    
    sql = "INSERT INTO ChangesSql (Participant, Race, EventID, Place, Bib, Time, WhenMade, Action) VALUES ('" & Left(sParticipant, 10) & "', '" 
    sql = sql & Left(sRace, 6) & "', " & lEventID & ", " & iThisPlace & ", " & iBib & ", '" & sThisTime & "', '" & Now() & "', 'Insert Finisher')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End Sub

Public Sub UpdateTimes(lThisPart, lThisRace, sIndDelay, sRaceDelay)
    Dim rs2, sql2
    Dim sngDelay, sngElapsedTime, sngChipStart, sngChipTime, sngFnlTime, sngFnlScnds
    
    sngDelay = ConvertToSeconds(sRaceDelay) + ConvertToSeconds(sIndDelay)

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT ChipStart FROM RaceData WHERE RaceID = " & lThisRace
    rs2.Open sql2, conn, 1, 2
    sChipStart = rs2(0).Value
    rs2.Close
    Set rs2 = Nothing

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT ElpsdTime, ChipStart, ChipTime, FnlTime, FnlScnds FROM IndResults WHERE RaceID = " & lThisRace & " AND ParticipantID = " & lThisPart
    rs2.Open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then
        sngElapsedTime = ConvertToSeconds(rs2(0).Value)                 'get elapsed time in seconds
        sngChipStart = Round(ConvertToSeconds(rs2(1).Value), 2)         'get chip start in seconds
        sngChipTime = sngElapsedTime - sngDelay - sngChipStart          'get chip time in seconds
        sngFnlTime = sngElapsedTime - sngDelay                          'get fnl time (gun time) in seconds

        If sChipStart = "y" Then
            sngFnlScnds = sngChipTime                                   'use chip time for fnl sncds
        Else
            sngFnlScnds = sngFnlTime                                    'use gun time for fnl sncds
        End If

        rs2(2).Value = ConvertToMinutes(sngChipTime)                    'get chip time in minutes
        rs2(3).Value = ConvertToMinutes(sngFnlTime)                     'get fnl time in minutes
        rs2(4).Value = sngFnlScnds
        
        rs2.Update
    End If
    rs2.Close
    Set rs2 = Nothing
End Sub
%>
<!--#include file = "../../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../../includes/convert_to_minutes.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>Results For <%=sEventName%></title>
<meta name="description" content="Fitness Event Results from Gopher State Events, a conventional timing service offered by H51 Software, LLC in Minnetonka, MN.">
<!--#include file = "../../../includes/js.asp" -->

<link rel="canonical" href="/results/fitness_events/results.asp" />

<script>
function chkFlds(){
 	if (document.insert_finisher.insert_place.value == '' || 
 	    document.insert_finisher.insert_bib.value == '' ||
 	    document.insert_finisher.insert_time.value == '')
		{
  		alert('All fields are required.');
  		return false
  		}
 	else
		if (isNaN(document.insert_finisher.insert_place.value) ||
		   isNaN(document.insert_finisher.insert_bib.value))
    		{
			alert('Bib and place must be numeric values!');
			return false
			} 	
	else
   		return true
}

function chkFlds2(){
 	if (document.delete_finisher.delete_place.value == '')
		{
  		alert('Please select a place to delete.');
  		return false
  		}
 	else
		if (isNaN(document.delete_finisher.delete_place.value))
    		{
			alert('Place must be numeric!');
			return false
			} 	
	else
   		return true
}
    
function chkFlds3(){
 	if (document.get_start.start_bib.value == '')
		{
  		alert('Please enter a bib.');
  		return false
  		}
 	else
		if (isNaN(document.get_start.start_bib.value))
    		{
			alert('Bib must be numeric!');
			return false
			} 	
	else
   		return true
}

    function chkFlds4(){
 	if (document.swap_places.place_1.value == '' ||
        document.swap_places.place_2.value == '' )
		{
  		alert('Please enter places to swap.');
  		return false
  		}
 	else
		if (isNaN(document.swap_places.place_1.value) ||
            isNaN(document.swap_places.place_2.value))
    		{
			alert('Both places must be numeric!');
			return false
			} 	
	else
   		return true
}
</script>
</head>

<<body>
<div class="container">
  	<!--#include file = "../../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h3 class="h3">Gopher State Events Results: <%=sEventName%> On <%=dEventDate%></h3>
			
			<form class="form-inline" name="which_event" method="post" action="results_mgr.asp?event_id=<%=lEventID%>">
			<label for="events">Events:</label>
			<select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				<option value="">&nbsp;</option>
				<%For i = 0 to UBound(Events, 2)%>
					<%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						<option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%Else%>
						<option value="<%=Events(0, i)%>"><%=Events(1, i)%> (<%=Events(2, i)%>)</option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			<input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event">
			</form>
			<br>

			<%If Not Clng(lEventID) = 0 Then%>
			    <!--#include file = "../../../includes/event_nav.asp" -->

                <h5 class="h5">Total Finishers:</span>&nbsp;<%=iTtlRcds%></h5>

                <div class="row">
                    <div class="col-sm-3 bg-success">
                        <h5 class="h5">Insert Finisher</h5>

	                    <form class="form" name="insert_finisher" method="post" action="results_mgr.asp?event_id=<%=lEventID%>" onsubmit="return chkFlds();">			
		                <div class="form-group">
			                <label for="insert_place" class="control-label col-xs-4">Place:</label>
			                <div class="col-xs-8">
                                <input type="text" class="form-control" name="insert_place" id="insert_place">
                            </div>
		                </div>
		                <div class="form-group">
			                <label for="insert_bib" class="control-label col-xs-4">Bib:</label>
			                <div class="col-xs-8">
                                <input type="text" class="form-control" name="insert_bib" id="insert_bib">
                            </div>
		                </div>
		                <div class="form-group">
			                <label for="insert_time" class="control-label col-xs-4">Time:</label>
			                <div class="col-xs-8">
                                <input type="text" class="form-control" name="insert_time" id="insert_time">
                            </div>
		                </div>
		                <div class="form-group">
				            <input type="hidden" name="submit_insert" id="submit_insert" value="submit_insert">
				            <input type="submit" class="form-control" name="submit1" id="submit1" value="Insert Finisher">
		                </div>
	                    </form>
                    </div>
                    <div class="col-sm-3 bg-info">
                        <h5 class="h5">Edit Time</h5>

	                    <form class="form" name="edit_time" method="post" action="results_mgr.asp?event_id=<%=lEventID%>">			
		                <div class="form-group">
			                <label for="bib_to_change" class="control-label col-xs-5">Bib:</label>
			                <div class="col-xs-7">
                                <input type="text" class="form-control" name="bib_to_change" id="bib_to_change">
                            </div>
		                </div>
		                <div class="form-group">
			                <label for="new_time" class="control-label col-xs-5">New Time:</label>
			                <div class="col-xs-7">
                                <input type="text" class="form-control" name="new_time" id="new_time">
                            </div>
		                </div>
		                <div class="form-group">
				            <input type="hidden" name="submit_new_time" id="submit_new_time" value="submit_new_time">
				            <input type="submit" class="form-control" name="submit7" id="submit7" value="Make Change">
		                </div>
	                    </form>
                    </div>
                    <div class="col-sm-3 bg-warning">
                        <h5 class="h5">Delete Finisher</h5>

	                    <form class="form" name="delete_finisher" method="post" action="results_mgr.asp?event_id=<%=lEventID%>" onsubmit="return chkFlds2();">			
		                <div class="form-group">
			                <label for="delete_bib" class="control-label col-xs-7">Bib to Delete:</label>
			                <div class="col-xs-5">
                                <input type="text" class="form-control" name="delete_bib" id="delete_bib">
                            </div>
		                </div>
		                <div class="form-group">
			                <label for="keep_open" class="control-label col-xs-7">Keep Open:</label>
			                <div class="col-xs-5">
                                <input type="checkbox" name="keep_open" id="keep_open">
                            </div>
		                </div>
			            <div class="form-group">
				            <input type="hidden" name="submit_delete" id="submit_delete" value="submit_delete">
				            <input type="submit" class="form-control" name="submit2" id="submit2" value="Delete Finisher">
			            </div>
	                    </form>
                    </div>
                    <div class="col-sm-3 bg-danger">
                        <h5 class="h5">Swap Places</h5>

	                    <form name="swap_places" method="post" action="results_mgr.asp?event_id=<%=lEventID%>" onsubmit="return chkFlds4();">	
		                <div class="form-group">
			                <label for="place_1" class="control-label col-xs-6">Place 1:</label>
			                <div class="col-xs-6">
                                <input type="text" class="form-control" name="place_1" id="place_1">
                            </div>
		                </div>
		                <div class="form-group">
			                <label for="place_2" class="control-label col-xs-6">Place 2:</label>
			                <div class="col-xs-6">
                                <input type="text" class="form-control" name="place_2" id="place_2">
                            </div>
		                </div>
			            <div class="form-group">
				            <input type="hidden" name="submit_swap" id="submit_swap" value="submit_swap">
				            <input type="submit" class="form-control" name="submit5" id="submit5" value="Swap Places">
			            </div>
	                    </form>
                    </div>
                </div>

                <div class="row">
                    <div class="col-sm-4 bg-danger">
                        <h5 class="h5">Edit Chip Start</h5>

	                    <form class="form-inline"name="get_start" method="post" action="results_mgr.asp?event_id=<%=lEventID%>" onsubmit="return chkFlds3();">			
			            <label for="start_bib">Start For Bib:</label>
			            <input type="text" class="form-control" name="start_bib" id="start_bib" value="<%=iStartBib%>">
				        <input type="hidden" name="submit_start_bib" id="submit_start_bib" value="submit_start_bib">
				        <input type="submit" class="form-control" name="submit3" id="submit3" value="Submit Bib">
	                    </form>

                        <%If Not sChipStart = vbNullString Then%>
                            <hr>
	                        <form class="form-inline" name="edit_start" method="post" action="results_mgr.asp?event_id=<%=lEventID%>&amp;start_bib=<%=iStartBib%>">			
			                <label>Old Time:</label>
			                <label><%=sActualStart%>&nbsp;&nbsp;&nbsp;</label>
			                <label for ="chip_start">New Time:</label>
			                <input type="text" class="form-control" name="chip_start" id="chip_start" value="<%=sChipStart%>">
				            <input type="hidden" name="submit_start" id="submit_start" value="submit_start">
				            <input type="submit" class="form-control" name="submit4" id="submit4" value="Submit Change">
	                        </form>
                        <%End If%>
                    </div>
                    <div class="col-sm-4 bg-warning">
                        <h5 class="h5">Refresh Times</h5>

	                    <form class="form" name="update_times" method="post" action="results_mgr.asp?event_id=<%=lEventID%>">	
                        <p>This utility will refresh all times for all participants in all races in this event.</p>	
                        <div>
		                    <input type="hidden" name="submit_refresh_times" id="submit_refresh_times" value="submit_refresh_times">
		                    <input type="submit" class="form-control" name="submit6x" id="submit6x" value="Refresh All Times">
                        </div>	
	                    </form>
                    </div>
                    <div class="col-sm-4 bg-success">
                        <h5 class="h5">Reset Places</h5>

	                    <form class="form" name="get_start" method="post" action="results_mgr.asp?event_id=<%=lEventID%>">	
                        <p>This utility will re-number all event places based on the existing place values and elapsed time.</p>	
                        <div>
		                    <input type="hidden" name="submit_reset" id="submit_reset" value="submit_reset">
		                    <input type="submit" class="form-control"name="submit6" id="submit6" value="Condense Places">
                        </div>	
	                    </form>
                    </div>
                </div>
                <div class="row bg-info" style="margin: 10px 0 10px 0;">
                    <div class="col-xs-4">
                        <br>
                        <form class="form-inline" name="find_bib" method="post" action="results_mgr.asp?event_id=<%=lEventID%>">
                        <label for="bib_to_find">Bib To Find:</label>
                        <input type="text" name="bib_to_find" id="bib_to_find" size="3" value="<%=iBibToFind%>">
                        <input type="hidden" name="submit_bib" id="submit_bib" value="submit_bib">
                        <input type="submit" class="form-control" name="submit_lookup" id="submit_lookup" value="Find Bib">
                        </form>
                        <br>
                    </div>
                    <div class="col-xs-8">
                        <%If Not CInt(iBibToFind) = 0 Then%>
                            <%If sErrMsg = vbNullString Then%>
                                <table class="table table-condensed">
                                    <tr><th>Place</th><th>Name</th><th>MF</th><th>Age</th><th>Chip Time</th><th>Gun Time</th><th>Chip Start</th></tr>
                                    <tr>
                                        <%For i = 0 To 6%>
                                            <td><%=BibRslts(i)%></td>
                                        <%Next%>
                                    </tr>
                                </table>
                            <%Else%>
                                <p><%=sErrMsg%></p>
                            <%End If%>
                        <%End If%>
                    </div>
                </div>
		        <table class="table table-striped table-condensed">
			        <tr>
                        <th>Race</th>
				        <th>Pl</th>
				        <th>Bib</th>
                        <th>Name</th>
                        <th>M/F</th>
  				        <th>Age</th>
				        <th>Elpsd</th>
				        <th>R-Delay</th>
                        <th>I-Delay</th>
				        <th>Start</th>
				        <th>Gun</th>
                        <th>Chip</th>
			        </tr>
			        <%For i = 0 To UBound(IndRslts, 2)%>
					    <tr>
						    <td><%=IndRslts(0, i)%></td>
                            <td><%=IndRslts(1, i)%></td>
                            <td><%=IndRslts(2, i)%></td>
						    <td><%=IndRslts(4, i)%>, <%=IndRslts(3, i)%></td>
						    <td><%=IndRslts(5, i)%></td>
			                <td><%=IndRslts(6, i)%></td>
                            <td><%=IndRslts(7, i)%></td>
						    <td><%=IndRslts(8, i)%></td>
						    <td><%=IndRslts(9, i)%></td>
						    <td><%=IndRslts(10, i)%></td>
						    <td><%=IndRslts(11, i)%></td>
                            <td><%=IndRslts(12, i)%></td>
					    </tr>
			        <%Next%>
		        </table>
            <%End If%>
        </div>
	</div>	
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>