<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID
Dim sEventName, sGender, sRaceName, sMF, sRacesToBlend, sRaceNames, sRaceID, sThisGender, sShowAge
Dim dEventDate
Dim Events(), Races(), IndRslts, SelRaces(), SelGenders(), Genders(1, 1)

lEventID = Request.QueryString("event_id")

Genders(0, 0) = "M"
Genders(1, 0) = "Male"
Genders(0, 1) = "F"
Genders(1, 1) = "Female"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

'get races
i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

ReDim SelRaces(0)
ReDim SelGenders(0)
If Request.Form.Item("submit_races") = "submit_races" Then
	sRacesToBlend = Request.Form.Item("races")

    j = 0
    For i = 1 To Len(sRacesToBlend)
        If Mid(sRacesToBlend, i, 1) = "," Then
            SelRaces(j) = sRaceID
            j = j + 1
            ReDim Preserve SelRaces(j)

            If sRaceNames = vbNullString Then
                sRaceNames = GetRaceName(CLng(sRaceID))
            Else
                sRaceNames = sRaceNames & ", " & GetRaceName(CLng(sRaceID))
            End If
            sRaceID = vbNullString
        Else
            sRaceID = sRaceID & Mid(sRacesToBlend, i, 1)
            If i = Len(sRacesToBlend) Then 
                SelRaces(j) = sRaceID
                j = j + 1
                ReDim Preserve SelRaces(j)

                sRaceNames = sRaceNames & ", " & GetRaceName(CLng(sRaceID))
            End If
        End If
    Next

	sGender = Request.Form.Item("gender")

    j = 0
    ReDim SelGenders(0)
    For i = 1 To Len(sGender)
        If Mid(sGender, i, 1) = "," Then
            SelGenders(j) = sThisGender
            j = j + 1
            ReDim Preserve SelGenders(j)
        Else
            sThisGender = Mid(sGender, i, 1)
            If i = Len(sGender) Then 
                SelGenders(j) = sThisGender
                j = j + 1
                ReDim Preserve SelGenders(j)
            End If
        End If
    Next

    sGender = vbNullString

    For i = 0 To UBound(SelGenders) - 1
        sGender = sGender & "'" & SelGenders(i) & "',"
    Next
    sGender = Left(sGender, Len(sGender) - 1)
End If

If Not sRacesToBlend = vbNullString Then
    sql = "SELECT pr.Bib, p.LastName, p.FirstName, p.Gender, pr.Age, ir.ChipTime, ir.FnlTime, ir.ChipStart, p.City, p.St, ir.RaceID "
    sql = sql & "FROM Participant p JOIN IndResults ir ON p.ParticipantID = ir.ParticipantID "
    sql = sql & "JOIN PartRace pr ON pr.RaceID = ir.RaceID AND pr.ParticipantID = p.ParticipantID "
    sql = sql & "WHERE ir.RaceID IN (" & sRacesToBlend & ") AND ir.FnlTime > '00:00:00.000' AND p.Gender IN (" & sGender & ") ORDER BY ir.FnlScnds"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        IndRslts = rs.GetRows()
    Else
        ReDim IndRslts(10, 0)
    End If
    rs.Close
    Set rs = Nothing
End If

%>
<!--#include file = "../../includes/convert_to_seconds.asp" -->
<!--#include file = "../../includes/convert_to_minutes.asp" -->
<%

Private Function GetRaceName(lThisRace)
    GetRaceName = "undetermined"
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lThisRace
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetRaceName = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Blend GSE Results</title>
<meta name="description" content="Gopher State Events (GSE) Blended Results.">
 <!--#include file = "../../includes/js.asp" --> 
<script>
function chkFlds() {
 	if (document.get_races.races.value == '')
		{
  		alert('You must select at least one race.');
  		return false
  		}
 	else
		if (document.get_races.gender.value == '')
    		{
			alert('You must select at least one gender.');
			return false
			}
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-sm-6">
            <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
        </div>
        <div class="col-sm-6">
            <h1 class="h1">GSE Blended Results</h1>
        </div>
    </div>

    <div class="bg-danger">
        <a href="javascript:window.print();" style="color:#fff;">Print</a>
    </div>
    
    <h2 class="h2"><%=sEventName%> On <%=dEventDate%></h2>

    <div class="bg-success">
        Sometimes it can be kind of cool to see how people compared between races and genders.  This utility allows you to get a listing of finishers by
        finish time based on races and genders that you select from that event.  Generally speaking, this utility is meaningless over different distances
        but could have some value if only modality has changed (ie: open, single-speed or fat tire in biking or classical and freestyle in nordic skiing).
        Just select at least one race and at least one gender and have a look.
    </div>

    <form class="form-inline" name="get_races" method="post" action="blended_results.asp?event_id=<%=lEventID%>" onsubmit="return chkFlds();">
    <label for="races">Races:</label>
    <select class="form-control" name="races" id="races" multiple size="<%=UBound(Races, 2)%>">
        <%For i = 0 to UBound(Races, 2) - 1%>
            <%If UBound(SelRaces) > 0 Then%>
                <%For j = 0 To UBound(SelRaces) - 1%>
                    <%If CLng(Races(0, i)) = CLng(SelRaces(j)) Then%>
                        <option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
                        <%Exit For%>
                    <%Else%>
                        <%If j = UBound(SelRaces) - 1 Then%>
                            <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
                        <%End If%>
                    <%End If%>
                <%Next%>
            <%Else%>
                <option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
            <%End If%>
        <%Next%>
    </select>
    <label for="gender">Gender:</label>
    <select class="form-control" name="gender" id="gender" multiple size="2">
        <%For i = 0 to UBound(Genders, 2)%>
            <%If UBound(SelGenders) > 0 Then%>
                <%For j = 0 To UBound(SelGenders) - 1%>
                    <%If Genders(0, i) = SelGenders(j) Then%>
                        <option value="<%=Genders(0, i)%>" selected><%=Genders(1, i)%></option>
                        <%Exit For%>
                    <%Else%>
                        <%If j = UBound(SelGenders) - 1 Then%>
                            <option value="<%=Genders(0, i)%>"><%=Genders(1, i)%></option>
                        <%End If%>
                    <%End If%>
                <%Next%>
            <%Else%>
                <option value="<%=Genders(0, i)%>"><%=Genders(1, i)%></option>
            <%End If%>
        <%Next%>
    </select>
    <input type="hidden" class="form-control" name="submit_races" id="submit_races" value="submit_races">
    <input type="submit" class="form-control" name="get_races" id="get_races" value="Blend These" style="font-size:0.9em;">
    </form>

	<%If Not sRacesToBlend = vbNullString Then%>
        <h4 class="h4">Blended Results For <%=sRaceNames%></h4>
		<table class="table table-striped">
			<tr>
				<th>Pl</th>
				<th>Bib-Name</th>
				<th>M/F</th>
  				<th>Age</th>
                <th>Race</th>
				<th>Chip Time</th>
				<th>Gun Time</th>
				<th>Start Time</th>
				<th>From</th>
			</tr>

			<%For i = 0 To UBound(IndRslts, 2)%>
				<tr>
					<td><%=i + 1%></td>
					<td><%=IndRslts(0, i)%> - <%=IndRslts(1, i)%>, <%=IndRslts(2, i)%></td>
					<td><%=IndRslts(3, i)%></td>
			        <td><%=IndRslts(4, i)%></td>
                    <td><%=GetRaceName(IndRslts(10, i))%></td>
					<td><%=IndRslts(5, i)%></td>
					<td><%=IndRslts(6, i)%></td>
					<td><%=IndRslts(7, i)%></td>
					<td><%=IndRslts(8, i)%>, <%=IndRslts(9, i)%></td>
				</tr>
			<%Next%>
		</table>
	<%End If%>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>