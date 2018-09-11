<%@ Language=VBScript%>
<%
Option Explicit
%>

<!--#include file = "../../includes/JSON_2.0.4.asp" -->

<%
Dim conn, rs, sql
Dim lRaceID, lEventID
Dim i, j
Dim sShowAge, sSortRsltsBy, sOrderBy, sGender
Dim iLength
Dim IndRslts

lRaceID = Request.QueryString("race_id")
sGender = Request.QueryString("gender")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If lRaceID = "0" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then lRaceID = rs(0).Value
    rs.Close
    Set rs = Nothing
End If

sql = "SELECT ShowAge, SortRsltsBy, EventID FROM RaceData WHERE RaceID = " & lRaceID
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

If sGender = "B" Then
    sql = "SELECT P.Country, PR.Bib, P.FirstName, P.LastName, P.Gender, PR.Age, IR.ChipTime, IR.FnlTime, IR.ChipStart, P.City, P.St, P.FBook "
    sql = sql & "FROM dbo.RaceData AS R INNER JOIN dbo.PartRace AS PR ON R.RaceID = PR.RaceID INNER JOIN "
    sql = sql & "dbo.Participant AS P ON PR.ParticipantID = P.ParticipantID INNER JOIN "
    sql = sql & "(SELECT DISTINCT RaceID, ParticipantID, ChipTime, FnlTime, ChipStart, FnlScnds, EventPl FROM dbo.IndResults "
    sql = sql & "WHERE (FnlTime IS NOT NULL) AND (FnlTime > '00:00:00.000')) AS IR ON R.RaceID = IR.RaceID AND P.ParticipantID = IR.ParticipantID "
    sql = sql & "WHERE (R.RaceID = " & lRaceID & ") ORDER BY " & sOrderBy   
Else
    sql = "SELECT P.Country, PR.Bib, P.FirstName, P.LastName, P.Gender, PR.Age, IR.ChipTime, IR.FnlTime, IR.ChipStart, P.City, P.St, P.FBook "
    sql = sql & "FROM dbo.RaceData AS R INNER JOIN dbo.PartRace AS PR ON R.RaceID = PR.RaceID INNER JOIN "
    sql = sql & "dbo.Participant AS P ON PR.ParticipantID = P.ParticipantID INNER JOIN "
    sql = sql & "(SELECT DISTINCT RaceID, ParticipantID, ChipTime, FnlTime, ChipStart, FnlScnds, EventPl FROM dbo.IndResults "
    sql = sql & "WHERE (FnlTime IS NOT NULL) AND (FnlTime > '00:00:00.000')) AS IR ON R.RaceID = IR.RaceID AND P.ParticipantID = IR.ParticipantID "
    sql = sql & "WHERE (R.RaceID = " & lRaceID & ") AND P.Gender = '" & sGender & "' ORDER BY " & sOrderBy   
End If
Set rs = conn.Execute(sql)
If True = rs.BOF Then
    ReDim IndRslts(11, 0)
Else
    IndRslts = rs.GetRows()
End If
'rs.Close
Set rs = Nothing

For i = 0 To UBound(IndRslts, 2)
    IndRslts(0, i) = i + 1
    IndRslts(2, i) = "<a href=javascript:pop('my_rslts.asp?race_id=" & lRaceID & "&event_id=" & lEventID & "&bib=" & IndRslts(1, i) & "',350,550)>" & Replace(IndRslts(2, i), """", "") & "</a>"
'    IndRslts(2, i) = "<a nof='LS_E' href='javascript:void(0)' onMouseOver=window.open('my_rslts.asp?race_id=" & lRaceID & "&event_id=" & lEventID & "&bib=" & IndRslts(1, i) & "','popover','toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=0,copyhistory=0,width=400,height=525,left=20,top=0')>" & Replace(IndRslts(2, i), """", "") & "</a>"    
    IndRslts(3, i) = Replace(IndRslts(3, i), """","")
    IndRslts(6, i) = Replace(IndRslts(6, i), "-", "")
    IndRslts(7, i) = Replace(IndRslts(7, i), "-", "")
    IndRslts(8, i) = Replace(IndRslts(8, i), "-", "")

    If Not IndRslts(9, i) & "" = "" Then IndRslts(9, i) = Replace(IndRslts(9, i), """", "")
    If Not IndRslts(10, i) & "" = "" Then  IndRslts(10, i) = Replace(IndRslts(10, i), """", "")

    If IndRslts(5, i) = "99" Or sShowAge = "n" Then IndRslts(5, i) = "--"
    IndRslts(11, i) = "<a href='certificate.asp?race_id=" & lRaceID & "&event_id=" & lEventID & "&bib=" & IndRslts(1, i) & "' onclick='openThis(this.href,1024,768);return false;'>View</a>"
Next

conn.Close
Set conn = Nothing
%>

{
  "data": [
        <%For i = 0 To UBound(IndRslts, 2)%>
            [
            <%For j = 0 To 11%>
                "<%=IndRslts(j, i)%>"
                <%If j < 11 Then %>
                    <%Response.Write ","       '-- don't output a comma on the last element%>
                  <%End If%>
                <%Next%>
            ]

            <%If i < UBound(IndRslts, 2) Then %>
                <%Response.Write ","         '-- don't output a comma on the last section%>
            <%End If%>
        <%Next%>
]
}
