<%@ Language=VBScript%>
<%
Option Explicit
%>

<!--#include file = "../../../includes/JSON_2.0.4.asp" -->

<%
Dim conn, rs, sql
Dim i, j
Dim lRaceID
Dim sGender, sSortBy, sSortDir, sShowLength, sFirstName, sLastName
Dim iLength, iAge
Dim Results

lRaceID = Request.QueryString("series_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect "http://www.google.com"
If CLng(lRaceID) < 0 Then Response.Redirect "http://www.google.com"

sGender = Request.QueryString("gender")
If sGender = vbNullString Then sGender = "M"
If Len(sGender) > 1 Then Response.Redirect "http://www.google.com"

iAge = Request.QueryString("age")
If CStr(iAge) = vbNullString Then iAge = 0

sFirstName = Request.QueryString("first_name")
If sFirstName = "undefined" Then sFirstName = vbNullString

sLastName = Request.QueryString("last_name")
If sLastName = "undefined" Then sLastName = vbNullString

iLength = Request.QueryString("results_length")
If CStr(iLength) = "undefined" Then iLength = 100
If CStr(iLength) = vbNullString Then iLength = 100

sSortBy = Request.QueryString("results_sort")
sSortDir = Request.QueryString("results_sort_direction")
'sStdgsPage = Request.QueryString("results_page")

If CStr(iLength) = "-1" Then
    sShowLength = vbNullString
Else
    sShowLength = " TOP " & iLength
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TOP " & iLength & " PR.Bib, P.FirstName, P.LastName, P.Gender, PR.Age, IR.ChipTime, IR.FnlTime, IR.ChipStart, P.City, P.St "
sql = sql & "FROM RaceData AS R INNER JOIN PartRace AS PR ON R.RaceID = PR.RaceID INNER JOIN Participant AS P ON PR.ParticipantID = P.ParticipantID "
sql = sql & "INNER JOIN (SELECT DISTINCT RaceID, ParticipantID, ChipTime, FnlTime, ChipStart, FnlScnds FROM IndResults WHERE FnlTime IS NOT NULL "
sql = sql & "AND FnlTime > '00:00:00.000') AS IR ON R.RaceID = IR.RaceID AND P.ParticipantID = IR.ParticipantID WHERE R.RaceID = " & lRaceID
sql = sql & " ORDER BY " & sSortBy & " " & sSortDir
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    Results = rs.GetRows()
Else
    ReDim Results(9, 0)
End If
rs.Close
Set rs = Nothing

For i = 0 To UBound(Results, 2)
    Results(0, i) = i + 1
Next

conn.Close
Set conn = Nothing
%>

{
  "data": [
        <%For i = 0 To UBound(Results, 2)%>
            [
            <%For j = 0 To 9%>
                "<%=Results(j, i)%>"
                <%If j < 9 Then %>
                    <%Response.Write ","       '-- don't output a comma on the last element%>
                  <%End If%>
                <%Next%>
            ]

            <%If i < UBound(Results, 2) Then %>
                <%Response.Write ","         '-- don't output a comma on the last section%>
            <%End If%>
        <%Next%>
]
}
