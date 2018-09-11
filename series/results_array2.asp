<%@ Language=VBScript%>
<%
Option Explicit
%>

<!--#include file = "../includes/JSON_2.0.4.asp" -->

<%
Dim conn, rs, sql
Dim i, j
Dim lSeriesID
Dim iLength
Dim sGender, sFilter, sSortBy, sSortDir, sStdgsPage, sShowLength
Dim Results

lSeriesID = Request.QueryString("series_id")
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"
If CLng(lSeriesID) < 0 Then Response.Redirect "http://www.google.com"

sGender = Request.QueryString("gender")
If sGender = vbNullString Then sGender = "M"
If Len(sGender) > 1 Then Response.Redirect "http://www.google.com"

sFilter = Request.QueryString("standings_filter")
If sFilter = "undefined" Then sFilter = vbNullString

iLength = Request.QueryString("standings_length")
If CStr(iLength) = "undefined" Then iLength = 100
If CStr(iLength) = vbNullString Then iLength = 100

sSortBy = Request.QueryString("standings_sort")
sSortDir = Request.QueryString("standings_sort_direction")
sStdgsPage = Request.QueryString("standings_page")

If sSortBy = "Pts" Then sSortBy = "AgeScore"

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
sql = "SELECT" & sShowLength & " sp.ParticipantID, sp.PartName, sp.Age, sr.AgeScore FROM SeriesParts sp INNER JOIN SeriesResults sr "
sql = sql & "ON sp.SeriesPartsID = sr.SeriesPartsID WHERE sp.PartName LIKE '%" & sFilter & "%' AND sp.SeriesID = " & lSeriesID 
sql = sql & " AND sp.Gender = '" & sGender & "' ORDER BY " & sSortBy & " " & sSortDir
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    Results = rs.GetRows()
Else
    ReDim Results(3, 0)
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
            <%For j = 0 To 3%>
                "<%=Results(j, i)%>"
                <%If j < 3 Then %>
                    <%Response.Write ","       '-- don't output a comma on the last element%>
                  <%End If%>
                <%Next%>
            ]

            <%If i < UBound(Results, 2) Then %>
                <%Response.Write ","           '-- don't output a comma on the last section%>
            <%End If%>
        <%Next%>
]
}
