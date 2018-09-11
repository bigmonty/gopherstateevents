<%@ Language=VBScript%>
<%
Option Explicit
%>

<!--#include file = "../../includes/JSON_2.0.4.asp" -->

<%
Dim conn, rs, sql, rs2, sql2
Dim i, j
Dim lSeriesID
Dim Results, Races()

Server.ScriptTimeout = 1200

lSeriesID = Request.QueryString("series_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

j = 0
ReDim Races(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT sr.RacesID FROM CCSeriesRaces sr INNER JOIN CCSeriesMeets se ON sr.CCSeriesMeetsID = se.CCSeriesMeetsID "
sql = sql & "WHERE se.CCSeriesID = " & lSeriesID & " AND se.MeetDate < '" & Date & "' ORDER BY se.MeetDate"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Races(j) = rs(0).Value
    j = j + 1
    ReDim Preserve Races(j)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT sp.RosterID, sp.PartName, sp.School, sr.Score FROM CCSeriesParts sp INNER JOIN CCSeriesResults sr ON sp.CCSeriesPartsID = sr.CCSeriesPartsID "
sql = sql & "WHERE sp.CCSeriesID = " & lSeriesID & " ORDER BY sr.Score DESC"
rs.Open sql, conn, 1, 2
Results = rs.GetRows()
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
        <%For i = 0 To UBound(Results, 2) - 1%>
            [
            <%For j = 0 To 3%>
                "<%=Results(j, i)%>"
                <%If j < 3 Then %>
                    <%Response.Write ","       '-- don't output a comma on the last element%>
                  <%End If%>
                <%Next%>
            ]

            <%If i < UBound(Results, 2) - 1 Then %>
                <%Response.Write ","           '-- don't output a comma on the last section%>
            <%End If%>
        <%Next%>
]
}
