<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisMeet, lThisRace
Dim sMeetName
Dim Bibs(), Excludes()
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 600

lThisMeet = Request.QueryString("meet_id")
lThisRace = Request.QueryString("this_race")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lThisMeet
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing
   
If Request.Form.Item("submit_restore") = "submit_restore" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Excludes FROM IndRslts WHERE MeetsID = " & lMeetID & " AND Bib = " & Request.Form.Item("excludes")
    rs.Open sql, conn, 1, 2
    rs(0).Value = "y"
    rs.Update
    rs.Close
    Set rs = Nothing

    'update team scores
    Call UpdateTmScores
ElseIf Request.Form.Item("submit_exclude") = "submit_exclude" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Excludes FROM IndRslts WHERE MeetsID = " & lMeetID & " AND Bib = " & Request.Form.Item("excludes")
    rs.Open sql, conn, 1, 2
    rs(0).Value = "n"
    rs.Update
    rs.Close
    Set rs = Nothing

    'update team scores
    Call UpdateTmScores
End If 

i = 0
ReDim Bibs(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND RaceTime <> '00:00' AND Excludes = 'n' ORDER BY Bib"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Bibs(i) = rs(0).Value
    i = i + 1
    ReDim Preserve Bibs(i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
    
i = 0
ReDim Excludes(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Bib FROM IndRslts WHERE MeetsID = " & lThisMeet & " AND RaceTime <> '00:00' AND Excludes = 'y' ORDER BY Bib"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Excludes(i) = rs(0).Value
    i = i + 1
    ReDim Preserve Excludes(i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Private Sub UpdateTmScores()
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<title>GSE CC/Nordic Results Manager: Exclude Finisher</title>
<!--#include file = "../../../includes/meta2.asp" -->



</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<%If Not CLng(lThisMeet) = 0 Then%>
				<!--#include file = "../manage_meet_nav.asp" -->
			<%End If%>

			<h4 class="h4">Results Manager for <%=sMeetName%> on <%=dMeetDate%>:&nbsp;Exclude Finisher</h4>

			<!--#include file = "results_nav.asp" -->

            <p style="margin: 0;padding: 0 0 10px 0;">(Leaves finisher in the database but removes them from team scoring and does not 
            assign them a place.)</p>
		
            <div style="width: 300px;float: left;">
                <h4 style="margin-bottom: 10px;">Bib To Exclude:</h4>
                <form name="exclude_bib" method="post" action="exclude_fnshr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>">
                <select name="excludes" id="excludes" size="20">
                    <%For i = 0 To UBound(Bibs) - 1%>
                        <option value="<%=Bibs(i)%>"><%=Bibs(i)%></option>
                    <%Next%>
                </select>
				<input type="hidden" name="submit_exclude" id="submit_exclude" value="submit_exclude">
				<input type="submit" name="submit1" id="submit1" value="Exclude This">
                </form>
            </div>
		
            <div style="margin-left: 310px;">
                <h4 style="margin-bottom: 10px;">Bibs To Restore:</h4>
                <form name="restore_bib" method="post" action="exclude_fnshr.asp?meet_id=<%=lThisMeet%>&amp;this_race=<%=lThisRace%>">
                <select name="restores" id="restores" size="<%=UBound(Excludes)%>">
                    <%For i = 0 To UBound(Excludes) - 1%>
                        <option value="<%=Excludes(i)%>"><%=Excludes(i)%></option>
                    <%Next%>
                </select>
				<input type="hidden" name="submit_restore" id="submit_restore" value="submit_restore">
				<input type="submit" name="submit2" id="submit2" value="Restore This">
                </form>
            </div>
        </div>
    </div>	
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
