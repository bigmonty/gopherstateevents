<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID
Dim Events, SurveyResults, SurveyPrompts

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate BETWEEN '11/25/2014' AND '" & Date & "' ORDER BY EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

i = 0
ReDim SurveyPrompts(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SurveyPromptsID, PromptType, Prompt FROm SurveyPrompts ORDER BY PromptType, Prompt"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    SurveyPrompts(0, i) = rs(0).Value
    SurveyPrompts(1, i) = rs(1).Value
    SurveyPrompts(2, i) = Replace(rs(2).Value, "''", "'")
    i = i + 1
    ReDim Preserve SurveyPrompts(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

If Not CLng(lEventID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SurveyResultsID, RaceID, ParticipantID, WhenSent, Consent, Overall, WordPhrase, GenerallySpeaking, EventComments, TimingComments, "
    sql = sql & "FinalThoughts FROM SurveyResults WHERE EventID = " & lEventID & " ORDER BY RaceID, WhenSent DESC"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        SurveyResults = rs.GetRows()
    Else
        ReDim SurveyResults(10, 0)
    End If
    rs.Close
    Set rs = Nothing
End If

Private Function GetPartName(lPartID)
    GetPartName = "Undetermined"

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        GetPartName = Replace(rs(0).Value, "''", "''") & " " & Replace(rs(1).Value, "''", "''")
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function GetRaceName(lRaceID)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        GetRaceName = Replace(rs(0).Value, "''", "''")
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function GetResponse(lPromptID, lSurveyResultsID)
    GetResponse = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Score FROM SurveyPrmptRslts WHERE SurveyPromptsID = " & lPromptID & " AND SurveyResultsID = " & lSurveyResultsID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        GetResponse = rs(0).Value
    End If
    rs.Close
    Set rs = Nothing
End Function%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>Gopher State Events Participant Surveys</title>
<meta name="description" content="GSE (Gopher State Events) Participant Surveys">

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
<!--
li{
	padding-top: 5px;
	}
    
ul{
    font-size: 0.9em;
    }
-->
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">GSE Participant Surveys</h4>

		    <div style="margin-left:10px;font-size:0.85em;">	
			    <form name="which_event" method="post" action="part_surveys.asp?event_id=<%=lEventID%>">
			    <span style="font-weight:bold;">Event:</span>
			    <select name="events" id="events" onchange="this.form.get_event.click()">
				    <option value="">&nbsp;</option>
				    <%For i = 0 to UBound(Events, 2)%>
					    <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						    <option value="<%=Events(0, i)%>" selected><%=Replace(Events(1, i), "''", "'")%>&nbsp;(<%=Events(2, i)%>)</option>
					    <%Else%>
						    <option value="<%=Events(0, i)%>"><%=Replace(Events(1, i), "''", "'")%>&nbsp;(<%=Events(2, i)%>)</option>
					    <%End If%>
				    <%Next%>
			    </select>
			    <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			    <input type="submit" name="get_event" id="get_event" value="Get These">
			    </form>
		    </div>

            <%If Not CLng(lEVentID) = 0 Then%>
                <%If UBound(SurveyResults, 2) > 0 Then%>
                    <ul style="padding: 10px 0 0 25px;">
                        <%For i = 0 To UBound(SurveyResults, 2)%>
                            <li>
                                <span style="font-weight:bold;"><%=GetPartName(SurveyResults(2, i))%></span> Sent&nbsp;<%=SurveyResults(3, i)%>
                                <ul>
                                    <li>Race: &nbsp;<%=GetRaceName(SurveyResults(1, i))%></li>
                                    <li>Consent: &nbsp;<%=SurveyResults(4, i)%></li>
                                    <li>Overall: &nbsp;<%=SurveyResults(5, i)%></li>
                                    <li>WordPhrase: &nbsp;<%=SurveyResults(6, i)%></li>
                                    <li>GenerallySpeaking: &nbsp;<%=SurveyResults(7, i)%></li>
                                    <li>EventComments: &nbsp;<%=SurveyResults(8, i)%></li>
                                    <li>TimingComments: &nbsp;<%=SurveyResults(9, i)%></li>
                                    <li>FinalThoughts: &nbsp;<%=SurveyResults(10, i)%></li>
                                </ul>
                                <span style="font-weight:bold;font-size: 0.85em;margin: 10px 0 0 15px;">Survey Prompts (1 to 4):</span>
                                <ul style="margin-left: 25px;background-color:#ececd8;">
                                    <%For j = 0 To UBound(SurveyPrompts, 2) - 1%>
                                        <li style="padding-top: 0;"><%=SurveyPrompts(2, j)%>&nbsp;(<%=SurveyPrompts(1, j)%>):
                                            <%=GetResponse(SurveyPrompts(0, j), SurveyResults(0, i))%></li>
                                    <%Next%>
                                </ul>
                            </li>
                        <%Next%>
                    </ul>
                <%End If%>
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