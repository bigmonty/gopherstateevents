<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lPartID, lEventID, lSurveyResultsID
Dim sEventDirEmail, sEventName, sLocation, sRaceName, sSuggestions, sComments, sMsg, sPartName
Dim sWordPhrase, sEventComments, sTimingComments, sGenerallySpeaking, sFinalThoughts, sConsent, sClickPage
Dim iOverall, iScore
Dim cdoMessage, cdoConfig
Dim Prompts(), Scale(1, 3)
Dim dEventDate, dExpDate
Dim bAccessDenied

lRaceID = Request.QueryString("race_id")
lPartID = Request.QueryString("part_id")

If CStr(lRaceID) = vbNullString Then lRaceID = "0"
If CStr(lPartID) = vbNullString Then lpartID = "0"

'If CLng(lRaceID) = 0 Or CLng(lPartID) = 0 Then bAcccessDenied = True

sClickPage = Request.ServerVariables("URL")

Scale(0, 0) = "1"
Scale(1, 0) = "Strongly Disagree"
Scale(0, 1) = "2"
Scale(1, 1) = "Disagree"
Scale(0, 2) = "3"
Scale(1, 2) = "Agree"
Scale(0, 3) = "4"
Scale(1, 3) = "Strongly Agree"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Prompts(2, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SurveyPromptsID, Prompt, PromptType FROM SurveyPrompts ORDER BY PromptType, Prompt"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Prompts(0, i) = rs(0).Value
    Prompts(1, i) = Replace(rs(1).Value, "''", "'")
    Prompts(2, i) = rs(2).Value
    i = i + 1
    ReDim Preserve Prompts(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = sErver.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventID, e.EventName, e.EventDate, e.Location, r.RaceName FROM Events e INNER JOIN RaceData r ON e.EventID = r.EventID WHERE r.RaceID = " 
sql = sql & lRaceID
rs.Open sql, conn, 1, 2
lEventID = rs(0).Value
sEventName = Replace(rs(1).Value, "''", "'")
dEventDate = rs(2).Value
If Not rs(3).Value & "" = "" Then sLocation = Replace(rs(3).Value, "''", "'")
sRaceName = Replace(rs(4).Value, "''", "'")
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ed.Email FROM EventDir ed INNER JOIN Events e ON ed.EventDirID = e.EventDirID WHERE e.EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventDirEmail = rs(0).Value
rs.Close
Set rs = Nothing

dExpDate = CDate(dEventDate) + 7

Set rs = sErver.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
sPartName = Replace(rs(1).Value, "''", "'") & ", " & Replace(rs(0).Value, "''", "'")
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_survey") = "submit_survey" Then
    'insert into survey
    iOverall = Request.Form.Item("overall")
    sConsent = Request.Form.Item("consent")

    sWordPhrase = Request.Form.Item("word_phrase")
    sEventComments = Request.Form.Item("event_comments")
    sTimingComments = Request.Form.Item("timing_comments")
    sFinalThoughts = Request.Form.Item("final_thoughts")
    sGenerallySpeaking = Request.Form.Item("generally_speaking")

    If Not sWordPhrase & "" = "" Then sWordPhrase = Left(Replace(sWordPhrase, "'", "''"), 100)
    If Not sEventComments & "" = "" Then sEventComments = Left(Replace(sEventComments, "'", "''"), 2000)
    If Not sTimingComments & "" = "" Then sTimingComments = Left(Replace(sTimingComments, "'", "''"), 2000)
    If Not sFinalThoughts & "" = "" Then sFinalThoughts = Left(Replace(sFinalThoughts, "'", "''"), 2000)
    If Not sGenerallySpeaking & "" = "" Then sGenerallySpeaking = Left(Replace(sGenerallySpeaking, "'", "''"), 2000)

	sql = "INSERT INTO SurveyResults(EventID, RaceID, ParticipantID, WhenSent, Overall, Consent, WordPhrase, EventComments, TimingComments, FinalThoughts, "
    sql = sql & " GenerallySpeaking) VALUES (" & lEventID & ", " & lRaceID & ", " & lPartID & ", '" & Now() & "', " & iOverall & ", '" & sConsent & "', '" 
    sql = sql & sWordPhrase & "', '" & sEventComments & "', '" & sTimingComments & "', '" & sFinalThoughts & "', '" & sGenerallySpeaking & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    'get survey results id
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SurveyResultsID FROM SurveyResults WHERE EventID = " & lEventID & " AND RaceID = " & lRaceID & " AND ParticipantID = " & lPartID
    sql = sql & " ORDER BY SurveyResultsID DESC"
    rs.Open sql, conn, 1, 2
    lSurveyResultsID = rs(0).Value
    rs.Close
    Set rs = Nothing

    'insert prompt values
    For i = 0 To UBound(Prompts, 2) - 1
        iScore = Request.Form.Item("prompt_" & Prompts(0, i))
        If CStr(iScore) = vbNullString Then iScore = 0

        If Not CInt(iScore) = 0 Then
	        sql = "INSERT INTO SurveyPrmptRslts(SurveyResultsID, SurveyPromptsID, Score) VALUES (" & lSurveyResultsID & ", " & Prompts(0, i) & ", " 
            sql = sql & iScore & ")"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    Next

    'notify us and event director
    sMsg = "Survey Results for " & sEventName & " by " & sPartName & vbCrLf
    sMsg = sMsg & "Race: " & sRaceName & vbCrLf & vbCrLf

    sMsg = sMsg & "Overall Rating (Out of 10): " & iOverall & vbCrLf
    sMsg = sMsg & "In a word or phrase: " & Replace(sWordPhrase, "''", "'") & vbCrLf & vbCrLf

    sMsg = sMsg & "SCALE: " & vbCrLf
    sMsg = sMsg & "1 = Strongly Disagree " & vbCrLf
    sMsg = sMsg & "2 = Disagree " & vbCrLf
    sMsg = sMsg & "3 = Agree " & vbCrLf
    sMsg = sMsg & "4 = Strongly Agree " & vbCrLf & vbCrLf

    For i = 0 To UBound(Prompts, 2) - 1
        sMsg = sMsg & Prompts(1, i) & ": " & Request.Form.Item("prompt_" & Prompts(0, i)) & vbCrLf & vbCrLf
    Next
    sMsg = sMsg & vbCrLf

    sMsg = sMsg & "Event Comments: " & Replace(sEventComments, "''", "'") & vbCrLf & vbCrLf
    sMsg = sMsg & "Timing Comments: " & Replace(sTimingComments, "''", "'") & vbCrLf & vbCrLf
    sMsg = sMsg & "Generally Speaking: " & Replace(sGenerallySpeaking, "''", "'") & vbCrLf & vbCrLf
    sMsg = sMsg & "Final Thoughts: " & Replace(sFinalThoughts, "''", "'") & vbCrLf & vbCrLf
%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.com;" & sEventDirEmail
'		.To = "bob.schneider@gopherstateevents.com;"
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "GSE Survey for " & sEventName
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
End If

bAccessDenied = False
'check for prior submission by this participant
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SurveyResultsID FROM SurveyResults WHERE EventID = " & lRaceID & " AND ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then bAccessDenied = True
rs.Close
Set rs = Nothing

If Date > CDate(dExpDate) Then bAccessDenied = True
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Race Experience Survey</title>
<meta name="description" content="Gopher State Events (GSE) event survey page.">
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-md-6">
            <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
        </div>
        <div class="col-md-6">
	       <h3 class="h3">Gopher State Events Race Experience Survey</h3>
        </div>
    </div>
	        
    <%If bAccessDenied = True Then%>
        <div class="row">
            <h3 class="h3">Gopher State Events Race Experience Survey</h3>
            <div class="bg-danger">
                We're sorry.  Either the window for returning a survey for this event has closed or you have already submitted a response for this event.
                If you have not submitted a response but the survey window has closed we would still love to hear your thoughts, 
                good or bad.  Please send any comments to <a href="mailto:bob.schneider@gopherstateevents.com">Gopher State Events, LLC</a>.
            </div>
        </div>
    <%Else%>
        <div class="row">
            
            <div class="bg-warning">
                At GSE we care about your experience.  Please feel free to comment on any or all of the following prompts.  Note that all prompts are 
                phrased positively so that you can "Strongly Disagree", "Disagree", "Agree" or "Strongly Agree".  Only one response will be accepted per 
                participant.  NOTE: The survey for this event will expire on <%=dExpDate%>.  This survey will be recorded in reference to:
            </div>

            <ul class="list-inline"  style="margin-top: 10px;">
                <li><span style="font-weight: bold;">Event:</span> <%=sEventName%></li>
                <li><span style="font-weight: bold;">Date:</span> <%=dEventDate%></li>
                <li><span style="font-weight: bold;">Location:</span> <%=sLocation%></li>
                <li><span style="font-weight: bold;">Participant:</span> <%=sPartName%></li>
                <li><span style="font-weight: bold;">Race:</span> <%=sRaceName%></li>
            </ul>
        </div>

        <form class="form" role="form" name="race_survey" method="post" action="race_survey.asp?race_id=<%=lRaceID%>&amp;part_id=<%=lPartID%>">
        <div class="row">
            <div class="form-group">
                <label for="word_phrase">In a word...or phrase...</label>
                <input type="text" class="form-control" name="word_phrase" id="word_phrase" maxlength="100">
            </div>
        </div>

        <div class="row">
            <div class="col-sm-6 bg-warning" style="padding: 10px;">
                <h4 class="h4">First, let's talk about the event itself...</h4>

                <%k = 0%>
                <%For i = 0 To UBound(Prompts, 2) - 1%>
                    <%If Prompts(2, i) = "Event" Then%>
                        <div class="form-group">
                            <%=k + 1%>) <%=Prompts(1, i)%>
                            <br>
                            <%For j = 0 to UBound(Scale, 2)%>
                                <label  class="radio-inline" style="margin-left: 25px;">
                                    <input type="radio" name="prompt_<%=Prompts(0, i)%>" id="prompt_<%=Prompts(0, i)%>" 
                                        value="<%=Scale(0, j)%>"><%=Scale(1, j)%>
                                </label>
                            <%Next%>
                        </div>

                        <%k = k + 1%>
                    <%End If%>
                <%Next%>

                <div class="form-group">
                    <label for="event_comments"> Other comments about the event itself:</label>
                    <textarea class="form-control" name="event_comments" id="event_comments"></textarea>
                </div>
            </div>
            <div class="col-sm-6 bg-success" style="padding: 10px;">
                <h4 class="h4">Lets not forget the timers...</h4>

                <%k = 0%>
                <%For i = 0 To UBound(Prompts, 2) - 1%>
                    <%If Prompts(2, i) = "Timing" Then%>
                        <div class="form-group">
                            <%=k + 1%>) <%=Prompts(1, i)%>
                            <br>
                            <%For j = 0 to UBound(Scale, 2)%>
                                <label  class="radio-inline" style="margin-left: 25px;">
                                    <input type="radio" name="prompt_<%=Prompts(0, i)%>" id="prompt_<%=Prompts(0, i)%>" 
                                        value="<%=Scale(0, j)%>"><%=Scale(1, j)%>
                                </label>
                            <%Next%>
                        </div>

                        <%k = k + 1%>
                    <%End If%>
                <%Next%>
                
                <div class="form-group">    
                    <label for="timing_comments"> Other comments about the timing:</label>
                    <textarea class="form-control" name="timing_comments" id="timing_comments"></textarea>
                </div>
            </div>
        </div>

        <div class="row">
            <div class="form-group">
                <label for="generally_speaking">Generally speaking...</label>
                <textarea class="form-control" name="generally_speaking" id="generally_speaking"></textarea>
            </div>
        </div>
        <div class="row">
            <div class="form-group">
                <label for="final_thoughts">Final Thoughts...</label>
                <textarea class="form-control" name="final_thoughts" id="final_thoughts"></textarea>
            </div>
        </div>
        <div class="row">
            <div class="form-group">
                <label for="overall">Overall, I would rate my experience in this event...</label>
                Very Unfulfilling&nbsp;&nbsp;&nbsp;
                <%For i = 1 To 10%>
                    <input type="radio" name="overall" id="overall" value="<%=i%>">&nbsp;<%=i%>&nbsp;&nbsp;&nbsp;
                <%Next%>
                Very Fulfilling
            </div>
        </div>
        <div class="row">
            <div class="form-group">
                <p>You 
                    &nbsp;&nbsp;<input type="radio" name="consent" id="consent" value="y" checked>&nbsp;<span style="font-weight: bold;">Do</span>
                    <input type="radio" name="consent" id="consent" value="n">&nbsp;<span style="font-weight: bold;">Do Not</span>
                    &nbsp;&nbsp;have my consent to use any or all of my survey results for promotional purposes as long as only my first name and last initial 
                    are used.
                </p>
                <input type="hidden" name="submit_survey" id="submit_survey" value="submit_survey">
                <input type="submit" name="submit1" id="submit1" value="Submit Survey">
            </div>
        </div>
        </form>
    <%End If%>
   	<!--#include file = "../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>