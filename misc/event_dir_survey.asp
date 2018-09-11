<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lEventID, lEventSurveyResultsID
Dim sEventDirEmail, sEventName, sMsg, sSentBy, sWordPhrase, sFinalThoughts, sConsent, sRole
Dim iScore, iPricing, iExpectations
Dim cdoMessage, cdoConfig
Dim Prompts(), Scale(1, 3)
Dim dEventDate, dExpDate
Dim bAccessDenied

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = "0"
If Not IsNumeric(lEventID) Then Response.Redirect "http://www.google.com"
If CLng(lEventID) < 0  Then Response.Redirect "http://www.google.com"
If CLng(lEventID) = 0 Then bAcccessDenied = True

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
ReDim Prompts(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventSurveyPromptsID, Prompt FROM EventSurveyPrompts ORDER BY Prompt"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Prompts(0, i) = rs(0).Value
    Prompts(1, i) = Replace(rs(1).Value, "''", "'")
    i = i + 1
    ReDim Preserve Prompts(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ed.Email, e.EventName, e.EventDate FROM EventDir ed INNER JOIN Events e ON ed.EventDirID = e.EventDirID WHERE e.EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventDirEmail = rs(0).Value
sEventName = Replace(rs(1).Value, "''", "'")
dEventDate = rs(2).Value
rs.Close
Set rs = Nothing

bAccessDenied = False
dExpDate = CDate(dEventDate) + 45
If Date > CDate(dExpDate) Then bAccessDenied = True

If Request.Form.Item("submit_survey") = "submit_survey" Then
    sConsent = Request.Form.Item("consent")
    sWordPhrase = Request.Form.Item("word_phrase")
    sFinalThoughts = Request.Form.Item("final_thoughts")
    sSentBy = Replace(Request.Form.Item("sent_by"), "''", "'")
    sRole = Request.Form.Item("role")
    iPricing = Request.Form.Item("pricing")
    iExpectations = Request.Form.Item("expectations")

    If Not sWordPhrase & "" = "" Then sWordPhrase = Left(Replace(sWordPhrase, "'", "''"), 100)
    If Not sFinalThoughts & "" = "" Then sFinalThoughts = Left(Replace(sFinalThoughts, "'", "''"), 2000)

	sql = "INSERT INTO EventSurveyResults(EventID, SentBy, WhenSent, Consent, WordPhrase, FinalThoughts, Pricing, Role, "
    sql = sql & "Expectations) VALUES ("  & lEventID & ", '"  & sSentBy & "', '" & Now() & "', '" & sConsent & "', '" & sWordPhrase 
    sql = sql & "', '" & sFinalThoughts & "', " & iPricing & ", '" & sRole & "', " & iExpectations & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    'get survey results id
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventSurveyResultsID FROM EventSurveyResults WHERE EventID = " & lEventID & " ORDER BY EventSurveyResultsID DESC"
    rs.Open sql, conn, 1, 2
    lEventSurveyResultsID = rs(0).Value
    rs.Close
    Set rs = Nothing

    'insert prompt values
    For i = 0 To UBound(Prompts, 2) - 1
        iScore = Request.Form.Item("prompt_" & Prompts(0, i))
        If CStr(iScore) = vbNullString Then iScore = 0

        If Not CInt(iScore) = 0 Then
	        sql = "INSERT INTO EventSurveyPrmptRslts(EventSurveyResultsID, EventSurveyPromptsID, Score) VALUES (" & lEventSurveyResultsID & ", " & Prompts(0, i) 
            sql = sql & ", " & iScore & ")"
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    Next

    'notify us
    sMsg = "Event Survey Results for " & sEventName & vbCrLf
    sMsg = "Sent By: " & sSentBy & vbCrLf & vbCrLf

    sMsg = sMsg & "OPEN ENDED RESPONSES: " & vbCrLf & vbCrLf
    sMsg = sMsg & "In a word or phrase: " & Replace(sWordPhrase, "''", "'") & vbCrLf
    sMsg = sMsg & "Expectations (1=Not Met, 5=Met): " & iExpectations & vbCrLf
    sMsg = sMsg & "Pricing (1=Low, 5=High): " & iPricing & vbCrLf & vbCrLf
    sMsg = sMsg & "Consent to Use for Promo: " & sConsent & vbCrLf
    sMsg = sMsg & "Final Thoughts: " & vbCrLf & Replace(sFinalThoughts, "''", "'") & vbCrLf & vbCrLf

    sMsg = sMsg & "QUICK RESPONSE SCALE: " & vbCrLf
    sMsg = sMsg & "1 = Strongly Disagree " & vbCrLf
    sMsg = sMsg & "2 = Disagree " & vbCrLf
    sMsg = sMsg & "3 = Agree " & vbCrLf
    sMsg = sMsg & "4 = Strongly Agree " & vbCrLf & vbCrLf

    sMsg = sMsg & "QUICK RESPONSES: " & vbCrLf & vbCrLf
    For i = 0 To UBound(Prompts, 2) - 1
        sMsg = sMsg & Prompts(1, i) & ": " & Request.Form.Item("prompt_" & Prompts(0, i)) & vbCrLf & vbCrLf
    Next

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.com;"
		.From = "bob.schneider@gopherstateevents.com"
		.Subject = "GSE Event Director Survey for " & sEventName
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Event Director Survey</title>
<meta name="description" content="Gopher State Events (GSE) event survey page.">
<!--#include file = "../includes/js.asp" -->

<script>
function chkFlds() {
if (document.evnt_dir_survey.sent_by.value == '') 
{
 	alert('Please submit your name!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body>
<div class="container">
    <img class="img-responsive" src="http://www.gopherstateevents.com/graphics/html_header.png" alt="GSE Header">
	        
    <%If bAccessDenied = True Then%>
        <div class="row">
            <h3 class="h3">GSE Event Director Survey</h3>
            <div class="bg-danger">
                We're sorry.  The window for returning a survey for this event has closed.
                If you have not submitted a response but the survey window has closed we would still love to hear your thoughts, 
                good or bad (although we prefer good :)).  Please send any comments to <a href="mailto:bob.schneider@gopherstateevents.com">Gopher State Events, LLC</a>.
            </div>
        </div>
    <%Else%>
        <div class="row">
            <h3 class="h3">GSE Event Director Survey: <%=sEventName%> on <%=dEventDate%></h3>
            <div class="bg-success">
                At GSE we care about your experience.  Please feel free to comment on any or all of the following prompts.  Note that all prompts are 
                phrased positively so that you can "Strongly Disagree", "Disagree", "Agree" or "Strongly Agree".  The survey for this event will 
                expire on <%=dExpDate%>.  <span class="text-danger">You may forward this link to any members of your staff that you would like to submit
                input.</span>
            </div>
        </div>

        <form class="form" role="form" name="evnt_dir_survey" method="post" action="event_dir_survey.asp?event_id=<%=lEventID%>" 
            onSubmit="return chkFlds();">
        <div class="row">
            <div class="col-sm-6">
                <div class="form-group">
                    <label for="sent_by">My Name:</label>
                    <input type="text" class="form-control" name="sent_by" id="sent_by" maxlength="50">
                </div>

                <div class=" bg-warning">
                    <h4 class="h4">Quick Response</h4>

                    <%For i = 0 To UBound(Prompts, 2) - 1%>
                        <div class="form-group">
                            <%=i + 1%>) <%=Prompts(1, i)%>
                            <br>
                            <%For j = 0 to UBound(Scale, 2)%>
                                <label  class="radio-inline" style="margin-left: 25px;">
                                    <input type="radio" name="prompt_<%=Prompts(0, i)%>" id="prompt_<%=Prompts(0, i)%>" 
                                        value="<%=Scale(0, j)%>"><%=Scale(1, j)%>
                                </label>
                            <%Next%>
                        </div>
                    <%Next%>
                </div>
            </div>
            <div class="col-sm-6">
                <div class="form-group">
                    <label for="word_phrase">My role in this event is...</label>
                    <input type="text" class="form-control" name="role" id="role" maxlength="50">
                </div>

                <div class="form-group">
                    <label for="word_phrase">In a word or a phrase...</label>
                    <input type="text" class="form-control" name="word_phrase" id="word_phrase" maxlength="100">
                </div>

                <div class="form-group">
                    <label for="final_thoughts">Final Thoughts...</label>
                    <textarea class="form-control" name="final_thoughts" id="final_thoughts"></textarea>
                </div>

                <div class="form-group text-danger">
                    <label for="overall"> Overall, our expectations of Gopher State Events, LLC were...</label><br>
                    <span style="font-weight: bold;">Not Met</span>&nbsp;&nbsp;&nbsp;
                    <%For i = 1 To 5%>
                        <input type="radio" name="expectations" id="expectations" value="<%=i%>">&nbsp;<%=i%>&nbsp;&nbsp;&nbsp;
                    <%Next%>
                    <span style="font-weight: bold;">Met</span>
                </div>

                <div class="form-group text-warning">
                    <label for="overall"> The pricing for Gopher State Events, LLC services are...</label><br>
                    <span style="font-weight: bold;">Low</span>&nbsp;&nbsp;&nbsp;
                    <%For i = 1 To 5%>
                        <input type="radio" name="pricing" id="pricing" value="<%=i%>">&nbsp;<%=i%>&nbsp;&nbsp;&nbsp;
                    <%Next%>
                    <span style="font-weight: bold;">High</span>
                </div>
            
                <div class="form-group bg-info">
                    You &nbsp;&nbsp;<input type="radio" name="consent" id="consent" value="y" checked>&nbsp;<span style="font-weight: bold;">Do</span>
                    <input type="radio" name="consent" id="consent" value="n">&nbsp;<span style="font-weight: bold;">Do Not</span>
                    &nbsp;&nbsp;have my consent to use any or all of my survey results for promotional purposes as long as only my first name and last 
                    initial are used.
                </div>

                <div class="form-group bg-success text-center">
                    <input type="hidden" name="submit_survey" id="submit_survey" value="submit_survey">
                    <input type="submit" name="submit1" id="submit1" value="Submit Survey">
                </div>
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