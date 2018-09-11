<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i
Dim lPartID, lCellProvider, lEventID
Dim iMyBib
Dim sMobileNumber, sErrMsg, sEventRaces, sEventName
Dim CellProviders, Events
Dim cdoMessage, cdoConfig
Dim bFound, bSuccess

lEventID = Request.QueryString("event_id")
If CStr(lEventID) & "" = "" Then lEventID = "0"
If Not IsNumeric(lEventID) Then Response.Redirect "http://www.google.com"
If CLng(lEventID) < 0 Then Response.Redirect "http://www.google.com"

iMyBib = Request.QueryString("my_bib")
If CStr(iMyBib) & "" = "" Then iMyBib = "0"
If Not IsNumeric(iMyBib) Then Response.Redirect "http://www.google.com"
If CLng(iMyBib) < 0 Then Response.Redirect "http://www.google.com"

lPartID = Request.QueryString("part_id")
If CStr(lPartID) & "" = "" Then lPartID = "0"
If Not IsNumeric(lPartID) Then Response.Redirect "http://www.google.com"
If CLng(lPartID) < 0 Then Response.Redirect "http://www.google.com"

bSuccess = False

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '" & Date - 6 & "' ORDER By EventDate"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn2, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%

If Request.Form.Item("submit_this") = "submit_this" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'sms_kiosk' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_sms_kiosk") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_sms_kiosk") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
		
		iMyBib = CleanInput(Trim(Request.Form.Item("my_bib")))
        If sHackMsg = vbNullString Then lEventID = CleanInput(Trim(Request.Form.Item("event_id")))
		If sHackMsg = vbNullString Then sMobileNumber = CleanInput(Trim(Request.Form.Item("mobile_number")))
		If sHackMsg = vbNullString Then lCellProvider = CleanInput(Trim(Request.Form.Item("cell_provider")))
		
		If sHackMsg = vbNullString Then
            sMobileNumber = Replace(sMobileNumber, "-", "")
            sMobileNumber = Replace(sMobileNumber, ".", "")
            sMobileNumber = Replace(sMobileNumber, "(", "")
            sMobileNumber = Replace(sMobileNumber, ")", "")
            sMobileNumber = Replace(sMobileNumber, " ", "")
            sMobileNumber = Trim(sMobileNumber)

            'get part_id if missing
            If CLng(lPartID) = 0 Then
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

                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT ParticipantID FROM PartRace WHERE Bib = " & iMyBib & " AND RaceID IN (" & sEventRaces & ")"
                rs.Open sql, conn, 1, 2
                lPartID = rs(0).Value
                rs.Close
                Set rs = Nothing
            End If

            bFound = False
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MobileNumber, CellProvider, PartID FROM MobileSettings WHERE Bib = " & iMyBib & " AND EventID = " & lEventID
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then
                rs(0).Value = sMobileNumber
                rs(1).Value = lCellProvider
                If rs(2).Value = "0" Then rs(2).Value = lPartID
                rs.Update
                bFound = True
            End If
            rs.Close
            Set rs = Nothing

            If bFound = False Then
                sql = "INSERT INTO MobileSettings (Bib, EventID, MobileNumber, CellProvider, PartID) VALUES (" & iMyBib & ", " & lEventID & ", '" 
                sql = sql & sMobileNumber & "', " & lCellProvider & ", " & lPartID & ")"
                Set rs = conn.execute(sql)
                Set rs = Nothing
            End If

            sql = "SELECT EventName FROM Events WHERE EventID = " & lEventID
            Set rs = conn.Execute(sql)
            sEventName = Replace(rs(0).Value, "''", "'")
            Set rs = Nothing

            'send test sms to this person
            Set cdoMessage = Server.CreateObject("CDO.Message")
		    With cdoMessage
                Set .Configuration = cdoConfig
                .From = "bob.schneider@gopherstateevents.com"
			    .To = sMobileNumber & GetSendURL(lCellProvider)
			    .TextBody = "Welcome to Gopher State Events SMS Race Results.  Good luck at the races."
			    .Send
		    End With
	        Set cdoMessage = Nothing

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = "bob.schneider@gopherstateevents.com"
				.From = "bobs@h51softtware.net"
				.Subject = "New SMS Results Registration"
				.TextBody = "Bib " & iMyBib & " in " & sEventName & " has just signed up to receive thier results via text messaging."
				.Send
			End With
			Set cdoMessage = Nothing

            bSuccess = True
        End If
   End If
End If

'log this user if they are just entering the site
If Session("access_sms_kiosk") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'sms_kiosk')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'sms_kiosk'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_sms_kiosk")
End If

Private Function GetSendURL(lProviderID)
	If Not CStr(lProviderID) & "" = ""  Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT SendURL FROM CellProviders WHERE CellProvidersID = " & lProviderID
		rs.Open sql, conn2, 1, 2
		If rs.RecordCount > 0 Then GetSendURL = rs(0).Value
        rs.Close
		Set rs = Nothing
	End If
End Function

%>
<!--#include file = "../includes/clean_input.asp" -->
<%

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; SMS Results Kiosk</title>
<meta name="description" content="SMS Results Set-Up a Gopher State Events (GSE) timed event.">
<!--#include file = "../includes/js.asp" --> 

<script>
    function chkFlds() {
        if (document.sms_info.mobile_number.value == '' ||
        document.sms_info.cell_provider.value == '' ||
        document.sms_info.event_id.value == '' ||
        document.sms_info.my_bib.value == '') {
            alert('All fields are required!');
            return false;
        }
        else
		    if (isNaN(document.sms_info.mobile_number.value) ||
                isNaN(document.sms_info.my_bib.value))
    		    {
			    alert('The mobile number and bib can not contain non-numeric values');
			    return false
			    }
        else
            return true;
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
            <h1 class="h1">SMS Results Input</h1>
        </div>
    </div>
    <div class="bg-danger">
        <a href="sms_kiosk.asp">Refresh Page</a>
    </div>

    <%If bSuccess = True Then%>
        <div class="bg-info">
            <h4 class="h4">Success!</h4>

            <p>
                You are now set up to receive your results for this race by text message.  You should be receiving a confirmation text very soon.  "
                PLEASE REPLY "Got It" to that message so we know you received it.
            </p>
        </div>
    <%Else%>
        <div>
            <h4 class="h4">Enter Mobile Data</h4>
            <p>
                In order to send you your results via text message, we need you to identify the event you are participating in, your bib number, your 
                mobile number & provider.  We will not use this information for anything other than sending YOU your results.  By submitting this 
                information you are agreeing to allow us to do this.
            </p>
            <p>
                Is your provider not listed?  Encountered problems using this page?  Let us know about it 
                <a href="contact.asp" style="font-weight: bold;color: red;">here.</a>
            </p>

            <%If Not sErrMsg = vbNullString Then%>
                <div class="bg-danger"><%=sErrMsg%></div>
            <%End If%>
                <form class="form-horizontal" role="form" name="sms_info" method="post" 
                    action="sms_kiosk.asp?part_id=<%=lPartID%>&amp;my_bib=<%=iMyBib%>&amp;event_id=<%=lEventID%>" onsubmit="return chkFlds();">
                <div class="form-group">
 	                <label for="event_id" class="control-label col-xs-3">Event:</label>
                    <div class="col-xs-9">
                        <select class="form-control" name="event_id" id="event_id">
                            <option value="">&nbsp;</option>
                            <%For i = 0 To UBound(Events, 2)%>
                                <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
                                    <option value="<%=Events(0, i)%>" selected><%=Replace(Events(1, i), "''", "'")%> (<%=Events(2, i)%>)</option>
                                <%Else%>
                                    <option value="<%=Events(0, i)%>"><%=Replace(Events(1, i), "''", "'")%> (<%=Events(2, i)%>)</option>
                                <%End If%>
                            <%Next%>
                        </select>
                    </div>
                </div>
                <div class="form-group">
 	                <label for="my_bib" class="control-label col-xs-3">Bib #:</label>
                    <div class="col-xs-9"><input class="form-control" type="text" name="my_bib" id="my_bib" value="<%=iMyBib%>"></div>
                </div>
                <div class="form-group">
	               <label for="mobile_number" class="control-label col-xs-3">Mobile Phone</label>
                    <div class="col-xs-9"><input class="form-control" type="text" name="mobile_number" id="mobile_number" value="<%=sMobileNumber%>"></div>
                </div>
                <div class="form-group">
                    <label for="cell_provider" class="control-label col-xs-3">Provider:</label>
                    <div class="col-xs-9">
                        <select class="form-control" name="cell_provider" id="cell_provider">
                            <option value="">&nbsp;</option>
                            <%For i = 0 To UBound(CellProviders, 2)%>
                                <%If CLng(lCellProvider) = CLng(CellProviders(0, i)) Then%>
                                    <option value="<%=CellProviders(0, i)%>" selected><%=CellProviders(1, i)%></option>
                                <%Else%>
                                    <option value="<%=CellProviders(0, i)%>"><%=CellProviders(1, i)%></option>
                                <%End If%>
                            <%Next%>
                        </select>
                    </div>
                </div>
                <div class="form-group">
		            <input class="form-control" type="hidden" name="submit_this" id="submit_this" value="submit_this">
		            <input class="form-control" type="submit" name="submit1" id="submit1" value="Submit This">
                </div>
                </form>
            </div>
        </div>
    <%End If%>
</div>
</body>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</html>
