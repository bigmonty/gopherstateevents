<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i
Dim lRosterID, lCellProvider, lMeetsID
Dim iMyBib
Dim sMobileNumber, sErrMsg, sMeetRaces, sMeetName
Dim CellProviders, Meets
Dim cdoMessage, cdoConfig
Dim bFound, bSuccess

lMeetsID = Request.QueryString("meet_id")
If CStr(lMeetsID) & "" = "" Then lMeetsID = "0"
If Not IsNumeric(lMeetsID) Then Response.Redirect "http://www.google.com"
If CLng(lMeetsID) < 0 Then Response.Redirect "http://www.google.com"

iMyBib = Request.QueryString("my_bib")
If CStr(iMyBib) & "" = "" Then iMyBib = "0"
If Not IsNumeric(iMyBib) Then Response.Redirect "http://www.google.com"
If CLng(iMyBib) < 0 Then Response.Redirect "http://www.google.com"

lRosterID = Request.QueryString("roster_id")
If CStr(lRosterID) & "" = "" Then lRosterID = "0"
If Not IsNumeric(lRosterID) Then Response.Redirect "http://www.google.com"
If CLng(lRosterID) < 0 Then Response.Redirect "http://www.google.com"

bSuccess = False

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
												
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MeetsID, MeetName, MeetName FROM Meets WHERE MeetDate >= '" & Date - 3 & "' ORDER By MeetDate"
rs.Open sql, conn, 1, 2
Meets = rs.GetRows()
rs.Close
Set rs = Nothing

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
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
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'cc_results_text' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn2, 1, 2
	If rs.RecordCount > 0 Then Session("access_cc_results_text") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_cc_results_text") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
		
		iMyBib = CleanInput(Trim(Request.Form.Item("my_bib")))
        If sHackMsg = vbNullString Then lMeetsID = CleanInput(Trim(Request.Form.Item("meet_id")))
		If sHackMsg = vbNullString Then sMobileNumber = CleanInput(Trim(Request.Form.Item("mobile_number")))
		If sHackMsg = vbNullString Then lCellProvider = CleanInput(Trim(Request.Form.Item("cell_provider")))
		
		If sHackMsg = vbNullString Then
            sMobileNumber = Replace(sMobileNumber, "-", "")
            sMobileNumber = Replace(sMobileNumber, ".", "")
            sMobileNumber = Replace(sMobileNumber, "(", "")
            sMobileNumber = Replace(sMobileNumber, ")", "")
            sMobileNumber = Replace(sMobileNumber, " ", "")
            sMobileNumber = Trim(sMobileNumber)

            'get roster_id if missing
            If CLng(lRosterID) = 0 Then
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT RaceID FROM Races WHERE MeetsID = " & lMeetsID
                rs.Open sql, conn, 1, 2
                Do While Not rs.EOF
                    sMeetRaces = sMeetRaces & rs(0).Value & ", "
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing

                If Not sMeetRaces = vbNullString Then sMeetRaces = Left(sMeetRaces, Len(sMeetRaces) - 2)

                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT RosterID FROM IndRslts WHERE Bib = " & iMyBib & " AND RaceID IN (" & sMeetRaces & ")"
                rs.Open sql, conn, 1, 2
                lRosterID = rs(0).Value
                rs.Close
                Set rs = Nothing
            End If

            bFound = False
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT MobileNumber, CellProvider, RosterID FROM MobileSettings WHERE Bib = " & iMyBib & " AND MeetsID = " & lMeetsID
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then
                rs(0).Value = sMobileNumber
                rs(1).Value = lCellProvider
                If rs(2).Value = "0" Then rs(2).Value = lRosterID
                rs.Update
                bFound = True
            End If
            rs.Close
            Set rs = Nothing

            If bFound = False Then
                sql = "INSERT INTO MobileSettings (Bib, MeetsID, MobileNumber, CellProvider, RosterID) VALUES (" & iMyBib & ", " & lMeetsID & ", '" 
                sql = sql & sMobileNumber & "', " & lCellProvider & ", " & lRosterID & ")"
                Set rs = conn.execute(sql)
                Set rs = Nothing
            End If

            sql = "SELECT MeetName FROM Meets WHERE MeetsID = " & lMeetsID
            Set rs = conn.Execute(sql)
            sMeetName = Replace(rs(0).Value, "''", "'")
            Set rs = Nothing

            'send test sms to this person
            Set cdoMessage = Server.CreateObject("CDO.Message")
		    With cdoMessage
                Set .Configuration = cdoConfig
                .From = "bob.schneider@gopherstateMeets.com"
			    .To = sMobileNumber & GetSendURL(lCellProvider)
			    .TextBody = "Welcome to Gopher State Meets Cross-Country/Nordic Ski Text Mesage Results.  Good luck at the races."
			    .Send
		    End With
	        Set cdoMessage = Nothing

			Set cdoMessage = CreateObject("CDO.Message")
			With cdoMessage
				Set .Configuration = cdoConfig
				.To = "bob.schneider@gopherstateMeets.com"
				.From = "bobs@h51softtware.net"
				.Subject = "New Cross-Country/Nordic Ski Text Message Registration"
				.TextBody = "Bib " & iMyBib & " in " & sMeetName & " has just signed up to receive thier results via text messaging."
				.Send
			End With
			Set cdoMessage = Nothing

            bSuccess = True
        End If
   End If
End If

'log this user if they are just entering the site
If Session("access_cc_results_text") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'cc_results_text')"
	Set rs = conn2.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'cc_results_text'"
	Set rs = conn2.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_cc_results_text")
End If

Private Function GetSendURL(lProviderID)
	If Not CStr(lProviderID) & "" = ""  Then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT SendURL FROM CellProviders WHERE CellProvidersID = " & lProviderID
		rs.Open sql, conn, 1, 2
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
<title>GSE&copy; CC/Nordic Ski Results by Text</title>
<meta name="description" content="Text Results Set-Up a Gopher State Meets (GSE) cross-country/Nordic ski meet.">
<!--#include file = "../includes/js.asp" --> 

<script>
    function chkFlds() {
        if (document.sms_info.mobile_number.value == '' ||
        document.sms_info.cell_provider.value == '' ||
        document.sms_info.meet_id.value == '' ||
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
    <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Meets">
    <h1 class="h1">Gopher State Meets Cross-Country/Nordic Ski Text Results Input</h1>

    <div class="bg-info">
        <a href="cc_results_text.asp">Refresh Page</a>
    </div>

    <%If bSuccess = True Then%>
        <div class="bg-success">
            <h4 class="h4">Success!</h4>

            <p>
                You are now set up to receive your results for this race by text message.  You should be receiving a confirmation text very soon.  "
                PLEASE REPLY "Got It" to that message so we know you received it.
            </p>
        </div>
    <%Else%>
        <div class="bg-warning">
            <h4 class="h4">Enter Mobile Data</h4>
            <p>
                In order to send you your results via text message, we need you to identify the meet you are participating in, your bib number, your 
                mobile number & provider.  We will not use this information for anything other than sending your results to up to three numbers.  By 
                submitting this information you are agreeing to allow us to do this.  IMPORTANT NOTE:  THE RESULTS YOU RECEIVE ARE UNOFFICIAL AND
                SUBJECT TO CHANGE!
            </p>
            <p>
                Is your provider not listed?  Encountered problems using this page?  Let us know about it 
                <a href="contact.asp" style="font-weight: bold;color: red;">here.</a>
            </p>

            <%If Not sErrMsg = vbNullString Then%>
                <div class="bg-danger"><%=sErrMsg%></div>
            <%End If%>
                <form class="form-horizontal" role="form" name="sms_info" method="post" 
                    action="cc_results_text.asp?roster_id=<%=lRosterID%>&amp;my_bib=<%=iMyBib%>&amp;meet_id=<%=lMeetsID%>" onsubmit="return chkFlds();">
                <div class="form-group">
 	                <label for="meet_id" class="control-label col-xs-3">Meet:</label>
                    <div class="col-xs-9">
                        <select class="form-control" name="meet_id" id="meet_id">
                            <option value="">&nbsp;</option>
                            <%For i = 0 To UBound(Meets, 2)%>
                                <%If CLng(lMeetsID) = CLng(Meets(0, i)) Then%>
                                    <option value="<%=Meets(0, i)%>" selected><%=Replace(Meets(1, i), "''", "'")%> (<%=Meets(2, i)%>)</option>
                                <%Else%>
                                    <option value="<%=Meets(0, i)%>"><%=Replace(Meets(1, i), "''", "'")%> (<%=Meets(2, i)%>)</option>
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
