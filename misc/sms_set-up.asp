<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i
Dim lPartID, lCellProvider, lMyID, lEventID
Dim iMyBib
Dim sMobileNumber, sErrMsg, sMyName
Dim CellProviders
Dim cdoMessage, cdoConfig
Dim bFound, bSuccess

lPartID = Request.QueryString("part_id")
If CStr(lPartID) & "" = "" Then Response.Redirect "http://www.google.com"
If Not IsNumeric(lPartID) Then Response.Redirect "http://www.google.com"
If CLng(lPartID) < 0 Then Response.Redirect "http://www.google.com"

iMyBib = Request.QueryString("my_bib")
If CStr(iMyBib) = vbNullString Then iMyBib = 0
If Not IsNumeric(iMyBib) Then Response.Redirect "http://www.google.com"
If CInt(iMyBib) < 0 Then Response.Redirect "http://www.google.com"

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect "http://www.google.com"
If CLng(lEventID) < 0 Then Response.Redirect "http://www.google.com"

bSuccess = False

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

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

If Request.Form.Item("submit_sms") = "submit_sms" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'sms_set-up' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_sms_set-up") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_sms_set-up") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
		
		lMyID = CleanInput(Trim(Request.Form.Item("my_id")))
		If sHackMsg = vbNullString Then sMobileNumber = CleanInput(Trim(Request.Form.Item("mobile_number")))
		If sHackMsg = vbNullString Then lCellProvider = CleanInput(Trim(Request.Form.Item("cell_provider")))
		
		If sHackMsg = vbNullString Then
            If Not CLng(lMyID) = CLng(lPartID) Then
                sErrMsg = "I am sorry.  This is not the code you were given.  You might be able to use your back button to find it and try again."
            Else
                sMobileNumber = Replace(sMobileNumber, "-", "")
                sMobileNumber = Replace(sMobileNumber, ".", "")
                sMobileNumber = Replace(sMobileNumber, "(", "")
                sMobileNumber = Replace(sMobileNumber, ")", "")

                'get bib if it is 0
                If CInt(iMyBib) = 0 Then
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    sql = "SELECT Bib FROM PartRace WHERE EventID = " & lEventID & " AND ParticipantID = " & lPartID
                    rs.Open sql, conn, 1, 2
                    iMyBib = rs(0).Value
                    rs.Close
                    Set rs = Nothing
                End If

                bFound = False
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT MobileNumber, CellProvider FROM MobileSettings WHERE EventID = " & lEventID & " AND Bib = " & iMyBib
                rs.Open sql, conn, 1, 2
                If rs.RecordCount > 0 Then
                    rs(0).Value = sMobileNumber
                    rs(1).Value = lCellProvider
                    rs.Update
                    bFound = True
                End If
                rs.Close
                Set rs = Nothing

                If bFound = False Then
                    sql = "INSERT INTO MobileSettings (MobileNumber, CellProvider, Bib, EventID) VALUES ('" & sMobileNumber & "', " & lCellProvider
                    sql = sql & ", " & iMyBib & ", " & lEventID & ")"
                    Set rs = conn.execute(sql)
                    Set rs = Nothing
                End If

                'send test sms to this person
                Set cdoMessage = Server.CreateObject("CDO.Message")
                Set cdoMessage.Configuration = cdoConfig
		        With cdoMessage
                    .From = "bob.schneider@gopherstateevents.com"
			        .To = sMobileNumber & GetSendURL(lCellProvider)
			        .TextBody = "Welcome to Gopher State Events SMS Race Results.  Good luck at the races."
			        .Send
		        End With
	            Set cdoMessage = Nothing

                'send me an email when they register
                sql = "SELECT FirstName, LastName FROM Participant WHERE ParticipantID = " & lPartID
                Set rs = conn.Execute(sql)
                sMyName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
                Set rs = Nothing

			    Set cdoMessage = CreateObject("CDO.Message")
			    With cdoMessage
				    Set .Configuration = cdoConfig
				    .To = "bob.schneider@gopherstateevents.com"
				    .From = "bobs@h51softtware.net"
				    .Subject = "New SMS Results Registration"
				    .TextBody = sMyName & " has just signed up to receive thier results via text messaging."
				    .Send
			    End With
			    Set cdoMessage = Nothing

                bSuccess = True
            End If
        End If
    End If
End If

'log this user if they are just entering the site
If Session("access_sms_set-up") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'sms_set-up')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'sms_set-up'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_sms_set-up")
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

'see if they have a mobile number on file
lCellProvider = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MobileNumber, CellProvider FROM MobileSettings WHERE ParticipantID = " & lPartID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sMobileNumber = rs(0).Value
    lCellProvider = rs(1).Value
End If
rs.Close
Set rs = Nothing

%>
<!--#include file = "../includes/clean_input.asp" -->
<%

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE&copy; SMS Results Set-Up</title>
<meta name="description" content="SMS Results Set-Up a Gopher State Events (GSE) timed event.">
<!--#include file = "../includes/js.asp" --> 

<script>
    function chkFlds() {
        if (document.sms_info.mobile_number.value == '' ||
        document.sms_info.cell_provider.value == '' ||
        document.sms_info.my_id.value == '') {
            alert('All fields are required!');
            return false;
        }
        else
		    if (isNaN(document.sms_info.mobile_number.value) ||
                isNaN(document.sms_info.my_id.value))
    		    {
			    alert('The mobile number and id can not contain non-numeric values');
			    return false
			    }
        else
            return true;
    }
</script>
</head>

<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Gopher State Events">
    <h4 class="h4">SMS Results Settings</h4>

    <%If bSuccess = True Then%>
        <h4 class="h4">Success!</h4>

        <p>You should be receiving a confirmation text very soon.  Please reply to that message so we know you received it.</p>
    <%Else%>
        <p>In order to send you your results via text message, we need your mobile number and your provider.  We also need you to enter the number that your
        attention was drawn to on the previous screen.  Once again, we will not use this information for anything other than sending YOU your results.  By 
        submitting this information you are agreeing to allow us to do this.</p>

        <p class="bg-danger text-danger">NOTE:  This form must be submitted before the race starts!</p>

        <p>
            Is your provider not listed?  Encountered problems using this page?  Let us know about it 
            <a href="contact.asp" style="font-weight: bold;color: red;">here.</a>
        </p>

        <%If Not sErrMsg = vbNullString Then%>
            <div class="bg-danger"><%=sErrMsg%></div>
        <%End If%>

        <div>
            <form role="form" class="form-inline" name="sms_info" method="post" 
                action="sms_set-up.asp?part_id=<%=lPartID%>&amp;my_bib=<%=iMyBib%>&amp;event_id=<%=lEventID%>" onsubmit="return chkFlds();">
            <div class="form-group">
	           <label for="mobile_number"> Mobile Phone:</label>
                <input class="form-control" type="text" name="mobile_number" id="mobile_number" value="<%=sMobileNumber%>">
            </div>
            <div class="form-group">
                <label for="cell_provider">Provider:</label>
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
            <div class="form-group">
 	            <label for="my_id">Code From Previous Page:</label>
                <input class="form-control" type="text" name="my_id" id="my_id">
            </div>
            <div class="form-group">
		        <input class="form-control" type="hidden" name="submit_sms" id="submit_sms" value="submit_sms">
		        <input class="form-control" type="submit" name="submit1" id="submit1" value="Submit This">
            </div>
            </form>
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
