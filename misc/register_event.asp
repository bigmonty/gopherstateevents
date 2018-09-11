<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventID, lEventDirID, lEventType, lRaceType
Dim sFirstName, sLastName, sPhone, sEmail, sUserID, sPassword, sEventName, sLocation, sComments, sTimingMethod, sNeedBibs, sNeedPins, sWaiver, sRaceName
Dim sDist, sErrMsg, sMsg, sStartTime, sCertified, sStartType, sWebsite, sMyName
Dim iNumRaces, iMileage, iRaceFee, iPartFee, iNumParts, iOldRaceFee, iPinFee
Dim sngMlgFee, sngTotal
Dim dThisDate, dEventDate
Dim EventTypes()
Dim cdoMessage, cdoConfig
Dim sHackMsg, sMsgText
Dim bHasRaces

If Session("role") = "event_dir" Then lEventDirID = Session("event_dir_id")

If CStr(lEventDirID) = vbNullString Then lEventDirID = Request.QueryString("event_dir_id")
lEventID = Request.QueryString("event_id")
iNumRaces = Request.QueryString("num_races")
iMileage = Request.QueryString("mileage")
iNumParts = Request.QueryString("num_parts")
sTimingMethod = Request.QueryString("timing_method")
sNeedPins = Request.QueryString("need_pins")

bHasRaces = False

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
	
'get event types
i = 0
ReDim EventTypes(1, 0)
sql = "SELECT EvntRaceTypesID, EvntRaceType FROM EvntRaceTypes ORDER BY EvntRaceType"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventTypes(0, i) = rs(0).Value
	EventTypes(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve EventTypes(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("submit_event_dir") = "submit_event_dir" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'register_event' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_evnt_reg") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_evnt_reg") = "y" Then	'if they are an authorized user allow them to proceed
        sFirstName = CleanInput(Trim(Replace(Request.Form.Item("first_name"), "''", "'")))
	    If sHackMsg = vbNullString Then sLastName = Replace(Request.Form.Item("last_name"), "''", "'")
	    If sHackMsg = vbNullString Then sPhone =  Replace(Request.Form.Item("phone"), "''", "'")
	    If sHackMsg = vbNullString Then sEmail =  Replace(Request.Form.Item("email"), "''", "'")

		If sHackMsg = vbNullString Then
        	sUserID =  Left(sFirstName, 2) & "_" & Left(sLastName, 4)
	        sPassword =  Left(sFirstName, 2) & "_" & Left(sLastName, 4)

            sql = "INSERT INTO EventDir (FirstName, LastName, Phone, Email, UserID, Password) VALUES ('" & sFirstName & "', '" & sLastName & "', '" & sPhone 
            sql = sql & "', '" & sEmail & "', '" & sUserID & "', '" & sPassword & "')"
            Set rs = conn.Execute(sql)
            Set rs = Nothing

            'get event director id
		    Set rs = Server.CreateObject("ADODB.Recordset")
		    sql = "SELECT EventDirID FROM EventDir WHERE FirstName = '" & sFirstName & "' AND LastName = '" & sLastName & "' ORDER BY EventDirID DESC"
		    rs.Open sql, conn, 1, 2
		    lEventDirID = rs(0).Value
		    rs.Close
		    Set rs = Nothing

	        Session.Contents.Remove("access_evnt_reg")
        End If
    End If
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'register_event' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_evnt_reg") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_evnt_reg") = "y" Then	'if they are an authorized user allow them to proceed
        sEventName = CleanInput(Trim(Request.Form.Item("event_name")))
        If sHackMsg = vbNullString Then sLocation = CleanInput(Trim(Request.Form.Item("location")))
         sComments = CleanInput(Trim(Request.Form.Item("comments")))
		 
		If sHackMsg = vbNullString Then
	        dThisDate = Request.Form.Item("month") & "/" & Request.Form.Item("day") & "/" & Request.Form.Item("year")

	        If IsDate(dThisDate) And dThisDate <> "1/1/1900" Then
                If CDate(dThisDate) >= Date Then
		            dEventDate = dThisDate
		            lEventType = Request.Form.Item("event_type")
                    sEventName = Replace(sEventName, "'", "''")
                    sLocation = Replace(sLocation, "'", "''")
	                If Not sComments & "" = "" Then sComments = Replace(sComments, "'", "''")
		            iNumRaces = Request.Form.Item("num_races")
                    iMileage = Request.Form.Item("mileage")
                    iNumParts = Request.Form.Item("num_parts")
                    sTimingMethod = Request.Form.Item("timing_method")
                    sWebsite =  Request.Form.Item("website")

                    sNeedBibs = "n"
                    sNeedPins = "n"

                    If Request.Form.Item("need_bibs") = "on" Then sNeedBibs = "y"
                    If Request.Form.Item("need_pins") = "on" Then sNeedPins = "y"

		            sql = "INSERT INTO Events (EventName, EventDate, EventType, EventDirID, Comments, DateReg, WhenShutdown, FeeIncrDate, Location, "
                    sql = sql & "TimingMethod, NeedBibs, NeedPins, Website) VALUES ('" & sEventName & "', '" & dEventDate & "', " & lEventType & ", " 
                    sql = sql & lEventDirID & ", '" & sComments & "', '" & Date & "', '" & CDate(dEventDate) - 1  & "', '" & Date & "', '" & sLocation 
                    sql = sql & "', '"  & sTimingMethod & "', '" & sNeedBibs & "', '" & sNeedPins & "', '" & sWebsite & "')"
		            Set rs = conn.Execute(sql)
		            Set rs = Nothing
		
		            'get event id
		            Set rs = Server.CreateObject("ADODB.Recordset")
		            sql = "SELECT EventID FROM Events WHERE EventName = '" & sEventname & "' AND EventDate = '" & dEventDate & "'"
		            rs.Open sql, conn, 1, 2
		            lEventID = rs(0).Value
		            rs.Close
		            Set rs = Nothing
 
                    sWaiver = "I understand that running a road race is a potentially dangerous activity. I do hereby "
                    sWaiver = sWaiver & "waive and release any and all claims for damages that I may incur as a result of my "
                    sWaiver = sWaiver & "participation in this event against the event and its organizers, all sponsors, "
                    sWaiver = sWaiver & "employees, volunteers, or officials of these organizations. I further certify that have full "
                    sWaiver = sWaiver & "knowledge of the risks involved in this event and that I am physically fit and sufficiently "
                    sWaiver = sWaiver & "trained to participate. If, however, as a result of my participation in the race I require "
                    sWaiver = sWaiver & "medical attention, I hereby give consent to authorize medical personnel to provide "
                    sWaiver = sWaiver & "such medical care as deemed necessary.  " & vbCrLf & vbCrLf
                    sWaiver = sWaiver & "I have read the foregoing and certify my agreement by clicking the button below. "
        
                     'insert into waiver table
                    sql = "INSERT INTO Waiver (EventID, Waiver) VALUES (" & lEventID & ", '" & sWaiver & "')"
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing
         
                     'insert into site info table
                    sql = "INSERT INTO SiteInfo (EventID) VALUES (" & lEventID & ")"
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing
         
                     'insert into site info table
                    sql = "INSERT INTO EventsWeb (EventsID, MetaDescription) VALUES (" & lEventID & ", 'Meta Description for " & sEventname & " on " & dEventDate & "')"
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing
         
                     'insert into event asgmt table
                    sql = "INSERT INTO EventAsgmt (EventID) VALUES (" & lEventID & ")"
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing

	                Session.Contents.Remove("access_evnt_reg")
	
                    'get event dir info
	                sql = "SELECT FirstName, LastName, Email, Phone FROM EventDir WHERE EventDirID = " & lEventDirID
	                Set rs = conn.Execute(sql)
	                sMyName = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
                    sEmail = rs(2).Value
                    sPhone = rs(3).Value
	                Set rs = Nothing

	                sMsg = vbCrLf & "This is notification that a new event has been registered at Gopher State Events:" & vbCrLf & vbCrLf
	                sMsg = sMsg & "Event Name: " & sEventName & vbCrLf
	                sMsg = sMsg & "Event Date: " & dEventDate & vbCrLf
                    sMsg = sMsg & "Event Type: " & GetThisType(lEventType) & vbCrLf
                    sMsg = sMsg & "Location: " & sLocation & vbCrLf
	                sMsg = sMsg & "Timing Method: " & sTimingMethod & vbCrLf
	                sMsg = sMsg & "Need Bibs: " & sNeedBibs & vbCrLf
	                sMsg = sMsg & "Need PIns: " & sNeedPins & vbCrLf
	                sMsg = sMsg & "Number of Races: " & iNumRaces & vbCrLf & vbCrLf

                    sMsg = sMsg & "Submitted By: " & sMyName & vbCrLf
                    sMsg = sMsg & "Email: " & sEmail & vbCrLf
                    sMsg = sMsg & "Phone: " & sPhone & vbCrLf & vbCrLf

                    sMsg = sMsg & "Comments: " & sComments & vbCrLf

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%
	
	                Set cdoMessage = CreateObject("CDO.Message")
	                With cdoMessage
		                Set .Configuration = cdoConfig
		                .To = "bob.schneider@gopherstateevents.com;"
		                .From = "bob.schneider@gopherstateevents.com"
	                    .Subject = "New GSE Event Registration"
		                .TextBody = sMsg
		                .Send
	                End With
	                Set cdoMessage = Nothing
	                Set cdoConfig = Nothing
                Else
                    sErrMsg = "Please select a future date."
                End If
            Else
                 sErrMsg = "Please select a valid date."
            End If
        End If
    End If
ElseIf Request.Form.Item("submit_races") = "submit_races" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'register_event' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_evnt_reg") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_evnt_reg") = "y" Then	'if they are an authorized user allow them to proceed
        iOldRaceFee = 0
        iPinFee = 0

        For i = 1 To iNumRaces
            sRaceName = CleanInput(Trim(Replace(Request.Form.Item("race_name_" & i), "'", "''")))
		 
            If Not sRaceName = vbNullString Then
		        If sHackMsg = vbNullString Then
                    sDist = Request.Form.Item("dist_" & i)
                    lRaceType = Request.Form.Item("race_type_" & i)
                    sStartTime = Request.Form.Item("start_time_" & i) & Request.Form.Item("am_pm_" & i)
                    sCertified = Request.Form.Item("certified_" & i)
                    sStartType = Request.Form.Item("start_type_" & i)

		            sql = "INSERT INTO RaceData (EventID, RaceName, Dist, Type, StartTime, Certified, StartType) VALUES (" & lEventID & ", '" & sRaceName 
                    sql = sql & "', '" & sDist & "', " & lRaceType & ", '" & sStartTime & "', '" & sCertified & "', '" & sStartType  & "')"
		            Set rs = conn.Execute(sql)
		            Set rs = Nothing

                    bHasRaces = True

                    'calculate estimate
                    iRaceFee = GetRaceFee(sDist)

                    If CInt(iRaceFee) > CInt(iOldRaceFee) Then 
                        iOldRaceFee = iRaceFee
                    Else
                        iRaceFee = iOldRaceFee
                    End If
                End If
            End If
        Next

        If sNeedPins = "y" Then iPinFee = 15
        iPartFee = iNumParts
        sngMlgFee = GetMileageFee(iMileage)
        sngTotal = CInt(iRaceFee) + CInt(iPartFee) + CSng(sngMlgFee) + CInt(iPinFee)

	    Session.Contents.Remove("access_evnt_reg")
	
	    sMsg = vbCrLf & "Races have been added to: " & EventName(lEventID) & vbCrLf & vbCrLf

	    sMsg = sMsg & "The details of the estimate are found below: " & vbCrLf
	    sMsg = sMsg & "Race Fee: $" & iRaceFee & vbCrLf
        If sNeedPins = "y" Then sMsg = sMsg & "Pin Fee: $15 " & vbCrLf
        sMsg = sMsg & "Participant Fee: $" & iPartFee & vbCrLf
        sMsg = sMsg & "Mileage Fee: $" & sngMlgFee & vbCrLf
	    sMsg = sMsg & "Estimate Total: $" & sngTotal & vbCrLf

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%
	
	    Set cdoMessage = CreateObject("CDO.Message")
	    With cdoMessage
		    Set .Configuration = cdoConfig
		    .To = "bob.schneider@gopherstateevents.com;"
		    .From = "bob.schneider@gopherstateevents.com"
	        .Subject = "GSE Event Estimate"
		    .TextBody = sMsg
		    .Send
	    End With
	    Set cdoMessage = Nothing
    End If
End If

'log this user if they are just entering the site
If Session("access_evnt_reg") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'register_event')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If

If CStr(lEventDirID) = vbNullString Then lEventDirID = 0
If CStr(lEventType) = vbNullString Then lEventType = 0
If CStr(dEventDate) = vbNullString Then dEventDate = "1/1/1900"
If CStr(iNumRaces) = vbNullString Then iNumRaces = 1

Private Function GetThisType(lEventType)
	sql = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs = conn.Execute(sql)
	GetThisType = rs(0).Value
	Set rs = Nothing
End Function

Private Function EventName(lThisEvent)
	sql = "SELECT EventName FROM Events WHERE EventID = " & lThisEvent
	Set rs = conn.Execute(sql)
	EventName = Replace(rs(0).Value, "''", "'")
	Set rs = Nothing
End Function

Private Function GetRaceFee(sThisDist)
    Select Case sThisDist
        Case "26.2_mi"
            GetRaceFee = 1500
        Case "13.1_mi"
            GetRaceFee = 1000
        Case "15_mi"
            GetRaceFee = 1000
        Case "10_mi"
            GetRaceFee = 1000
        Case "5_mi"
            GetRaceFee = 450
        Case "4_mi"
            GetRaceFee = 400
        Case "8_km"
            GetRaceFee = 450
        Case "50_km"
            GetRaceFee = 1500
        Case "25_km"
            GetRaceFee = 1500
        Case "20_km"
            GetRaceFee = 1000
        Case "15_km"
            GetRaceFee = 1000
        Case "10_km"
            GetRaceFee = 500
        Case Else
            GetRaceFee = 400
    End Select

    If sTimingMethod = "RFID" Then 
        If CSng(GetRaceFee) < 1000 Then GetRaceFee = CSng(GetRaceFee) + 300
    End If
End Function

Private Function GetMileageFee(iThisMileage)
    If CInt(iThisMileage) < 150 Then
        GetMileageFee = Round(CInt(iMileage)*0.75)
    Else
        GetMileageFee = 350
    End If
End Function

%>
<!--#include file = "../includes/clean_input.asp" -->
<%

Set cdoConfig = Nothing    
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Event Registration</title>
<!--#include file = "../includes/js.asp" -->

<script>
function chkEvntDir(){
 	if (document.add_event_dir.first_name.value == '' || 
 	    document.add_event_dir.last_name.value == '' ||
 	    document.add_event_dir.email.value == '' ||
	 	document.add_event_dir.phone.value == '')
		{
  		alert('All fields are required.');
  		return false
  		}
	else
   		return true
}

function chkEvent(){
 	if (document.new_event.event_name.value == '' ||
        document.new_event.location.value == '' || 
        document.new_event.timing_method.value == '' || 
	 	document.new_event.event_type.value == '')
		{
  		alert('All fields are required.');
  		return false
  		}
	else
   		return true
}
</script>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
        <div id="main" style="padding-left: 10px;">
			<h3 class="admin_hdr">Register Your Event With Gopher State Events</h3>

            <p>The purpose of this utility is to gather information about your event for scheduling purposes and/or for creating estimates for event
            management services.  You are in no way bound by any obligation, financial or otherwise, upon submitting this form.  Upon completing this
            process you will be given an estimate for the management of this event.</p>

            <%If CLng(lEventDirID) = 0 Then%>
			    <h4 class="h4">Enter Your (Event Director) Information</h4>
			
                <p style="color: #00e;">IMPORTANT NOTE:  If you are a current GSE event director you should <a href="/default.asp?sign_out=y" 
                    style="font-weight: bold;color: #039;">sign in</a> in to your account and then return to this page.  That will
                attach this event to your existing account rather than creating a completely new account.</p>

			    <form name="add_event_dir" method="Post" action="register_event.asp" onsubmit="return chkEvntDir();">
			    <table style="margin:10px;">
				    <tr>
					    <th>First Name:</th>
					    <td><input type="text" name="first_name" id="first_name"></td>
					    <th>Last Name:</th>
                        <td><input type="text" name="last_name" id="last_name"></td>
				    </tr>
				    <tr>
					    <th>Mobile Phone:</th>
					    <td><input type="text" name="phone" id="phone"></td>
					    <th>Email:</th>
					    <td><input type="text" name="email" id="email"></td>
				    </tr>
				    <tr>
					    <td style="background-color:#ececd8;text-align:center;" colspan="4">
						    <input type="hidden" name="submit_event_dir" id="submit_event_dir" value="submit_event_dir">
						    <input type="submit" name="submit2" id="submit2" value="Submit Event Director">
					    </td>
				    </tr>
			    </table>
			    </form>
            <%Else%>
                <%If CLng(lEventID) = 0 Then%>
			        <h4 class="h4">Enter Event Information</h4>
			
			        <%If Not sErrMsg = vbNullString Then%>
				        <p><%=sErrMsg%></p>
			        <%End If%>
			
			        <form name="new_event" method="Post" action="register_event.asp?event_dir_id=<%=lEventDirID%>" onsubmit="return chkEvent();">
			        <table>
				        <tr>
					        <th>Event Name:</th>
					        <td><input type="text" name="event_name" id="event_name" value="<%=sEventName%>"></td>
					        <th>Event Date:</th>
					        <td>
						        <select name="month" id="month">
							        <%For i = 1 To 12%>
								        <%If Month(dEventDate) = i Then%>
                                            <option value="<%=i%>" selected><%=i%></option>
                                        <%Else%>
                                            <option value="<%=i%>"><%=i%></option>
                                        <%End If%>
							        <%Next%>
						        </select>
						        /
						        <select name="day" id="day">
							        <%For i = 1 To 31%>
								        <option value="<%=i%>"><%=i%></option>
							        <%Next%>
						        </select>
						        /
						        <select name="year" id="year">
							        <%For i = Year(Date) To Year(Date) + 2%>
								        <option value="<%=i%>"><%=i%></option>
							        <%Next%>
						        </select>
					        </td>
                        </tr>
                        <tr>
					        <th>Event Type:</th>
					        <td>
						        <select name="event_type" id="event_type">
							        <option value="">&nbsp;</option>
							        <%For i = 0 To UBound(EventTypes, 2) - 1%>
								        <%If CLng(lEventType) = CLng(EventTypes(0, i)) Then%>
                                            <option value="<%=EventTypes(0, i)%>" selected><%=EventTypes (1, i)%></option>
                                        <%Else%>
                                            <option value="<%=EventTypes(0, i)%>"><%=EventTypes (1, i)%></option>
                                        <%End If%>
							        <%Next%>
						        </select>
					        </td>
					        <th>Location (City, St/Prov):</th>
					        <td> <input type="text" name="location" id="location" value="<%=sLocation%>"></td>
				        </tr>
                        <tr>
						    <th>Timing Method:</th>
						    <td>
						        <select name="timing_method" id="timing_method">
                                    <option value="">&nbsp;</option>
   							        <%If sTimingMethod = "Conv" Then%>
                                        <option value="Conv" selected>Conv</option>
								        <option value="RFID">RFID</option>
                                    <%ElseIf sTimingMethod = "RFID" Then%>
                                        <option value="Conv">Conv</option>
								        <option value="RFID" selected>RFID</option>
                                    <%Else%>
                                        <option value="Conv">Conv</option>
								        <option value="RFID">RFID</option>
                                    <%End If%>
						        </select>
                            </td>
                            <th>
                                <%If sNeedBibs = "y" Then%>
                                    <input type="checkbox" name="need_bibs" id="need_bibs" checked>&nbsp;Need Bibs (Free)
                                <%Else%>
                                    <input type="checkbox" name="need_bibs" id="need_bibs">&nbsp;Need Bibs (Free)
                                <%End If%>
                            </th>
                            <th>
                                <%If sNeedPins = "y" Then%>
                                    <input type="checkbox" name="need_pins" id="need_pins" checked>&nbsp;Need Pins ($15)
                                <%Else%>
                                    <input type="checkbox" name="need_pins" id="need_pins">&nbsp;Need Pins ($15)
                                <%End If%>
                            </th>
                        </tr>
				        <tr>
					        <th valign="top">Number of Races:</th>
					        <td>
                                <select name="num_races" id="num_races">
                                    <%For i = 1 To 9%>
                                        <option value="<%=i%>"><%=i%></option>
                                    <%Next%>
                                </select>
                            </td>
					        <th valign="top">Distance From Minnetonka, MN:</th>
					        <td>
                                <select name="mileage" id="mileage">
                                    <%For i = 50 To 300 Step 10%>
                                        <option value="<%=i%>"><%=i%></option>
                                    <%Next%>
                                </select>
                            </td>
				        </tr>
				        <tr>
					        <th>Website (optional):</th>
					        <td> <input type="text" name="website" id="website" value="<%=sWebsite%>"></td>
					        <th style="text-align: left;padding-left: 20px;" colspan="2">
                                Anticipated Field Size (all races):
                                <select name="num_parts" id="num_parts">
                                    <%For i = 50 To 2000 Step 50%>
                                        <option value="<%=i%>"><%=i%></option>
                                    <%Next%>
                                </select>
                            </th>
				        </tr>
				        <tr>
					        <th valign="top">Comments:</th>
					        <td colspan="3"><textarea name="comments" id="comments" cols="75" rows="3"><%=sComments%></textarea></td>
				        </tr>
				        <tr>
					        <td style="background-color:#ececd8;text-align:center;" colspan="4">
						        <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
						        <input type="submit" name="submit1" id="submit1" value="Submit Event">
					        </td>
				        </tr>
			        </table>
			        </form>
                <%Else%>
                    <%If bHasRaces = False Then%>
			            <h4 class="h4">Enter Race Information</h4>
                    
			            <form name="add_races" method="Post" 
                            action="register_event.asp?need_pins=<%=sNeedPins%>&amp;timing_method=<%=sTimingMethod%>&amp;event_dir_id=<%=lEventDirID%>&amp;event_id=<%=lEventID%>&amp;num_parts=<%=iNumParts%>&amp;num_races=<%=iNumRaces%>&amp;mileage=<%=iMileage%>">
			            <table style="margin:10px;">
                            <%For i = 1 To iNumRaces%>
                                <tr>
                                    <td>
                                        <h5>Race <%=i%></h5>
                                        <table>
				                            <tr>
					                            <th>Race Name:</th>
					                            <td><input type="text" name="race_name_<%=i%>" id="race_name_<%=i%>"></td>
					                            <th>Distance:</th>
                                                <td>
                                                    <select name="dist_<%=i%>" id="dist_<%=i%>">
                                                        <option value="5_km">5K</option>
                                                        <option value="8_km">8K</option>
                                                        <option value="10_km">10K</option>
                                                        <option value="15_km">15K</option>
                                                        <option value="20_km">20K</option>
                                                        <option value="50_km">50K</option>
                                                        <option value="1_mi">1 Mile</option>
                                                        <option value="2_mi">2 Mile</option>
                                                        <option value="3_mi">3 Mile</option>
                                                        <option value="4_mi">4 Mile</option>
                                                        <option value="5_mi">5 Mile</option>
                                                        <option value="10_mi">10 Mile</option>
                                                        <option value="15_mi">15 Mile</option>
                                                        <option value="13.1_mi">Hale-Marathon</option>
                                                        <option value="26.2_mi">Marathon</option>
                                                        <option value="Other">Other</option>
                                                    </select>
                                                </td>
					                            <th>Race Type:</th>
					                            <td>
						                            <select name="race_type_<%=i%>" id="race_type_<%=i%>">
							                            <%For j = 0 To UBound(EventTypes, 2) - 1%>
                                                            <%If j = 6 Then%>
                                                                <option value="<%=EventTypes(0, j)%>" selected><%=EventTypes (1, j)%></option>
                                                            <%Else%>
                                                                <option value="<%=EventTypes(0, j)%>"><%=EventTypes (1, j)%></option>
                                                            <%End If%>
							                            <%Next%>
						                            </select>
					                            </td>
				                            </tr>
				                            <tr>
					                            <th>Start Time:</th>
					                            <td>
                                                    <input type="text" name="start_time_<%=i%>" id="start_time_<%=i%>">
                                                    <select name="am_pm_<%=i%>" id="am_pm_<%=i%>">
                                                         <option value="am">AM</option>
                                                         <option value="pm">PM</option>
                                                    </select>
                                                </td>
					                            <th>Start Type:</th>
					                            <td>
                                                    <select name="start_type_<%=i%>" id="start_type_<%=i%>">
                                                        <option value="mass">Mass Start</option>
                                                        <option value="interval">Interval Start</option>
                                                        <option value="wave">Wave Start</option>
                                                    </select>
                                                </td>
					                            <th>Certified Course:</th>
					                            <td>
                                                    <select name="certified_<%=i%>" id="certified_<%=i%>">
                                                        <option value="n">No</option>
                                                        <option value="y">Yes</option>
                                                    </select>
                                                </td>
				                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            <%Next%>
				            <tr>
					            <td style="background-color:#ececd8;text-align:center;" colspan="6">
						            <input type="hidden" name="submit_races" id="submit_races" value="submit_races">
						            <input type="submit" name="submit3" id="submit3" value="Submit Race(s)">
					            </td>
				            </tr>
			            </table>
			            </form>
                    <%Else%>
                        <h4 class="h4">Submission Complete!</h4>

                        <p>Thank you for submitting this event for consideration.  We have received your submission and will be in contact with you.</p>

                        <p>NOTE:  If your event has an informational flyer/brochure please send it to <a href="mailto:bob.schneider@gopherstateevents.com"
                        style="font-weight:bold;">bob.schneider@gopherstateevents.com</a>.</p>
            
                        <p>We are capable of managing 3 races simultaneously.  You can determine if we have room on our schedule by checking our 
                        <a href="http://www.gopherstateevents.com/calendar/calendar.asp" style="font-weight: bold;">calendar</a>.  You can find an 
                        estimate for managing your event below.  Please understand that this is just an estimate.  On some occassions we will make 
                        modifications in our standard pricing to accommodate special circumstances.</p>

                        <h5>Event Management Estimate:</h5>

                        <ul  style="margin-left: 25px;font-size: 0.8em;">
                            <%If CInt(iPinFee) > 0 Then%>
                                <li><span style="font-weight: bold;">Pin Fee:</span>&nbsp;$15</li>
                            <%End If%>
                            <li><span style="font-weight: bold;">Race Fee:</span>&nbsp;$<%=iRaceFee%></li>
                            <li><span style="font-weight: bold;">Participant Fee:</span>&nbsp;$<%=iPartFee%></li>
                            <%If CSng(sngMlgFee) = 350 Then%>
                                <li><span style="font-weight: bold;">Mileage/Lodging:</span>&nbsp;$<%=sngMlgFee%> ($200 Mileage + $150 Lodging/Meals)</li>
                            <%Else%>
                                <li><span style="font-weight: bold;">Mileage:</span>&nbsp;$<%=sngMlgFee%> (@ 0.75/mile)</li>
                            <%End If%>
                            <li><span style="font-weight: bold;">-----------------------------------</span></li>
                            <li><span style="font-weight: bold;">Estimate Total:&nbsp;$<%=sngTotal%></span></li>
                        </ul>
                    <%End If%>
                <%End If%>
            <%End If%>
		</div>
	</div>
</div>
<%	
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
