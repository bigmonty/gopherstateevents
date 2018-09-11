<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, conn2, conn3, rs2, sql2
Dim i
Dim lPerfTrkrID, lEtraxcID, lPartID, lTeamID
Dim sErrMsg, sMsg, sGender, sUserName, sPassword, sConfirmPassword, sFirstName, sLastName, sTeamName, sGradeYear, sEmail
Dim iMonth, iDay, iYear, iWhichStep, iGrade
Dim cdoMessage, cdoConfig
Dim Teams
Dim dDOB, dExpiration

If Session("role") = "perf_trkr" Then Response.Redirect("profile.asp")

iWhichStep = Request.QueryString("which_step")
IF CStr(iWhichStep) = vbNullString Then iWhichstep = 1

lPartID = Request.QueryString("part_id")
lTeamID = Request.QueryString("team_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
	
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"
	
Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get year for roster grades
If Month(Date) <= 7 Then
    sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
Else
    sGradeYear = Right(CStr(Year(Date)), 2)
End If

sql = "SELECT TeamsID, TeamName, Gender, Sport FROM Teams ORDER BY TeamName"
Set rs= conn.Execute(sql)
Teams = rs.GetRows()
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'create_accnt' ORDER BY AuthAccessID DESC"
rs.Open sql, conn3, 1, 2
If rs.RecordCount > 0 Then Session("access_create_accnt") = "y"
rs.Close
Set rs = Nothing

If Session("access_create_accnt") = "y" Then	'if they are an authorized user allow them to proceed
	Dim sHackMsg

    If Request.Form.Item("submit_step1") = "submit_step1" Then
		lPartID = CleanInput(CLng(Request.Form.Item("part_id")))
	    If sHackMsg = vbNullString Then lTeamID = CleanInput(Trim(Request.Form.Item("teams")))

        If sHackMsg = vbNullString Then
            If Not IsNumeric(lPartID) Then sErrMsg = "Your participant id must be numeric.  Please check the value you have been given and try again."

            'if everything checks out
            If sErrMsg = vbNullString Then
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT FirstName FROM Roster WHERE RosterID = " & lPartID & " AND TeamsID = " & lTeamID
                rs.Open sql, conn, 1, 2
                If rs.RecordCount > 0 Then
                    Call GetMyData(lPartID, lTeamID)

                    iWhichStep = 2
                Else
                    sErrMsg = "I'm sorry.  That participant id does not exist on that team.  Please check your data and re-enter.  If you would like "
                    sErrMsg = sErrMsg & "your participant id sent to you please email <a href='mailto:bob.schneider@gopherstateevents.com'>GSE</a>."
                End If
                rs.Close
                Set rs = Nothing
            End If
        End If
    ElseIf Request.Form.Item("submit_step2") = "submit_step2" Then
        If Request.Form.Item("confirmation") = "y" Then 
            iWhichStep = 3
        Else
            sErrMsg = "Not you?  We're sorry.  Please send an email to <a href='mailto:bob.schneider@gopherstateevents.com'>GSE</a> and explain what is "
            sErrMsg = sErrMsg & "incorrect so we can get your account set up.  Depending on the volume of requests this could take a day or two.  Please "
            sErrMsg = sErrMsg & " include your first name, last name, current grade, gender, team, and sport. "
        End If
    ElseIf Request.Form.Item("submit_step3") = "submit_step3" Then
	    sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
	    If sHackMsg = vbNullString Then sEmail = CleanInput(Trim(Request.Form.Item("email")))
	    If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))
	    If sHackMsg = vbNullString Then sConfirmPassword = CleanInput(Trim(Request.Form.Item("confirm_password")))

        If sHackMsg = vbNullString Then
            iMonth = Request.Form.Item("dob_month")
            iDay = Request.Form.Item("dob_day")
            iYear = Request.Form.Item("dob_year")

	        dDOB = iMonth & "/" & iDay & "/" & iYear
	
            If IsDate(dDOB) = False Then sErrMsg = "Your date of birth is not a valid date.  Please adjust."

            If Not CStr(sPassword) = CStr(sConfirmPassword) Then sErrMsg = "Your passwords do not match.  Please adjust."

            If sErrMsg = vbNullString Then
                'check for user name validity
                If ValidUserName(sUserName) = False Then 
                    sErrMsg = "Your user name is not valid.  It is either already in use or not between 5 and 12 characters in length.  "
                    sErrMsg = sErrMsg & "Please adjust and re-enter."
                End If
            End If

            If sErrMsg = vbNullString Then
                'check for password validity
                If ValidPassword(sUserName) = False Then 
                    sErrMsg = "Your password is not valid.  It is either already in use or not between 5 and 12 characters in length.  "
                    sErrMsg = sErrMsg & "Please adjust and re-enter."
                End If
            End If

            If sErrMsg = vbNullString Then
                'check for email uniqueness
                If UniqueEmail(sEmail) = False Then 
                    sErrMsg = "Your email address is already in our system.  If you believe you are the only human using this address, "
                    sErrMsg = sErrMsg & "please log in to your existing account.  If a friend or family member may be using this email "
                    sErrMsg = sErrMsg & "address please use another one. If all else fails, email us at <a href='mailto:bob.schneider@gopherstateevents.com'>GSE</a>."
                End If
            End If

            If sErrMsg = vbNullString Then
                'check for email validity
                If ValidEmail(sEmail) = False Then sErrMsg = "Your email address does not appear to be in a valid format.  Please re-enter."
            End If

            'if everything checks out
            If sErrMsg = vbNullString Then
                Call GetMyData(lPartID, lTeamID)

  	            'insert into perftrkr table
	            sql = "INSERT INTO PerfTrkr (RosterID, UserName, Password, Email, DOB, WhenSubscr, Expiration) VALUES (" & lPartID & ", '" 
                sql = sql & sUserName & "', '" & sPassword & "', '" & sEmail & "', '" & dDOB & "', '" & Now() & "', '" & Date + 7 & "')"
	            Set rs=conn.Execute(sql)
	            Set rs=Nothing

	            'get participant id
	            sql = "SELECT PerfTrkrID FROM PerfTrkr WHERE UserName = '" & sUserName & "' AND Password = '" & sPassword & "'"
	            Set rs = conn.Execute(sql)
	            lPerfTrkrID = rs(0).Value
	            Set rs=Nothing

                'set up my-etraxc account
                Dim lIndSubscrID
                Dim iPIN, iCurrSeas
                Dim PinArray()
                Dim bFound
                Dim dSeasStart
	
			    'assign PIN and check for uniqueness
			    Randomize
			    iPIN = Int(9000*Rnd + 1000)
	
			    'get all PINs and User Names to check for duplicates
			    sql = "SELECT UserName, PIN FROM PartData"
			    Set rs = conn2.Execute(sql)
			    i = 0
			    ReDim PINArray(1, 0)
			    Do While Not rs.EOF
				    PINArray(0, i) = rs(0).Value
				    PINArray(1, i) = rs(1).Value
				    i = i + 1
				    ReDim Preserve PINArray(1, i)
				    rs.MoveNext
			    Loop
			    Set rs = Nothing

			    Do
				    bFound = False
				    For i = 0 to UBound(PINArray, 2) - 1
					    If UCase(CStr(sUserName)) = UCase(CStr(PINArray(0, i))) Then
						    If CInt(iPIN) = CInt(PINArray(1, i)) Then
							    iPIN = Int(9000*Rnd + 1000)
							    bFound = True
							    Exit For
						    End If
					    End If
				    Next
			    Loop While bFound = True

			    sql = "INSERT INTO PartData (TeamID, FirstName, LastName, Grade, Gender, Email, PIN, MinMilesConv, UserName, "
			    sql = sql & "BirthDate, Archive, PerfTrkrID) VALUES (225, '" & sFirstName & "', '" & sLastName & "', 0, '" & sGender & "', '" 
			    sql = sql & sEmail & "', " & iPIN & ", '480', '" & sUserName & "', '" & dDOB & "', 'N', " & lPerfTrkrID & ")"
			    Set rs=conn2.Execute(sql)
			    Set rs=Nothing
	
			    'get part id
			    sql = "SELECT PartID FROM PartData WHERE TeamID = 225 AND FirstName = '" & sFirstName & "' AND LastName = '" & sLastName
			    sql = sql & "' AND Grade = 0 AND Gender = '" & sGender & "' AND TeamID = 225 ORDER BY PartID DESC"
			    Set rs = conn2.Execute(sql)
			    lEtraxcID = rs(0).Value
			    Set rs = Nothing
			
			    'add to the ind expiration table
			    sql = "INSERT INTO IndExpiration (PartID, ExpirationDate) VALUES (" & lEtraxcID & ", '1/1/2020')"
			    Set rs = conn2.Execute(sql)
			    Set rs = Nothing
			
			    'add to part profile
			    sql = "INSERT INTO PartProfile (PartID) VALUES (" & lEtraxcID & ")"
			    Set rs=conn2.Execute(sql)
			    Set rs=Nothing
	
			    'add to infopart table
			    sql = "INSERT INTO InfoPart (PartID) VALUES (" & lEtraxcID & ")"
			    Set rs=conn2.Execute(sql)
			    Set rs=Nothing
	
			    'add to ind subscr table
			    sql = "INSERT INTO IndSubscr (PartID, RegDate) VALUES (" & lEtraxcID & ", '" & Date & "')"
			    Set rs=conn2.Execute(sql)
			    Set rs=Nothing
		
			    'get ind subsc id
			    sql = "SELECT IndSubscrID FROM IndSubscr WHERE PartID = " & lEtraxcID 
			    Set rs = conn2.Execute(sql)
			    lIndSubscrID = rs(0).Value
			    Set rs = Nothing
		
			    'add to ind seasons table
			    dSeasStart = Date - 91
			    sql = "SELECT SeasonsID FROM Seasons WHERE BegDate >= '" & CDate(dSeasStart) & "' ORDER BY BegDate"
			    Set rs = conn2.Execute(sql)
			    iCurrSeas = rs(0).Value
			    Set rs = Nothing

			    For i = CInt(iCurrSeas) To CInt(iCurrSeas) + 3
				    sql = "INSERT INTO IndSeas (IndSubscrID, SeasonsID) VALUES (" & lIndSubscrID & ", " & i & ")"
				    Set rs=conn2.Execute(sql)
				    Set rs=Nothing
			    Next

	            sMsg = "Welcome to GSE's Performance Tracker!" & vbCrLf & vbCrLf
	            sMsg = sMsg & "The purpose of this program is to assist you in tracking the race results of you, your teammates, and your opponents. "
	            sMsg = sMsg & "We hope you will find this utility to be a valuable asset to you and would ask that you provide "
                sMsg = sMsg & "feedback about additional functionality that you would like to see. " & vbCrLf & vbCrLf
	
                sMsg = sMsg & "By creating your Performance Tracker account through GSE you have also received an account at http://www.my-etraxc.com/ "
                sMsg = sMsg & "My-eTRaXC is a FREE online fitness manager and training log for individuals.  It renders well on mobile devices "
                sMsg = sMsg & "and has more in terms of features and functionality than we could list here.  You are encouraged to log in to "
                sMsg = sMsg & "My-eTRaXC using your user name and the PIN listed below and look around."  & vbCrLf & vbCrLf
	
	            sMsg = sMsg & "Your Name: "  & sFirstName & " " & sLastName & vbCrLf
	            sMsg = sMsg & "Gender: "  & sGender  & vbCrLf
	            sMsg = sMsg & "Date of Birth: "  & dDOB  & vbCrLf
            	sMsg = sMsg & "My-eTRaXC PIN: " & iPIN  & vbCrLf  & vbCrLf

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	            Set cdoMessage = CreateObject("CDO.Message")
	            With cdoMessage
		            Set .Configuration = cdoConfig
		            .To = sEmail
		            .From = "bob.schneider@gopherstateevents.com"
	                .CC = "bob.schneider@gopherstateevents.com;"
	                .Subject = "Your GSE Performance Tracker Account"
		            .TextBody = sMsg
		            .Send
	            End With
	            Set cdoMessage = Nothing
                Set cdoConfig = Nothing

	            Session.Contents.Remove("access_create_accnt")

                Response.Redirect "perf_trkr.asp"
            End If
        End If
    End If
End If
		
Private Sub GetMyData(lThisPart, lThisTeam)
    sFirstName = vbNullString
    sLastName = vbNullString
    sGender = vbNullString
    sTeamName = vbNullString
    iGrade = 0

	sql2 = "SELECT r.Gender, t.TeamName, r.FirstName, r.LastName FROM Roster r INNER JOIN Teams t ON r.TeamsID = t.TeamsID WHERE r.RosterID = " & lThisPart
    sql2 = sql2 & " AND r.TeamsID = " & lThisTeam
	Set rs2 = conn.Execute(sql2)
	sGender = rs2(0).Value
    sTeamName  = Replace(rs2(1).Value, "''", "'")
    sFirstName = Replace(rs2(2).Value, "''", "'")
    sLastName = Replace(rs2(3).Value, "''", "'")
	Set rs2 = Nothing

    'get grade
    sql2 = "SELECT Grade" & sGradeYear & " FROM Grades WHERE RosterID = " & lThisPart
    Set rs2 = conn.Execute(sql2)
    iGrade = rs2(0).Value
    Set rs2 = Nothing
End Sub

'log this user if they are just entering the site
If Session("access_create_accnt") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'create_accnt')"
	Set rs = conn3.Execute(sql)
	Set rs = Nothing
End If

%>
<!--#include file = "../../includes/clean_input.asp" -->
<!--#include file = "../../includes/valid_email.asp" -->
<%

Function ValidUserName(sThisUserName) 
	ValidUserName = True

	If Len(sThisUserName) < 5 Or Len(sThisUserName) > 12 Then ValidUserName = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UserName FROM PerfTrkr WHERE UserName = '" & sThisUserName & "'"
    rs.open sql, conn, 1, 2
    If rs.RecordCount > 0 Then ValidUserName = False
    rs.Close
    Set rs = Nothing
End Function

Function ValidPassword(sThisPassword) 
	ValidPassword = True

	If Len(sThisPassword) < 5 Or Len(sThisPassword) > 12 Then ValidPassword = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Password FROM PerfTrkr WHERE Password = '" & sThisPassword & "'"
    rs.open sql, conn, 1, 2
    If rs.RecordCount > 0 Then ValidPassword = False
    rs.Close
    Set rs = Nothing
End Function

Function UniqueEmail(sThisEmail) 
	UniqueEmail = True

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Email FROM PerfTrkr WHERE Email = '" & sThisEmail & "'"
    rs.open sql, conn, 1, 2
    If rs.RecordCount > 0 Then UniqueEmail = False
    rs.Close
    Set rs = Nothing
End Function

Function createRandomString(pwLength)
 	Dim charArray, arrayLength, pw, x
	
 	charArray = "ABCDEFGHIJKLMNPQRSTUVWXYZ0123456789"
 	arrayLength = Len(charArray)
 	pw = ""
	
 	Randomize
	
	For x = 1 To pwLength
  		pw = pw & Mid(charArray, 1 + Int(Rnd * arrayLength), 1)
 	Next
	
 	createRandomString = pw
End Function

Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Create A My GSE&copy; Performance Tracker Account</title>
<meta name="description" content="Create Gopher State Events (GSE) Performance Tracker account.">

<script>
function chkFlds_step1() {
if (document.step_1.part_id.value == '' || 
    document.step_1.teams.value == '')

{
 	alert('ParticipantID and School are required fields!');
 	return false
 	}
else
 	return true;
}

function chkFlds_step2() {
if (document.step_2.confirmation.value == '')

{
 	alert('Please confirm your data before submitting!');
 	return false
 	}
else
 	return true;
}

function chkFlds_step3() {
if (document.step_3.dob_month.value == '' || 
    document.step_3.dob_day.value == '' || 
    document.step_3.dob_year.value == '' || 
    document.step_3.email.value == '' || 
    document.step_3.user_name.value == '' || 
    document.step_3.confirm_password.value == '' || 
    document.step_3.password.value == '' )

{
 	alert('All fields are required!');
 	return false
 	}
else
 	return true;
}</script>
</head>

<body onload="document.new_accnt.first_name.focus();">
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->

    <div class="row">
        <div class="col-md-4">
	        <%If Not sHackMsg = vbNullString Then%>
		        <p class="bg-warning"><%=sHackMsg%></p>
	        <%Else%>
		        <%If Not sErrMsg = vbNullString Then%>
			        <p class="bg-danger"><%=sErrMsg%></p>
		        <%End If%>

			    <div class="bg-warning" style="text-align: center;">
				    Already have an account? <a style="color:#fff;" href="login.asp">Sign In Here</a>
                </div>
                        
                <h4 class="h4">Create Account</h4>

                <%Select Case iWhichStep%>
                    <%Case 1%>
                        <h5 class="h5">Step 1: You and Your Team</h5>

  		                <form class="form" name="step_1" method="Post" action="create_accnt.asp?which_step=1" onSubmit="return chkFlds_step1();">
			            <div class="form-group row">
				            <label class="col-sm-5 form-control-label text-nowrap" for="part_id">Participant ID:</label>
				            <div class="col-sm-7">
                                <input type="text" class="form-control input-sm" name="part_id" id="part_id" value="<%=lPartID%>">
                            </div>
                        </div>
                        <p style="font-size:0.8em;" class="text-warning">
                            (Your coach has been given participant ids for all of their team members.  You may also get it by
                            <a href="mailto:bob.schneider@gopherstateevents.com">contacting us</a>.)
                        </p>
                        <div class="form-group row">
				            <label class="col-sm-5 form-control-label text-nowrap" for="teams">Team-Gender (Sport):</label>
				            <div class="col-sm-7">
                                <select class="form-control" name="teams" id="teams">
                                    <option value=""></option>
                                    <%For i = 0 To UBound(Teams, 2) - 1%>
                                        <option value="<%=Teams(0, i)%>"><%=Teams(1, i)%> - <%=Teams(2, i)%> (<%=Teams(3, i)%>)</option>
                                    <%Next%>
                                </select>
                            </div>
                        </div>
				        <input class="form-control input-sm" type="hidden" name="submit_step1" id="submit_step1" value="submit_step1">
				        <input class="form-control input-sm" type="submit" name="submit1" id="submit1" value="Step 1: Find My School/Team">
		                </form>
                    <%Case 2%>
                        <h5 class="h5">Step 2: Is This You?</h5>

                        <ul class="list-group">
                            <li class="list-group-item">First Name: <%=sFirstName%></li>
                            <li class="list-group-item">Last Name: <%=sLastName%></li>
                            <li class="list-group-item">School: <%=sTeamName%></li>
                            <li class="list-group-item">Gender: <%=sGender%></li>
                            <li class="list-group-item">Grade: <%=iGrade%></li>
                        </ul>

  		                <form class="form-inline" name="step_2" method="Post" 
                            action="create_accnt.asp?which_step=2&amp;part_id=<%=lPartID%>&amp;team_id=<%=lTeamID%>" 
                            onsubmit="return chkFlds_step2();">
                        <select class="form-control" name="confirmation" id="confirmation">
                            <option value=""></option>
                            <option value="y">Yes</option>
                            <option value="n">No</option>
                        </select>
				        <input class="form-control input-sm" type="hidden" name="submit_step2" id="submit_step2" value="submit_step2">
				        <input class="form-control input-sm" type="submit" name="submit2" id="submit2" value="Step 2: Verify My Identity">
		                </form>
                    <%Case 3%>
                        <h5 class="h5">Step 3: Complete Your Registration</h5>

  		                <form class="form" name="step_3" method="Post" 
                            action="create_accnt.asp?which_step=3&amp;part_id=<%=lPartID%>&amp;team_id=<%=lTeamID%>" 
                            onSubmit="return chkFlds_step3();">
			            <div class="form-group row">
				            <label class="col-sm-3 form-control-label" for="dob_month">DOB:</label>
				            <div class="col-sm-3">
					            <select class="form-control" name="dob_month" id="dob_month">
                                    <option value="">&nbsp;</option>
                                    <%For i = 1 To 12%>
								        <option value="<%=i%>"><%=i%></option>
                                    <%Next%>
					            </select>
                            </div>
                            <div class="col-sm-3">
					            <select class="form-control" name="dob_day" id="dob_day">
                                    <option value="">&nbsp;</option>
                                    <%For i = 1 To 31%>
								        <option value="<%=i%>"><%=i%></option>
                                    <%Next%>
					            </select>
                            </div>
                            <div class="col-sm-3">
					            <select class="form-control" name="dob_year" id="dob_year">
                                    <option value="">&nbsp;</option>
                                    <%For i = Year(Date) - 5 To 1950 Step -1%>
								        <option value="<%=i%>"><%=i%></option>
                                    <%Next%>
					            </select>
                            </div>
                        </div>
			            <div class="form-group row">
				            <label class="col-sm-3 form-control-label text-nowrap" for="email">Email:</label>
				            <div class="col-sm-9">
                                <input type="text" class="form-control input-sm" name="email" id="email" value="<%=sEmail%>">
                            </div>
                        </div>
			            <div class="form-group row">
				            <label class="col-sm-3 form-control-label text-nowrap" for="user_name">User Name:</label>
				            <div class="col-sm-9">
                                <input type="text" class="form-control input-sm" name="user_name" id="user_name" value="<%=sUserName%>">
                            </div>
                        </div>
                        <div class="form-group row">
				            <label class="col-sm-3 form-control-label text-nowrap" for="last_name">Password:</label>
				            <div class="col-sm-9">
                                <input type="password" class="form-control input-sm" name="password" id="password">
                            </div>
                        </div>
                        <div class="form-group row">
				            <label class="col-sm-3 form-control-label text-nowrap" for="last_name">Confirm:</label>
				            <div class="col-sm-9">
                                <input type="password" class="form-control input-sm" name="confirm_password" id="confirm_password">
                            </div>
                        </div>
				        <input class="form-control input-sm" type="hidden" name="submit_step3" id="submit_step3" value="submit_step3">
				        <input class="form-control input-sm" type="submit" name="submit3" id="submit3" value="Step 3: Complete Registration">
		                </form>
                <%End Select%>
	        <%End If%>
            <div>
                Performance Tracker is SCHOOL CROSS-COUNTRY/NORDIC SKI PARTICIPANTS ONLY utility that affords subscribers
                the ability to follow and compare the performances of themselves, teammates and competitors.  These competitors are followed via 
                "packs" (a pack can be a single participant or a group of participants that are of the same gender and of the same sport).  
                <br><br>
                Among other features, it allows you to have your <span style="font-weight: bold;">results emailed or texted to you, your parents, and
                friends within a few minutes of finishing your race</span> (NOTE:  ONLY if the event was timed by GSE).
            </div>

            <div class="bg-warning">
                This service carries a one-time fee of $5 to cover administrative and server fees.  You will receive a 7-day free trial which begins when you 
                create your account.
            </div>

            <div class="bg-success">
                This service does not divulge any information that is not already "public" 
                via online results lists.  What it does do is make those results more personal and formats them in a manner that is more informational.
                <br>
                <span style="font-weight: bold;">Please be patient with us.  This utility is a work in progress.</span>
            </div>
        </div>
        <div class="col-md-4">
            <h4 class="h4">What's The Point?</h4>
            <ul class="list-group">
                <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Sends results texts/emails within minutes (GSE-timed events only) to you and whoever you wish to receive them (parents, siblings, etc.).</li>
                <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Create "Performance Packs" of competitors to follow.</li>
                <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Track yours and your competitors' performances.</li>
                <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Graphs of yours and your competitors' performances.</li>
                 <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Access to <a href="http://www.my-etraxc.com">My-eTRaXC</a> training utility.</li>
                <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Enter performances in events not timed by GSE.</li>
                <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Compare performances graphically over time.</li>
                <li class="list-group-item list-group-item-danger" style="padding: 3px 2px 3px 2px;">Network with willing participants.</li>
            </ul>
            <h4 class="h4">About Performance Tracker</h4>
            <iframe class="embed-responsive-item" src="https://www.youtube.com/embed/8pZQckevn70" frameborder="0" allowfullscreen></iframe>
        </div>
        <div class="col-md-4">
            <img src="images/sample_cc.jpg" alt="Sample Picture" class="img-responsive">
            <br>
		    <a href="http://www.etraxc.com/" onclick="openThis2(this.href,1024,760);return false;">
		        <img src="/graphics/banner_ads/etraxc_banner.png" alt="eTRaXC" class="img-responsive">
            </a>
            <hr>
            <a href="http://www.my-etraxc.com/" onclick="openThis2(this.href,1024,760);return false;">
		        <img src="/graphics/my-etraxc_ad.gif" alt="My-eTRaXC" class="img-responsive">
            </a>
            <hr>
            <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
            <!-- GSE Banner Ad -->
            <ins class="adsbygoogle"
                    style="display:inline-block;width:375px;height:90px"
                    data-ad-client="ca-pub-1381996757332572"
                    data-ad-slot="1411231449"></ins>
            <script>
            (adsbygoogle = window.adsbygoogle || []).push({});
            </script>
        </div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" --> 
<%
conn3.Close
Set conn3 = Nothing

conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>
</body>
</html>
