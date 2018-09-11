<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim i
Dim lPartID, lMyHistID, lEtraxcID
Dim sErrMsg, sMsg, sFirstName, sLastName, sGender, sEmail, sUserName, sPassword, sConfirmPassword
Dim cdoMessage, cdoConfig
Dim dDOB

lPartID = Request.QueryString("part_id")
If CStr(lPartID) & "" = "" Then lPartID = 0

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"

Dim sRandPic
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, PixName FROM RacePix ORDER BY NEWID()"
rs.Open sql, conn, 1, 2
sRandPic = "/gallery/" & rs(0).Value & "/" & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_accnt") = "submit_accnt" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'create_accnt' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_create_accnt") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_create_accnt") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg
		
		sFirstName = CleanInput(Trim(Request.Form.Item("first_name")))
		If sHackMsg = vbNullString Then sLastName = CleanInput(Trim(Request.Form.Item("last_name")))
	    If sHackMsg = vbNullString Then sGender = CleanInput(Trim(Request.Form.Item("gender")))
	    If sHackMsg = vbNullString Then sEmail = CleanInput(Trim(Request.Form.Item("email")))
	    If sHackMsg = vbNullString Then sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
	    If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))
	    If sHackMsg = vbNullString Then sConfirmPassword = CleanInput(Trim(Request.Form.Item("confirm_password")))
        If sHackMsg = vbNullString Then dDOB = Request.Form.Item("dob")

        If sHackMsg = vbNullString Then
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
                'check for email validity
                If ValidEmail(sEmail) = False Then sErrMsg = "Your email address does not appear to be in a valid format.  Please re-enter."
            End If

            If sErrMsg = vbNullString Then
                'check for email validity
                If Not IsDate(dDOB) Then sErrMsg = "Your date of birth is not a valid date.  Please re-enter."
            End If

            'if everything checks out
            If sErrMsg = vbNullString Then
                sFirstName = Replace(sFirstName, "'", "''")
                sLastName = Replace(sLastName, "'", "''")
  
                'first create participantID
                If CLng(lPartID) = 0 Then
                    'insert into the participant table
                    sql = "INSERT INTO Participant (FirstName, LastName, DOB, Gender)VALUES ('"
                    sql = sql & sFirstName & "', '" & sLastName & "', '" & dDOB & "', '" & sGender & "')"
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing
        
                    'get ParticipantID
                    sql = "SELECT ParticipantID FROM Participant WHERE (FirstName = '" & sFirstName & "' AND LastName = '"
                    sql = sql & sLastName & "' AND Gender = '" & sGender & "' AND DOB  = '" & dDOB & "')"
                    Set rs = conn.Execute(sql)
                    lPartID = rs(0).Value
                    Set rs = Nothing
                End If

  	            'insert into my hist table
	            sql = "INSERT INTO MyHist (UserName, Password, Email, WhenCreated, ParticipantID) VALUES ('" & sUserName & "', '" & sPassword & "', '" 
                sql = sql & sEmail & "', '" & Now() & "', " & lPartID & ")"
	            Set rs=conn.Execute(sql)
	            Set rs=Nothing

	            'get my hist id
	            sql = "SELECT MyHistID FROM MyHist WHERE UserName = '" & sUserName & "' AND Password = '" & sPassword & "'"
	            Set rs = conn.Execute(sql)
	            lMyHistID = rs(0).Value
	            Set rs=Nothing

                'set up my-etraxc account
                Dim lIndSubscrID
                Dim iPIN, iCurrSeas
                Dim sngPace
                Dim PinArray()
                Dim bFound
                Dim dSeasStart

	            sngPace = Request.Form("mintomiles")
	
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
			    sql = sql & "BirthDate, Archive, MyHistID) VALUES (225, '" & sFirstName & "', '" & sLastName & "', 0, '" & sGender & "', '" 
			    sql = sql & sEmail & "', " & iPIN & ", " & sngPace & ", '" & sUserName & "', '" & dDOB & "', 'N', " & lMyHistID & ")"
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

	            sMsg = "Welcome to GSE's My History Program!" & vbCrLf & vbCrLf
	            sMsg = sMsg & "The purpose of this program is to assist you in managing your training, racing, and other components of your "
	            sMsg = sMsg & "journey towards health and wellness.  We hope you will find the program to be a valuable asset to you and "
                sMsg = sMsg & "would ask that you provide feedback about additional functionality that you would like to see. " & vbCrLf & vbCrLf
	
                sMsg = sMsg & "By creating your My History account through GSE you have also received an account at http://www.my-etraxc.com/ "
                sMsg = sMsg & "My-eTRaXC is a FREE online fitness manager and training log for individuals.  It renders well on mobile devices "
                sMsg = sMsg & "and has more in terms of features and functionality than we could list here.  You are encouraged to log in to "
                sMsg = sMsg & "My-eTRaXC using your user name and the PIN listed below and look around."  & vbCrLf & vbCrLf
	
	            sMsg = sMsg & "Your Name: "  & sFirstName & " " & sLastName & vbCrLf
	            sMsg = sMsg & "Gender: "  & sGender  & vbCrLf
	            sMsg = sMsg & "Date of Birth: "  & dDOB  & vbCrLf
            	sMsg = sMsg & "My-eTRaXC PIN: " & iPIN  & vbCrLf  & vbCrLf

%>
<!--#include file = "../includes/cdo_connect.asp" -->
<%
	
	            Set cdoMessage = CreateObject("CDO.Message")
	            With cdoMessage
		            Set .Configuration = cdoConfig
		            .To = sEmail
		            .From = "bob.schneider@gopherstateevents.com"
	                .CC = "bob.schneider@gopherstateevents.com;"
	                .Subject = "Your My GSE History Account"
		            .TextBody = sMsg
		            .Send
	            End With
	            Set cdoMessage = Nothing

                Response.Redirect "my_hist.asp"
            End If
        End If
    End If
End If
		
If CLng(lPartID) > 0 Then
	sql = "SELECT FirstName, LastName, Gender, Email FROM Participant WHERE ParticipantID = " & lPartID
	Set rs = conn.Execute(sql)
	sFirstName = Replace(rs(0).Value, "''", "'") 
    sLastName  = Replace(rs(1).Value, "''", "'")
	sGender = rs(2).Value
    sEmail = rs(3).Value
	Set rs = Nothing
End If

'log this user if they are just entering the site
If Session("access_create_accnt") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'create_accnt')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'create_accnt'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_create_accnt")
End If

%>
<!--#include file = "../includes/clean_input.asp" -->

<!--#include file = "../includes/valid_email.asp" -->
<%

Function ValidUserName(sThisUserName) 
	ValidUserName = True

	If Len(sThisUserName) < 5 Or Len(sThisUserName) > 12 Then ValidUserName = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UserName FROM MyHist WHERE UserName = '" & sThisUserName & "'"
    rs.open sql, conn, 1, 2
    If rs.RecordCount > 0 Then ValidUserName = False
    rs.Close
    Set rs = Nothing
End Function

Function ValidPassword(sThisPassword) 
	ValidPassword = True

	If Len(sThisPassword) < 5 Or Len(sThisPassword) > 12 Then ValidPassword = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Password FROM MyHist WHERE Password = '" & sThisPassword & "'"
    rs.open sql, conn, 1, 2
    If rs.RecordCount > 0 Then ValidPassword = False
    rs.Close
    Set rs = Nothing
End Function

Function MinsSecs(sngPace)
	Select Case Right(sngPace, 2) 
		Case "75"
			MinsSecs = CStr(CSng(sngPace) - CSng(.75)) & ":45"
		Case "25"
			MinsSecs = CStr(CSng(sngPace) - CSng(.25)) & ":15"
		Case ".5"
			MinsSecs = CStr(CSng(sngPace) - CSng(.5)) & ":30"
		Case Else
			MinsSecs = sngPace & ":00"
	End Select	
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
<!--#include file = "../includes/meta2.asp" -->
<title>Create A My GSE&copy; Account</title>
<meta name="description" content="Create my history account for Gopher State Events (GSE).">
<!--#include file = "../includes/js.asp" -->

<script>
function chkFlds() {
if (document.new_accnt.first_name.value == '' || 
    document.new_accnt.last_name.value == '' || 
    document.new_accnt.gender.value == '' || 
    document.new_accnt.email.value == '' ||
    document.new_accnt.dob.value == '' || 
    document.new_accnt.user_name.value == '' || 
    document.new_accnt.confirm_password.value == '' || 
    document.new_accnt.password.value == '' )
{
 	alert('All fields are required!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body onload="document.new_accnt.first_name.focus();">
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Create My GSE History Account">
    <h3 class="h3">My GSE History</h3>

    <div class="bg-warning">
        <ul  class="list-inline">
            <li><a href="my_hist.asp">Sign In</a></li>
        </ul>
    </div>
		
    <h4 class="h4">Create My GSE History Account <small>(all fields are required)</small></h4>

    <div class="bg-warning">
        With GSE's My History app you  can use the full force of the web to track your racing, manage your fitness and 
        training, plan future races, and otherwise live an active life full of energy and exhuberance.  This FREE service coordinates with
        <a href="http://www.my-etraxc.com/" style="font-weight: bold;">My-eTRaXC</a>, an online training, lifestyle, and record keeping account.  
    </div>
		
    <br>
    	
	<div class="col-md-7">
		<%If Not sHackMsg = vbNullString Then%>
			<div class="bg-danger"><%=sHackMsg%></div>
		<%Else%>
			<%If Not sErrMsg = vbNullString Then%>
				<div class="bg-danger"><%=sErrMsg%></div>
			<%End If%>
				
  			<form role="form" class="form" name="new_accnt" method="Post" action="create_accnt.asp?part_id=<%=lPartID%>" onSubmit="return chkFlds();">
            <div class="row">
                <div class="col-sm-6">
			        <div class="form-group">
				        <label for="first_name">First Name:</label>
				        <input type="text" class="form-control" name="first_name" id="first_name" maxLength="30" value="<%=sFirstName%>" tabindex="1">
			        </div>
			        <div class="form-group">
				        <label for="gender">Gender:</label>
				        <select class="form-control" name="gender" id="gender" tabindex="3">
                            <option value="">&nbsp;</option>
                            <%Select Case UCase(sGender)%>
                                <%Case "M"%>
							        <option value="M" selected>M</option>
							        <option value="F">F</option>
                                <%Case "F"%>
							        <option value="M">M</option>
							        <option value="F" selected>F</option>
                                <%Case Else%>
							        <option value="M">M</option>
							        <option value="F">F</option>
                            <%End Select%>
				        </select>
			        </div>
			        <div class="form-group">
				        <label for="dob">Date of Birth:</label>
				        <input type="text" class="form-control" name="dob" id="dob" value="<%=dDOB%>" tabindex="5">
			        </div>
			        <div class="form-group">
				        <label for="user_name">User Name:</label>
				        <input type="text" class="form-control" name="user_name" id="user_name" maxLength="12" size="12" value="<%=sUserName%>" tabindex="7">
			        </div>
                </div>
                <div class="col-sm-6">
			        <div class="form-group">
				        <label for="last_name">Last Name:</label>
				        <input type="text" class="form-control" name="last_name" id="last_name" maxLength="30" value="<%=sLastName%>" tabindex="2">
			        </div>
			        <div class="form-group">
				        <label for="email">Email:</label>
				        <input type="text" class="form-control" name="email" id="email" maxLength="50" size="25"value="<%=sEmail%>" tabindex="4"> 
			        </div>
                    <div class="form-group">
				        <label for="mintomiles">Typical Training Pace:</label>
				        <select class="form-control" name="mintomiles" id="mintomiles" tabindex="6">
					        <option value="">&nbsp;</option>
					        <%For i = 5 to 12 Step .25%>
						        <%If CSng(sngPace) = CSng(i) Then%>
			 				        <option value="<%=i%>" selected><%=MinsSecs(i)%></option>
						        <%Else%>
			 				        <option value="<%=i%>"><%=MinsSecs(i)%></option>
			 			        <%End If%>
			 		        <%Next%>
				        </select>
                    </div>
			        <div class="form-group">
				        <label for="password">Password:</label>
				        <input type="password" class="form-control" name="password" id="password" maxLength="12" size="12" tabindex="8">
			        </div>
			        <div class="form-group">
				        <label for="confirm_password">Confirm:</label>
				        <input type="password" class="form-control" name="confirm_password" id="confirm_password" maxLength="12" size="12" tabindex="9">
			        </div>
                </div>
                <div class="row">
			        <div class="form-group">
				        <input type="hidden" class="form-control" name="submit_accnt" id="submit_accnt" value="submit_accnt">
				        <input type="submit" class="form-control" name="submit1" id="submit1" value="Create My Account" tabindex="10">
			        </div>
                </div>
            </div>
			</form>
		<%End If%>
    </div>
    <div class="col-md-5">
		<a href="<%=sRandPic%>" onclick="openThis2(this.href,1024,768);return false;">
            <img src="<%=sRandPic%>" alt="<%=sRandPic%>" class="img-responsive">
        </a>
    </div>
</div>
<!--#include file = "../includes/footer.asp" -->
<%
conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>
</body>
</html>
