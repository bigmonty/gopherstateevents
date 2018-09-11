<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim lRaceID, lPartID, lEventID
Dim sEventname, sRace, sFirstName, sLastName, sCity, sState, sPhone, sEmail, sGender, sBib, sTwitter, sFbook, sShrtSize
Dim sShowAge, sShowDOB, sShowPhone, sShowCity, sShowSt, sShowEmail, sShowSize, sShowBib, sShowFbook, sShowTwitter, sShowProvider
Dim sShowCell, sShowTeams, sFirstNames, sErrMsg, sEventRaces, sWhereReg
Dim iAgeTab, iDOBTab, iPhoneTab, iCityTab, iStTab, iEmailTab, iSizeTab, iBibTab, iFbookTab, iTwitterTab, iProviderTab, iCellTab, iTeamsTab
Dim iLeapYrs, iAgeDays, iAge
Dim dEventDate, dDOB
Dim RaceArray(), FirstNames, Events()
Dim bFound

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Events(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDate >= '" & Date & "' ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.eOF
	Events(0, i) = rs(0).Value
	Events(1, i) = Replace(rs(1).Value, "''", "'") & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve Events(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_event") = "submit_event" Then
    lEventID = Request.Form.Item("events")
End If

If CStr(lEventID) = vbNullString Then lEventID = 0

i = 0
ReDim FirstNames(0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT DISTINCT FirstName FROM Participant ORDER BY FirstName"
rs.Open sql, conn, 1, 2
sFirstNames = rs.GetString(,,"&#34;&#44;&#34;","&#34;&#44;&#34;","&vbCrLf")
Do While Not rs.EOF
    FirstNames(i) = rs(0).Value
    i = i + 1
    ReDim Preserve FirstNames(i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If CLng(lEventID) > 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SElECT ShowAge, AgeTab, ShowDOB, DOBTab, ShowPhone, PhoneTab, ShowCity, CityTab, ShowSt, StTab, ShowEmail, EmailTab, ShowSize, SizeTab, "
    sql = sql & "ShowBib, BibTab, ShowProvider, ProviderTab, ShowCell, CellTab, ShowTeams, TeamsTab, ShowFbook, FBookTab, "
    sql = sql & "ShowTwitter, TwitterTab FROM PartEntryTabs WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sShowAge = rs(0).value
        iAgeTab = rs(1).Value
        sShowDOB = rs(2).value
        iDOBTab = rs(3).Value
        sShowPhone = rs(4).value
        iPhoneTab = rs(5).Value
        sShowCity = rs(6).value
        iCityTab = rs(7).Value
        sShowSt = rs(8).value
        iStTab = rs(9).Value
        sShowEmail = rs(10).value
        iEmailTab = rs(11).Value
        sShowSize = rs(12).value
        iSizeTab = rs(13).Value
        sShowBib = rs(14).value
        iBibTab = rs(15).Value
        sShowProvider = rs(16).value
        iProviderTab = rs(17).Value
        sShowCell = rs(18).value
        iCellTab = rs(19).Value
        sShowTeams = rs(20).value
        iTeamsTab = rs(21).Value
        sShowFbook = rs(22).value
        iFbookTab = rs(23).Value
        sShowTwitter = rs(24).value
        iTwitterTab = rs(25).Value

        bFound = True
    Else
        bFound = False
    End If
    rs.Close
    Set rs = Nothing

    If bFound = False Then
        sql = "INSERT INTO PartEntryTabs (EventID) VALUES (" & lEventID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        sShowAge = "y"
        iAgeTab = 4
        sShowDOB = "y"
        iDOBTab = 5
        sShowPhone = "y"
        iPhoneTab = 6
        sShowCity = "y"
        iCityTab = 7
        sShowSt = "y"
        iStTab = 8
        sShowEmail = "y"
        iEmailTab = 9
        sShowSize = "y"
        iSizeTab = 10
        sShowBib = "y"
        iBibTab = 11
        sShowProvider = "n"
        iProviderTab = 12
        sShowCell = "n"
        iCellTab = 13
        sShowTeams = "n"
        iTeamsTab = 14
        sShowFbook = "n"
        iFbookTab = 15
        sShowTwitter = "n"
        iTwitterTab = 16
    End If

    'get event information
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    sEventName = rs(0).Value
    dEventDate = rs(1).Value
    rs.Close
    Set rs = Nothing

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

    i = 0
    ReDim RaceArray(1, 0)	
    sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    RaceArray(0, i) = rs(0).value
	    RaceArray(1, i) = rs(1).Value
	    i = i + 1
	    ReDim Preserve RaceArray(1, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing

    If UBound(RaceArray, 2) = 1 Then 
        lRaceID = RaceArray(0, 0)
    Else
        If CStr(lRaceID) = vbNullString Then lRaceID = 0
    End If
End If

If Request.Form.Item("submit_part") = "submit_part" Then
	If Clng(lRaceID) = 0 Then lRaceID = Request.Form.Item("races")
	
	sFirstName = Replace(Request.Form.Item("first_name"),"'","''")
	sLastName = Replace(Request.Form.Item("last_name"),"'","''")
	sGender = Request.Form.Item("gender")
	sCity = Replace(Request.Form.Item("city"),"'","''")
	sState = Trim(Request.Form.Item("state"))
	sPhone = Request.Form.Item("phone1") & "-" & Request.Form.Item("phone2") & "-"  & Request.Form.Item("phone3")
	sEmail = Request.Form.Item("email")
	dDOB = Request.Form.Item("dob")
	sBib = Request.Form.Item("bib")
	sTwitter = Request.Form.Item("twitter")
	sFbook = Request.Form.Item("fbook")
    sShrtSize = Request.Form.Item("shrt_size")
        	
	If Request.Form.Item("age") = vbNullString Then
		iAge = RaceDayAge()
	Else
		iAge = Request.Form.Item("age")
	End If
    
    'check for duplicate bib
    sErrMsg = vbNullString

    If sBib = vbNullString Then sErrMsg = "You must include a bib number."

    If sErrMsg = vbNullString Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Bib FROM PartRace WHERE Bib = '" & sBib & "' AND RaceID IN (" & sEventRaces & ")"
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then sErrMsg = "I'm sorry.  This bib has already been assigned in this event."
        rs.Close
        Set rs = Nothing
    End If

    If sErrMsg = vbNullString Then
        'insert into the participant table
        sql = "INSERT INTO Participant (FirstName, LastName, City, St, Phone, Email, DOB, Gender, Twitter, Fbook)VALUES ('"
        sql = sql & sFirstName & "', '" & sLastName & "', '" & sCity & "', '" & sState & "', '" & sPhone & "', '" & sEmail
        sql = sql & "', '" & dDOB & "', '" & sGender & "', '" & sTwitter & "', '" & sFbook & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
        
        'get ParticipantID
        sql = "SELECT ParticipantID FROM Participant WHERE (FirstName = '" & sFirstName & "' AND LastName = '"
        sql = sql & sLastName & "' AND City = '" & sCity & "' AND St = '" & sState & "' AND Phone = '" & sPhone & "' AND Email = '"
        sql = sql & sEmail & "' AND Gender = '" & sGender & "')"
        Set rs = conn.Execute(sql)
        lPartID = rs(0).Value
        Set rs = Nothing
	
	    'insert into part reg table
        If CDate(dEventDate) >= Date - 1 Then
            sWhereReg = "Race Day"
        Else
            sWhereReg = "Mail In"
        End If

	    sql = "INSERT INTO PartReg (ParticipantID, WhereReg, DateReg, RaceID, ShrtSize) VALUES (" & lPartID & ", '" & sWhereReg & "', '" & Date & "', " 
        sql = sql & lRaceID & ", '" & sShrtSize & "')"
	    Set rs=conn.Execute(sql)
	    Set rs=Nothing

	    'insert into part race table
	    sql = "INSERT INTO PartRace (ParticipantID, Age, RaceID, Bib, AgeGrp) VALUES (" & lPartID & ", " 
	    sql = sql & CInt(iAge) & ", " & lRaceID & ", '" & sBib & "', '" & GetAgeGrp(sGender, iAge, lRaceID) & "')"
	    Set rs=conn.Execute(sql)
	    Set rs=Nothing

        sFirstName = vbNullString
        sLastName = vbNullString
        sBib = vbNullString
        sGender = vbNullString
        dDOB = vbNullString
        sCity = vbNullString
        sEmail = vbNullString 
        sState = vbNullString
        sPhone = vbNullString
        sTwitter = vbNullString
        sFbook = vbNullString
        sShrtSize = vbNullString
        iAge = vbNullString
    End If
End If

'get participants age on race day
Function RaceDayAge()
    'first get leap years
    iLeapYrs = 0
    For i = Year(CDate(dEventDate)) To Year(CDate(dDOB)) Step -1
        If i / 4 = i \ 4 Then iLeapYrs = iLeapYrs + 1
    Next
        
    iAgeDays = DateDiff("d", CDate(dDOB), CDate(dEventDate))
    iAgeDays = iAgeDays - iLeapYrs
    RaceDayAge = iAgeDays \ 365
End Function

Public Function GetAgeGrp(sMF, iAge, lThisRaceID)
    Dim iBegAge, iEndAge
    
    iBegAge = 0
    
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT EndAge FROM AgeGroups WHERE Gender = '" & sMF & "' AND RaceID = " & lThisRaceID & " ORDER BY EndAge DESC"
    rs2.Open sql2, conn, 1, 2
    Do While Not rs2.EOF
        If CInt(iAge) <= CInt(rs2(0).Value) Then
            iEndAge = rs2(0).Value
        Else
            iBegAge = CInt(rs2(0).Value) + 1
            Exit Do
        End If
        rs2.MoveNext
    Loop
    rs2.Close
    Set rs2 = Nothing

    If iBegAge = 0 Then
        GetAgeGrp = iEndAge & " and Under"
    Else
        If iEndAge = 110 Then
            GetAgeGrp = CInt(iBegAge) & " and Over"
        Else
            GetAgeGrp = CInt(iBegAge) & " - " & iEndAge
        End If
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Enter <%=sEventName%> Participants</title>

<script>
  $(function() {
    var availableTags = [];
    <%
        For i = 0 To UBound(FirstNames)
            ' The Replace() is to escape apostrophes
            Response.Write( "availableTags.push('" & Replace(FirstNames(i), "'", "\'") & "');" & vbCrLf )
        Next
    %>    
    $( "#first_name" ).autocomplete({
      source: availableTags
    });
  });
</script>

<script>
function checkFields() {
 	if (document.part_reg.first_name.value == '' || 
 	    document.part_reg.last_name.value == '' ||
	 	document.part_reg.gender.value == ''|| 
        document.part_reg.bib.value == ''|| 
        document.part_reg.races.value == ''||
	 	(document.part_reg.age.value == ''&&
	 	document.part_reg.dob.value == ''))
		{
  		alert('Please fill in all required fields-they are in red!  This includes EITHER age or DOB');
  		return false
  		}
 	else
		if (isNaN(document.part_reg.bib.value))
    		{
			alert('The bib number field can not contain non-numeric values');
			return false
			}
	else
   		return true
}
</script>

</head>

<body onload="javascript:part_reg.first_name.focus()">
<div class="container">
    <h3 class="h3">GSE Enter <%=sEventName%> Participants</h3>
		
    <div>
        This page is designed for quick participant data input and is designed to be largely mouse-free.  Just tab from field-to-field and when
        all data is entered, hit the "enter" key on your keyboard.  The form will reset and you are ready for the next participant.
    </div>	

    <%If Not sErrMsg = vbNullString Then%>
        <br>
        <div class="bg-info"><%=sErrMsg%></div>
    <%End If%>

    <div>
        <a href="enter_parts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">Refresh Form</a>
    </div>

    <div class="row">
         <form role="form" class="form-inline" name="which_event" method="post" action="enter_parts.asp?event_id=<%=lEventID%>">
        <div class="form-group">
            <label for="events">Select Event:</label>
            <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()" tabindex="18">
                <option value="">&nbsp;</option>
                <%For i = 0 to UBound(Events, 2) - 1%>
                    <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
                        <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
                    <%Else%>
                        <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                    <%End If%>
                <%Next%>
            </select>
        </div>
        <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
        <input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event" tabindex="19">
        </form>
   </div>

    <%If CLng(lEventID) > 0 Then%>
	    <form role="form" class="form" name="part_reg" method="post" action="enter_parts.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>" onSubmit="return checkFields();">
	    <div class="row">
            <div class="col-sm-6">
                <div class="form-group">
                    <label for="first_name" class="text-danger">First Name:</label>
                    <input class="form-control" name="first_name" id="first_name" maxLength="30" tabindex="1" value="<%=sFirstName%>">
                </div>
                <div class="form-group">
                    <label for="gender" class="text-danger">Gender:	</label>
                    <select class="form-control" name="gender" id="gender" tabindex="3"> 
                        <option value="">&nbsp;</option> 
                        <%Select Case sGender%>
                            <%Case "M"%>
                                <option value="M" selected>M</option> 
                                <option value="F">F</option>
                                <option value="X">X</option>
                            <%Case "F"%>
                                <option value="M">M</option> 
                                <option value="F" selected>F</option>
                                <option value="X">X</option>
                            <%Case Else%>
                                <option value="M">M</option> 
                                <option value="F">F</option>
                                <option value="X">X</option>
                        <%End Select%>
                    </select>
                </div>
                <%If sShowAge = "y" Then%>
                    <div class="form-group">
                        <label for="age" class="text-danger">Age:</label>
                        <input class="form-control" name="age" id="age" maxlength="2" tabindex="5" value="<%=iAge%>">
                    </div>
                <%End If%>
                <%If sShowCity = "y" Then%>
                    <div class="form-group">
                        <label for="city">City:</label>
                        <input class="form-control" name="city" id="city" maxLength="30" tabindex="7" value="<%=sCity%>">
                    </div>
                <%End If%>
                <%If sShowEmail = "y" Then%>
                    <div class="form-group">
                        <label for="email">Email:</label>
                        <input class="form-control" name="email" id="email" maxLength="50" tabindex="9" value="<%=sEmail%>">
                    </div>
                <%End If%>
                <%If sShowDob = "y" Then%>
                    <div class="form-group">
                        <label for="dob" class="text-danger">DOB:</label>
                        <input class="form-control" name="dob" id="dob" maxLength="10" tabindex="<%=iDobTab%>" value="<%=dDOB%>">
                    </div>
                <%End If%>
                <%If sShowSize = "y" Then%>
                    <div class="form-group">
                        <label for="shrt_size">Shrt Size:</label>
                        <input class="form-control" name="shrt_size" id="shrt_size" maxLength="4" tabindex="<%=iSizeTab%>" value="<%=sShrtSize%>">
                    </div>
                <%End If%>
                <%If sShowFbook = "y" Then%>
                    <div class="form-group">
                        <label for="fbook">FBook:</label>
                        <input class="form-control" name="fbook" id="fbook" maxLength="30" tabindex="<%=iFbookTab%>">
                    </div>
                <%End If%>
                <%If sShowTwitter = "y" Then%>
                    <div class="form-group">
                        <label>Twitter:</label>
                        <input name="twitter" id="twitter" size="15" maxLength="30" tabindex="<%=iTwitterTab%>">
                    </div>
                <%End If%>
                <%If sShowProvider = "y" Then%>
                    <div class="form-group">
                        <label for="provider">Cell Provider:</label>
                        <select class="form-control" name="provider" id="provider"  tabindex="<%=iProviderTab%>">
                            <option value="">&nbsp;</option>
                        </select>
                    </div>
                <%End If%>
                <%If sShowCell = "y" Then%>
                    <div class="form-group">
                        <label for="cell_phone">Cell Phone:</label>
                        <input class="form-control" name="cell_phone" id="cell_phone" maxLength="30" tabindex="<%=iCellTab%>">
                    </div>
                    <%End If%>
                <%If sShowTeams = "y" Then%>
                    <div class="form-group">
                        <label for="teams">Team:</label>
                        <select name="teams" id="teams"  tabindex="<%=iTeamsTab%>">
                            <option value="">&nbsp;</option>
                        </select>
                    </div>
                <%End If%>
            </div>
            <div class="col-sm-6">
                <div class="form-group">
                    <label for="last_name" class="text-danger">Last Name:</label>
                    <input class="form-control" name="last_name" id="last_name" maxLength="30" tabindex="2" value="<%=sLastName%>">
                </div>
                <div class="form-group">
                    <label for="races" class="text-danger">Select Race:</label>
                    <select class="form-control" name="races" id="races" tabindex="4">
                        <option value="">&nbsp;</option>
                        <%For i = 0 to UBound(RaceArray, 2) - 1%>
                            <%If CLng(lRaceID) = CLng(RaceArray(0, i)) Then%>
                                <option value="<%=RaceArray(0, i)%>" selected><%=RaceArray(1, i)%></option>
                            <%Else%>
                                <option value="<%=RaceArray(0, i)%>"><%=RaceArray(1, i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </div>
                <%If sShowBib = "y" Then%>
                    <div class="form-group">
                        <label for="bib" class="text-danger">Bib:</label>
                        <input class="form-control" name="bib" id="bib" maxlength="4" tabindex="6" value="<%=sBib%>">
                    </div>
                <%End If%>
                <%If sShowSt = "y" Then%>
                    <div class="form-group">
                        <label for="state">State:</label>
                        <input class="form-control" name="state" id="state" maxLength="2" tabindex="8" value="<%=sState%>">
                    </div>
                <%End If%>
                <%If sShowPhone = "y" Then%>
                    <div class="form-group">
                        <label for="phone">Phone:</label>
                        <input class="form-control" name="phone" id="phone" maxLength="12" tabindex="<%=iPhoneTab%>" value="<%=sPhone%>">
                    </div>
                <%End If%>
            </div>
            <div class="form-group">
                <input type="hidden" name="submit_part" id="submit_part" value="submit_part">
                <input type="submit" class="form-control" name="submit1" id="submit1" value="Enter Participant" tabindex="17"> 
            </div>
            </form>
        </div>
    <%End If%>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>