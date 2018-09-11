<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lThisMeet, lThisClass
Dim sMeetName, sClassName, sGender, sDetails, sDeleteClass, sCourseMap, sMapLink, sMeetInfoSheet, sMeetSite, sMeetHost, sWebsite, sComments, sEntryFee
Dim Races(), MeetTeams(), MeetClasses(), ThisClass(3), MeetArr()
Dim dMeetDate, dWhenShutdown

If Not Session("role") = "meet_dir" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")
lThisClass = Request.QueryString("class_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_meet") = "submit_meet" Then 
    lThisMeet = Request.Form.Item("meets")
ElseIf Request.Form.Item("submit_class") = "submit_class" Then
    sClassName = Request.Form.Item("class_name")
    sGender = Request.Form.Item("gender")
    sDetails = Request.Form.Item("details")

    sql = "INSERT INTO MeetClasses (MeetsID, ClassName, Gender, Details) VALUES (" & lThisMeet & ", '" & sClassName & "', '" & sGender
    sql = sql & "', '" & sDetails & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_select") = "submit_select" Then
    lThisClass = Request.Form.Item("classes")
ElseIf Request.Form.Item("submit_edit") = "submit_edit" Then
    sClassName = Request.Form.Item("edit_class_name")
    sGender = Request.Form.Item("edit_gender")
    sDetails = Request.Form.Item("edit_details")
    sDeleteClass = Request.Form.Item("delete_class")

    If sDeleteClass = "on" Then
        sql = "DELETE FROM MeetClasses WHERE MeetClassesID = " & lThisClass
        Set rs = conn.Execute(sql)
        SEt rs = Nothing
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ClassName, Gender, Details FROM MeetClasses WHERE MeetClassesID = " & lThisClass 
        rs.Open sql, conn, 1, 2
        If sClassName & "" = "" Then
            rs(0).Value = rs(0).OriginalValue
        Else
            rs(0).Value = sClassName
        End If
        rs(1).Value = sGender
        If sDetails & "" = "" Then
            rs(2).Value = rs(2).OriginalValue
        Else
            rs(2).Value = sDetails
        End If
        rs.Update
        rs.Close
        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_races") = "submit_races" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RacesID, RaceClass FROM Races WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    If Request.Form.Item("race_classes_" & rs(0).Value) & "" = "" Then
            rs(1).Value = Null
        Else
            rs(1).Value = Request.Form.Item("race_classes_" & rs(0).Value)
        End If
        rs.Update
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_teams") = "submit_teams" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID, MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    If Request.Form.Item("team_classes_" & rs(0).Value) & "" = "" Then
            rs(1).Value = Null
        Else
            rs(1).Value = Request.Form.Item("team_classes_" & rs(0).Value)
        End If
        rs.Update
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

If CStr(lThisClass) = vbNullString Then lThisClass = 0
If CStr(lThisMeet) = vbNullString Then lThisMeet = 0

i = 0
ReDim MeetArr(1, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetDirID = " & Session("my_id") & " ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetArr(0, i) = rs(0).Value
	MeetArr(1, i) = rs(1).Value & " " & Year(rs(2).Value)
	i = i + 1
	ReDim Preserve MeetArr(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If UBound(MeetArr, 2) = 1 Then lThisMeet = MeetArr(0, 0)

If Not CLng(lThisMeet) = 0 Then
	sql = "SELECT MeetName, MeetDate, MeetSite, MeetHost, WebSite, Comments, EntryFee, WhenShutdown FROM Meets WHERE MeetsID = " & lThisMeet
	Set rs = conn.Execute(sql)
	sMeetName = Replace(rs(0).Value, "''", "'")
	dMeetDate = rs(1).Value
	If Not rs(2).Value & "" = "" Then sMeetSite = Replace(rs(2).Value, "''", "'")
	If Not rs(3).Value & "" = "" Then sMeetHost = Replace(rs(3).Value, "''", "'")
	sWebSite = rs(4).Value
	If Not rs(5).Value & "" = "" Then sComments = Replace(rs(5).Value, "''", "'")
	sEntryFee = rs(6).Value
	dWhenShutdown = rs(7).Value
	Set rs = Nothing
	
	'get maplink
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT MapLink FROM MapLinks WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMapLink = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get meet info sheet
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT InfoSheet FROM MeetInfo WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sMeetInfoSheet = rs(0).Value
	rs.Close
	Set rs = Nothing
	
	'get course map
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Map FROM CourseMap WHERE MeetsID = " & lThisMeet
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then sCourseMap = rs(0).Value
	rs.Close
	Set rs = Nothing

    i = 0
    ReDim MeetClasses(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClassesID, ClassName, Gender, Details FROM MeetClasses WHERE MeetsID = " & lThisMeet 
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MeetClasses(0, i) = rs(0).Value
        MeetClasses(1, i) = Replace(rs(1).Value, "''", "'")
        MeetClasses(2, i) = rs(2).Value
        MeetClasses(3, i) = Replace(rs(3).Value, "''", "'")
        i = i + 1
        ReDim Preserve MeetClasses(3, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    'get meet teams array
    i = 0
    ReDim MeetTeams(3, 0)
    sql = "SELECT mt.TeamsID, t.TeamName, t.Gender, mt.MeetClass FROM MeetTeams mt INNER JOIN Teams t ON mt.TeamsID = t.TeamsID "
    sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
    Set rs = conn.Execute(sql)
    Do While Not rs.EOF
	    MeetTeams(0,  i) = rs(0).Value
	    MeetTeams(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
        MeetTeams(2, i) = rs(2).Value
        If rs(3).Value & "" = "" Then
            MeetTeams(3, i) = 0
        Else
            MeetTeams(3, i) = rs(3).Value
        End If

	    i = i + 1
	    ReDim Preserve MeetTeams(3, i)
	    rs.MoveNext
    Loop
    Set rs = Nothing

    'get races in this meet
    i = 0
    ReDim Races(2, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT RacesID, RaceDesc, RaceClass FROM Races WHERE MeetsID = " & lThisMeet & " ORDER BY RaceTime"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
	    Races(0, i) = rs(0).Value
        Races(1, i) = Replace(rs(1).Value, "''", "'")
        If rs(2).Value & "" = "" Then
            Races(2, i) = 0
        Else
            Races(2, i) = rs(2).Value
        End If
	    i = i + 1
	    ReDim Preserve Races(2, i)
	    rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

If Not CLng(lThisClass) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT MeetClassesID, ClassName, Gender, Details FROM MeetClasses WHERE MeetClassesID = " & lThisClass 
    rs.Open sql, conn, 1, 2
    ThisClass(0) = rs(0).Value
    ThisClass(1) = Replace(rs(1).Value, "''", "'")
    ThisClass(2) = rs(2).Value
    ThisClass(3) = Replace(rs(3).Value, "''", "'")
    rs.Close
    Set rs = Nothing
End If

'migrate classes from last year
'Dim iTeamClass
'For i = 0 To UBound(MeetTeams, 2) - 1   
'    iTeamClass = 0
'    Set rs = Server.CreateObject("ADODB.Recordset")
'    sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = 285 AND TeamsID = " & MeetTeams(0, i)
'    rs.Open sql, conn, 1, 2
'    If rs.RecordCount > 0 Then  'make sure the team was in the meet last year
'        If Not rs(0).Value & "" = "" Then iTeamClass = CLng(rs(0).Value) + 4    'because this years corresponding classes are all +4 from last years classes
'    End If
'    rs.Close
'    Set rs = Nothing

'    If CLng(iTeamClass) > 0 Then
'        Set rs = Server.CreateObject("ADODB.Recordset")
'        sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = 330 AND TeamsID = " & MeetTeams(0, i)
'        rs.Open sql, conn, 1, 2
'        rs(0).Value = iTeamClass
'        rs.Update
'        rs.Close
'        Set rs = Nothing
        
'        MeetTeams(3, i) = iTeamClass
'    End If
'Next
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>GSE  Admin CC/Nordic Meet Classes</title>
<!--#include file = "../../../../includes/js.asp" -->

<script>
function chkFlds(){
 	if (document.add_class.class_name.value == '' || 
 	    document.add_class.gender.value == '' ||
 	    document.add_class.details.value == '')
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
    <!--#include file = "../../../../includes/header.asp" -->
	<!--#include file = "../../../../includes/meet_dir_menu.asp" -->
	<h4 class="h4">Meet Classes</h4>

	<form class="form-inline" name="get_meets" method="post" action="meet_classes.asp">
	<label for="meets">Select Meet:</label>
	<select class="form-control" name="meets" id="meets" onchange="this.form.submit1.click();">
		<option value="">&nbsp;</option>
		<%For i = 0 to UBound(MeetArr, 2) - 1%>
			<%If CLng(lThisMeet) = CLng(MeetArr(0, i)) Then%>
				<option value="<%=MeetArr(0, i)%>" selected><%=MeetArr(1, i)%></option>
			<%Else%>
				<option value="<%=MeetArr(0, i)%>"><%=MeetArr(1, i)%></option>
			<%End If%>
		<%Next%>
	</select>
	<input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
	<input type="submit" class="form-control" name="submit1" id="submit1" value="Get This">
	</form>

	<%If Not CLng(lThisMeet) = 0 Then%>
		<!--#include file = "../../meet_dir_nav.asp" -->
		<div class="bg-info">
            <h5 class="h5">Assigning Meet Classes to an event is a multi-step process:</h5>
		    <ol>
                <li>Create the classes</li>
                <li>Designate which race goes in which class (not all races need to be determined by class)</li>
                <li>Designate a class for each team.  Failure to designate a class for a team will prevent them from entering athletes in any race
                    that has a class designation.</li>
            </ol>
        </div>

        <div class="col-xs-4">
            <h4 class="h4">Create Meet Classes:</h4>

            <form class="form" name="add_class" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkFlds();">
            <table>
                <tr>
                    <th>Class Name:</th>
                    <td><input type="text" class="form-control" name="class_name" id="class_name"></td>
                </tr>
                <tr>
                    <th>Gender:</th>
                    <td>
                        <select class="form-control" name="gender" id="gender">
                            <option value="">&nbsp;</option>
                            <option value="Both">Both</option>
                            <option value="Male">Male</option>
                            <option value="Female">Female</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <th>Details:</th>
                    <td><textarea class="form-control" name="details" id="details"rows="5"></textarea></td>
                </tr>
                <tr>
                    <td style="text-align: center;" colspan="4">
                        <input type="hidden" name="submit_class" id="submit_class" value="submit_class">
                        <input type="submit" class="form-control" name="submit1" id="submit1" value="Submit Class">
                    </td>
                </tr>
            </table>
            </form>

            <br>

            <h4 class="h4">Edit/Delete Meet Classes:</h4>

            <form class="form" name="select_class" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
            <table>
                <tr>
                    <th>Select Class To Edit/Delete:</th>
                    <td>
                        <select class="form-control" name="classes" id="classes" onchange="this.form.submit5.click();">
                            <option value="">&nbsp;</option>
                            <%For i = 0 To UBound(MeetClasses, 2) - 1%>
                                <%If CLng(lThisClass) = CLng(MeetClasses(0, i)) Then%>
                                    <option value="<%=MeetClasses(0, i)%>" selected><%=MeetClasses(1, i)%></option>
                                <%Else%>
                                    <option value="<%=MeetClasses(0, i)%>"><%=MeetClasses(1, i)%></option>
                                <%End If%>
                            <%Next%>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td style="text-align: center;"  colspan="2">
                        <input type="hidden" name="submit_select" id="submit_select" value="submit_select">
                        <input type="submit" class="form-control" name="submit5" id="submit5" value="Select Class">
                    </td>
                </tr>
            </table>
            </form>

            <%If Not CLng(lThisClass) = 0 Then%>
                <br><br>
                <form class="form" name="edit_class" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>&amp;class_id=<%=lThisClass%>">
                <table>
                    <tr>
                        <th>Class Name:</th>
                        <td><input type="text" class="form-control" name="edit_class_name" id="edit_class_name" value="<%=ThisClass(1)%>"></td>
                    </tr>
                    <tr>
                        <th>Gender:</th>
                        <td>
                            <select class="form-control" name="edit_gender" id="edit_gender">
                                <%Select Case ThisClass(2)%>
                                    <%Case "Both"%>
                                        <option value="Both" selected>Both</option>
                                        <option value="Male">Male</option>
                                        <option value="Female">Female</option>
                                    <%Case "Male"%>
                                        <option value="Both">Both</option>
                                        <option value="Male" selected>Male</option>
                                        <option value="Female">Female</option>
                                    <%Case "Female"%>
                                        <option value="Both">Both</option>
                                        <option value="Male">Male</option>
                                        <option value="Female" selected>Female</option>
                                <%End Select%>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <th>Details:</th>
                        <td><textarea class="form-control" name="edit_details" id="edit_details" rows="5"><%=ThisClass(3)%></textarea></td>
                    </tr>
                    <tr>
                        <td style="text-align: center;" colspan="2">
                            <input type="checkbox" name="delete_class" id="delete_class">&nbsp;Delete This Class (No Undo!)
                        </td>
                    </tr>
                    <tr>
                        <td style="text-align: center;" colspan="2">
                            <input type="hidden" name="submit_edit" id="submit_edit" value="submit_edit">
                            <input type="submit" class="form-control" name="submit2" id="submit2" value="Save Changes">
                        </td>
                    </tr>
                </table>
                </form>
            <%End If%>

            <br>

            <h4 class="h4">Assign Races To Classes:</h4>

            <form class="form" name="assign_races" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
            <table class="table table-striped"><tr><th>No.</th><th>Race</th><th>Class</th></tr>
                <%For i = 0 To UBound(Races, 2) - 1%>
                    <tr>
                        <td><%=i + 1%>)</td>
                        <td><%=Races(1, i)%></td>
                        <td>
                            <select class="form-control" name="race_classes_<%=Races(0, i)%>">
                                <option value="">--</option>
                                <%For j = 0 To UBound(MeetClasses, 2) - 1%>
                                    <%If CLng(Races(2, i)) = CLng(MeetClasses(0, j)) Then%>
                                        <option value="<%=MeetClasses(0, j)%>" selected><%=MeetClasses(1, j)%></option>
                                    <%Else%>
                                        <option value="<%=MeetClasses(0, j)%>"><%=MeetClasses(1, j)%></option>
                                    <%End If%>
                                <%Next%>
                            </select>
                        </td>
                    </tr>
                <%Next%>
                <tr>
                    <td style="text-align: center;" colspan="3">
                        <input type="hidden" name="submit_races" id="submit_races" value="submit_races">
                        <input type="submit" class="form-control" name="submit3" id="submit3" value="Submit Races">
                    </td>
                </tr>
            </table>
            </form>
        </div>
        <div class="col-xs-8">
            <h4 class="h4">Assign Teams To Classes:</h4>

            <form class="form" name="assign_teams" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
            <div class="form-group">
                <input type="hidden" name="submit_teams" id="submit_teams" value="submit_teams">
                <input type="submit" class="form-control" name="submit4" id="submit4" value="Save Changes">
            </div>
            <div class="col-xs-6">
                <h5 class="h5">Boys Teams</h5>
                <table class="table table-striped">
                    <tr><th>No.</th><th>Team</th><th>Class</th></tr>
                    <%k = 0%>
                    <%For i = 0 To UBound(MeetTeams, 2) - 1%>
                        <%If MeetTeams(2, i) = "M" Then%>
                            <tr>
                                <td><%=k + 1%>)</td>
                                <td>
                                    <a href="javascript:pop('/ccmeet_admin/manage_meet/edit_this_team.asp?team_id=<%=MeetTeams(0, i)%>',600,150)"><%=MeetTeams(1, i)%></a>
                                </td>
                                <td>
                                    <select class="form-control" name="team_classes_<%=MeetTeams(0, i)%>">
                                        <option value="">&nbsp;</option>
                                        <%For j = 0 To UBound(MeetClasses, 2) - 1%>
                                            <%If CLng(MeetTeams(3, i)) = CLng(MeetClasses(0, j)) Then%>
                                                <option value="<%=MeetClasses(0, j)%>" selected><%=MeetClasses(1, j)%></option>
                                            <%Else%>
                                                <option value="<%=MeetClasses(0, j)%>"><%=MeetClasses(1, j)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>

                            <%k = k + 1%>
                        <%End If%>
                    <%Next%>
                </table>
            </div>
            <div class="col-xs-6">
                <h5 class="h5">Girls Teams</h5>
                <table class="table table-striped">
                    <tr><th>No.</th><th>Team</th><th>Class</th></tr>
                    <%k = 0%>
                    <%For i = 0 To UBound(MeetTeams, 2) - 1%>
                        <%If MeetTeams(2, i) = "F" Then%>
                            <tr>
                                <td><%=k + 1%>)</td>
                                <td>
                                    <a href="javascript:pop('/ccmeet_admin/manage_meet/edit_this_team.asp?team_id=<%=MeetTeams(0, i)%>',600,150)"><%=MeetTeams(1, i)%></a>
                                </td>
                                <td>
                                    <select class="form-control" name="team_classes_<%=MeetTeams(0, i)%>">
                                        <option value="">&nbsp;</option>
                                        <%For j = 0 To UBound(MeetClasses, 2) - 1%>
                                            <%If CLng(MeetTeams(3, i)) = CLng(MeetClasses(0, j)) Then%>
                                                <option value="<%=MeetClasses(0, j)%>" selected><%=MeetClasses(1, j)%></option>
                                            <%Else%>
                                                <option value="<%=MeetClasses(0, j)%>"><%=MeetClasses(1, j)%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>

                            <%k = k + 1%>
                        <%End If%>
                    <%Next%>
                </table>
            </div>
            </form>
        </div>
    <%End If%>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
