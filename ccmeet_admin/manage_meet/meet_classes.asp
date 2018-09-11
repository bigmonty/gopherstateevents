<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lThisMeet, lThisClass, lPrevMeet, lPrevClassID, lNewClassID
Dim sSport, sClassName, sGender, sDetails, sDeleteClass, sLockClasses, sPrevClass
Dim Races(), MeetTeams(), MeetClasses(), ThisClass(3), OtherMeets()
Dim dMeetDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lThisMeet = Request.QueryString("meet_id")
lThisClass = Request.QueryString("class_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

sql = "SELECT MeetDate, Sport, LockClasses FROM Meets WHERE MeetsID = " & lThisMeet 
Set rs = conn.Execute(sql)
dMeetDate = rs(0).Value
sSport = rs(1).Value
sLockClasses = rs(2).Value
Set rs = Nothing

'get other meets
i = 0
ReDim OtherMeets(1, 0)
sql = "SELECT MeetsID, MeetName, MeetDate FROM Meets WHERE MeetsID <> " & lThisMeet & " AND Sport = '" & sSport
sql = sql & "' ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	OtherMeets(0,  i) = rs(0).Value
	OtherMeets(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
 	i = i + 1
	ReDim Preserve OtherMeets(1, i)
	rs.MoveNext
Loop
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

If Request.Form.Item("auto_assign") = "auto_assign" Then
    lPrevMeet = Request.Form.Item("prev_meet")
    For i = 0 To UBound(MeetTeams, 2) - 1
        lPrevClassID = "0"
        lNewClassID = "0"
        sPrevClass = vbNullString

        'see if this team was in the meet and was assigned a class in the previous meet
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lPrevMeet & " AND TeamsID = " & MeetTeams(0, i)
        rs.Open sql, conn, 1, 2
        If rs.RecordCount > 0 Then lPrevClassID = rs(0).Value
        rs.Close
        Set rs = Nothing

        If CLng(lPrevClassID) > 0 Then
            'if they were, get the class name
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT ClassName FROM MeetClasses WHERE MeetClassesID = " & lPrevClassID 
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then sPrevClass = rs(0).Value
            rs.Close
            Set rs = Nothing

            'if they were in a class, put them in the class with the same name in this meet
            If Not sPrevClass = vbNullString Then
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT MeetClassesID FROM MeetClasses WHERE ClassName = '" & sPrevClass & "' AND MeetsID = " 
                sql = sql & lThisMeet 
                rs.Open sql, conn, 1, 2
                If rs.RecordCount > 0 Then lNewClassID = rs(0).Value
                rs.Close
                Set rs = Nothing

                If CLng(lNewClassID) > 0 Then
                    'enter this class...only if they are not already entered into a class for this meet
                    Set rs = Server.CreateObject("ADODB.Recordset")
                    sql = "SELECT MeetClass FROM MeetTeams WHERE MeetsID = " & lThisMeet & " AND TeamsID = " & MeetTeams(0, i)
                    rs.Open sql, conn, 1, 2
                    If rs(0).Value & "" = "" Then
                        rs(0).Value = lNewClassID
                        rs.Update
                    End If
                    rs.Close
                    Set rs = Nothing
                End If
            End If
        End If
    Next
ElseIf Request.Form.Item("submit_lock") = "submit_lock" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT LockClasses FROM Meets WHERE MeetsID = " & lThisMeet 
    rs.Open sql, conn, 1, 2
    rs(0).Value = Request.Form.Item("lock_classes")
    rs.Update
    rs.Close
    Set rs = Nothing
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

If CStr(lThisClass) = vbNullString Then lThisClass = 0

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
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE  Admin CC/Nordic Meet Classes</title>

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
    <!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
	        <!--#include file = "manage_meet_nav.asp" -->
			
			<h4 class="h4">Meet Classes</h4>
				
			<p>Assigning Meet Classes to an event is a multi-step process:</p>
			<ol>
                <li>Create the classes</li>
                <li>Designate which race goes in which class (not all races need to be determined by class)</li>
                <li>Designate a class for each team.  Failure to designate a class for a team will prevent them from entering athletes in any race
                    that has a class designation.</li>
            </ol>

            <table style="padding: 5px;">
                <tr>
                    <td style="width: 250px;" valign ="top">
                        <h4 class="h4">Create Class:</h4>

                        <form name="add_class" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>" onsubmit="return chkFlds();">
                        <table>
                            <tr>
                                <th>Class Name:</th>
                                <td><input class="form-control" type="text" name="class_name" id="class_name"></td>
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
                                <td><input class="form-control" type="text" name="details" id="details"></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;" colspan="4">
                                    <input type="hidden" name="submit_class" id="submit_class" value="submit_class">
                                    <input class="form-control" type="submit" name="submit1" id="submit1" value="Submit Class">
                                </td>
                            </tr>
                        </table>
                        </form>

                        <br>

                        <h4 class="h4">Edit/Delete Classes:</h4>

                        <form name="select_class" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
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
                                    <input class="form-control" type="submit" name="submit5" id="submit5" value="Select Class">
                                </td>
                            </tr>
                        </table>
                        </form>

                        <br>
                        
                        <h4 class="h4">Lock Classes:</h4>

                        <p>Locking classes simply prevents coaches from making changes in their team's class.  It does not prevent GSE Admins or
                        Meet Managers from making changes.</p>

                        <form name="lock_meet_classes" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
                        <table>
                            <tr>
                                <th>Lock:</th>
                                <td>
                                    <select class="form-control" name="lock_classes" id="lock_classes">
                                        <%If sLockClasses = "y" Then%>
                                            <option value="y" selected>Yes</option>
                                            <option value="n">No</option>
                                        <%Else%>
                                            <option value="y">Yes</option>
                                            <option value="n" selected>No</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <input type="hidden" name="submit_lock" id="submit_lock" value="submit_lock">
                                    <input class="form-control" type="submit" name="submit5x" id="submit5x" value="Lock Classes">
                                </td>
                            </tr>
                        </table>
                        </form>

                        <%If Not CLng(lThisClass) = 0 Then%>
                            <br><br>
                            <form name="edit_class" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>&amp;class_id=<%=lThisClass%>">
                            <table>
                                <tr>
                                    <th>Class Name:</th>
                                    <td><input class="form-control" type="text" name="edit_class_name" id="edit_class_name" value="<%=ThisClass(1)%>"></td>
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
                                    <td><input class="form-control" type="text" name="edit_details" id="edit_details" value="<%=ThisClass(3)%>" size="30"></td>
                                </tr>
                                <tr>
                                    <td style="text-align: center;" colspan="2">
                                        <input type="checkbox" name="delete_class" id="delete_class">&nbsp;Delete This Class (No Undo!)
                                    </td>
                                </tr>
                                <tr>
                                    <td style="text-align: center;" colspan="2">
                                        <input type="hidden" name="submit_edit" id="submit_edit" value="submit_edit">
                                        <input class="form-control" type="submit" name="submit2" id="submit2" value="Save Changes">
                                    </td>
                                </tr>
                            </table>
                            </form>
                        <%End If%>
                    </td>
                    <td valign="top">
                        <h4 class="h4">Assign Races To Classes:</h4>

                        <form name="assign_races" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
                        <table><tr><th>No.</th><th style="text-align: left;">Race</th><th style="text-align: left;">Class</th></tr>
                            <%For i = 0 To UBound(Races, 2) - 1%>
                                <%If i mod 2 = 0 Then%>
                                    <tr>
                                        <td class="alt"><%=i + 1%>)</td>
                                        <td class="alt"><%=Races(1, i)%></td>
                                        <td class="alt">
                                            <select class="form-control" name="race_classes_<%=Races(0, i)%>">
                                                <option value="">No Class</option>
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
                                <%Else%>
                                    <tr>
                                        <td><%=i + 1%>)</td>
                                        <td><%=Races(1, i)%></td>
                                        <td>
                                            <select class="form-control" name="race_classes_<%=Races(0, i)%>">
                                                <option value="">No Class</option>
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
                                <%End If%>
                            <%Next%>
                            <tr>
                                <td style="text-align: center;" colspan="2">
                                    <input type="hidden" name="submit_races" id="submit_races" value="submit_races">
                                    <input class="form-control" type="submit" name="submit3" id="submit3" value="Submit Races">
                                </td>
                            </tr>
                        </table>
                        </form>
                    </td>
                    <td valign="top">
                        <h4 class="h4">Assign Teams To Classes:</h4>
                                    
                        <%If UBound(MeetClasses, 2) > 0 Then%>
                            <a href="javascript:pop('teams_by_class.asp?meet_id=<%=lThisMeet%>',800,750)">Teams By Class</a>
                        <%End If%>
            
                        <div class="bg-success">
                            <h5 class="h5">Auto Assign From Previous Meet</h5>
                            <form class="form-inline" name"auto_assign_classes" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
                            <label for="which_meet">Meet:</label>
                            <select class="form-control" name="prev_meet" id="prev_meet">
                                <option value="">&nbsp;</option>
                                <%For i = 0 To UBound(OtherMeets, 2) - 1%>
                                    <option value="<%=OtherMeets(0, i)%>"><%=OtherMeets(1, i)%></option>
                                <%Next%>
                            </select>
                            <input type="hidden" name="auto_assign" id="auto_assign" value="auto_assign">
                            <input class="form-control" type="submit" name="submit4a" id="submit4a" value="Go">
                            </form>
                        </div>
                        <hr>
                        <form name="assign_teams" method="post" action="meet_classes.asp?meet_id=<%=lThisMeet%>">
                        <table>
                            <tr>
                                <td style="text-align: center;" colspan="2">
                                    <input type="hidden" name="submit_teams" id="submit_teams" value="submit_teams">
                                    <input class="form-control" type="submit" name="submit4" id="submit4" value="Submit Teams">
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <h5 style="font-size: 1.1em;">Boys Teams</h5>
                                    <table style="font-size: 1.0em;">
                                        <tr><th>No.</th><th style="text-align: left;">Team</th><th style="text-align: left;">Class</th></tr>
                                        <%k = 0%>
                                        <%For i = 0 To UBound(MeetTeams, 2) - 1%>
                                            <%If MeetTeams(2, i) = "M" Then%>
                                                <%If k mod 2 = 0 Then%>
                                                    <tr>
                                                        <td class="alt"><%=k + 1%>)</td>
                                                        <td class="alt"><%=MeetTeams(1, i)%></td>
                                                        <td class="alt">
                                                            <select name="team_classes_<%=MeetTeams(0, i)%>">
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
                                                <%Else%>
                                                    <tr>
                                                        <td><%=k + 1%>)</td>
                                                        <td><%=MeetTeams(1, i)%></td>
                                                        <td>
                                                            <select name="team_classes_<%=MeetTeams(0, i)%>">
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
                                                <%End If%>

                                                <%k = k + 1%>
                                            <%End If%>
                                        <%Next%>
                                    </table>
                                </td>
                                <td valign="top">
                                    <h5 style="font-size: 1.1em;">Girls Teams</h5>
                                    <table style="font-size: 1.0em;">
                                        <tr><th>No.</th><th style="text-align: left;">Team</th><th style="text-align: left;">Class</th></tr>
                                        <%k = 0%>
                                        <%For i = 0 To UBound(MeetTeams, 2) - 1%>
                                            <%If MeetTeams(2, i) = "F" Then%>
                                                <%If k mod 2 = 0 Then%>
                                                    <tr>
                                                        <td class="alt"><%=k + 1%>)</td>
                                                        <td class="alt"><%=MeetTeams(1, i)%></td>
                                                        <td class="alt">
                                                            <select name="team_classes_<%=MeetTeams(0, i)%>">
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
                                                <%Else%>
                                                    <tr>
                                                        <td><%=k + 1%>)</td>
                                                        <td><%=MeetTeams(1, i)%></td>
                                                        <td>
                                                            <select name="team_classes_<%=MeetTeams(0, i)%>">
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
                                                <%End If%>

                                                <%k = k + 1%>
                                            <%End If%>
                                        <%Next%>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        </form>
                    </td>
                </tr>
            </table>
		</div>
	</div>
<!--#include file = "../../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
