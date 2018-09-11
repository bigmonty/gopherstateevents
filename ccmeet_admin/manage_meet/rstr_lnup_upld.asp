<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim fs
Dim sBatchFileName, sErrMsg, sGradeYear, sWhichFile, sMeetName
Dim lMyID, lTeamID, lMeetID
Dim Filepath
Dim file    
Dim TextStream		
Dim Line
Dim sSplit,  sField
Dim field1, field2, field3, field4, field5, field6, field7
Dim Races(), MeetTeams()
Dim dMeetDate
Dim bPartFound

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lMeetID = Request.QueryString("meet_id")

If CStr(lMeetID) = vbNullString Then 
	lMeetID = Session("meet_id") 
	Session("meet_id") = vbNullString
Else
	Session("meet_id") = lMeetID
End If

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

sWhichFile = Request.QueryString("which_file")
sBatchFileName = "/ccmeet_admin/manage_meet/batch_upload/" & sWhichFile

Response.Buffer = true		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get meet name
sql = "SELECT MeetName, MeetDate FROM Meets WHERE MeetsID = " & lMeetID
Set rs = conn.Execute(sql)
sMeetName = Replace(rs(0).Value, "''", "'")
dMeetDate = rs(1).Value
Set rs = Nothing

i = 0
ReDim MeetTeams(2, 0)
sql = "SELECT t.TeamsID, t.TeamName, t.Gender FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
sql = sql & "WHERE mt.MeetsID = " & lMeetID & " ORDER BY t.TeamName, t.Gender"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	MeetTeams(0, i) = rs(0).Value
	MeetTeams(1, i) = Replace(rs(1).Value, "''", "'")
    MeetTeams(2, i) = rs(2).Value
	i = i + 1
	ReDim Preserve MeetTeams(2, i)
	rs.MoveNext
Loop
Set rs = Nothing

'get races in this meet
i = 0
ReDim Races(1, 0)
sql = "SELECT RacesID, RaceDesc FROM Races WHERE MeetsID = " & lMeetID & " ORDER BY ViewOrder, RaceTime"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Not sWhichFile = vbNullString Then
	Set fs = Server.CreateObject("Scripting.FileSystemObject")

	Filepath = Server.MapPath(sBatchFileName)

	If fs.FileExists(Filepath) Then
	    Set file = fs.GetFile(Filepath)
	    Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)
		
	    Do While Not TextStream.AtEndOfStream
	        Line = TextStream.readline
			sSplit =  Split(Line, vbTab)	

			field1 = Trim(sSplit(0))                            'team id
			field2 = Replace(Trim(sSplit(1)), "'", "''")		'first name
			field3 = Replace(Trim(sSplit(2)), "'", "''")		'last name	
            field4 = Trim(sSplit(3))							'gender
			field5 = Trim(sSplit(4))							'grade	
            field6 = Trim(sSplit(5))							'race id
			field7 = Trim(sSplit(6))							'bib	

			field1 = Replace(field1, chr(34), "")
			field2 = Replace(field2, chr(34), "")
			field3 = Replace(field3, chr(34), "")
			field4 = Replace(field4, chr(34), "")
            field5 = Replace(field5, chr(34), "")
			field6 = Replace(field6, chr(34), "")
            field7 = Replace(field7, chr(34), "")

            field4 = Left(field4, 1)                            'just take the first letter of gender

            'see if they exist in the db
            bPartFound = False
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT RosterID FROM Roster WHERE FirstName = '" & field2 & "' AND LastName = '" & field3 & "' AND Gender = '" & field4
            sql = sql & "' AND TeamsID = " & field1
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then 
                lMyID = rs(0).Value
                bPartFound = True
            End If
            rs.Close
            Set rs = Nothing
        
            If bPartFound = False Then
                'insert team member
                sql = "INSERT INTO Roster (TeamsID, FirstName, LastName, Gender) VALUES (" & field1 & ", '" & field2 & "', '" & field3 & "', '" 
		        sql = sql & field4 & "')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
        
                'get roster id
                sql = "SELECT RosterID FROM Roster WHERE TeamsID = " & field1 & " AND FirstName = '" & field2 & "' AND LastName = '"
                sql = sql & field3 & "' AND Gender = '" & field4 & "' ORDER BY RosterID DESC"
                Set rs = conn.Execute(sql)
                lMyID = rs(0).Value
                Set rs = Nothing
 
		        'get year for roster grades
		        If Month(Date) <=7 Then
			        sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
		        Else
			        sGradeYear = Right(CStr(Year(Date)), 2)	
		        End If
       
                'insert grade
                sql = "INSERT INTO Grades (RosterID, Grade" & sGradeYear & ") VALUES (" & lMyID & ", " & field5 & ")"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            End If

            'now add to race and enter bib
            If CLng(field6) > 0 Then    'if they have a race id listed
                bPartFound = False
		        Set rs = Server.CreateObject("ADODB.Recordset")
		        sql = "SELECT RacesID, Bib FROM IndRslts WHERE RosterID = " & lMyID & " AND MeetsID = " & lMeetID
		        rs.Open sql, conn, 1, 2
		        If rs.recordcount > 0 Then bPartFound = True
		        rs.Close
		        Set rs = Nothing
	
                If bPartFound = False Then
				    sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID, Bib) VALUES (" & lMeetID & ", " & field6
				    sql = sql & ", " & lMyID & ", " & field7 & ")"
				    Set rs = conn.Execute(sql)
				    Set rs = Nothing
                End If
            End If
        Loop
	Else
	    Response.Write sWhichFile & " can not be found."
	End If
	
'	If fs.FileExists("C:\inetpub\h51web\gopherstateevents\ccmeet_admin\manage_team\roster_upload\batch_upload\" & sWhichFile) = True Then
'		fs.DeleteFile("C:\inetpub\h51web\gopherstateevents\ccmeet_admin\manage_team\roster_upload\batch_upload\" & sWhichFile)
'	End If
	Set fs = nothing
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" 
"http://www.w3.org/TR/html4/strict.dtd">

<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Meet Roster & Line-Up Upload Utility</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
	<h3 class="h3">GSE Meet Roster & Line-Up Upload Utility</h3>
    <h4 class="h4"><%=sMeetName%> on <%=dMeetDate%></h4>

    <br>

    <div class="row">
        <div class="col-sm-3 bg-warning">
            <h4 class="h4">&nbsp;Race IDs:&nbsp;</h4>
	        <ul class="list-group">
                <%For i = 0 To UBound(Races, 2) - 1%>
                    <li class="list-group-item"><%=Races(1, i)%>: <%=Races(0, i)%></li>
                <%Next%>
            </ul>
        </div>
        <div class="col-sm-5 bg-success">
            <h4 class="h4">&nbsp;Team IDs:&nbsp;</h4>
	        <ul class="list-group">
                <%For i = 0 To UBound(MeetTeams, 2) - 1%>
                    <li class="list-group-item"><%=MeetTeams(1, i)%> (<%=MeetTeams(2, i)%>): <%=MeetTeams(0, i)%></li>
                <%Next%>

            </ul>
        </div>
        <div class="col-sm-4 bg-info">
            <h4 class="h4">&nbsp;Field Definitions:&nbsp;</h4>
	        <ul class="list-group">
                <li class="list-group-item">Team ID</li>
		        <li class="list-group-item">First Name</li>
		        <li class="list-group-item">Last Name</li>
		        <li class="list-group-item">Gender ("M" or "F")</li>
		        <li class="list-group-item">Numeric Grade</li>
                <li class="list-group-item">Race ID ("0" for none)</li>
                <li class="list-group-item">Bib ("0" for none)</li>
            </ul>
        </div>
    </div>
	
	<h4 class="h4">&nbsp;Select File:&nbsp;</h4>
	<form class="form-inline" name="upload_batch" method="Post"  enctype="multipart/form-data" action="receive_rstr_lnup.asp">
	<div class="form-group">
        <input class="form-control" type="file" name="file1" id="file1">
	    <input type="hidden" name="submit_batch" id="submit_batch" value="submit_batch">
	    <input class="form-control"  type="submit" id="submit1" name="submit1" value="Upload File">
    </div>
	</form>
    <br>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>