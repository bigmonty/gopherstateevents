<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim fs
Dim sBatchFileName, sGradeYear, sWhichFile
Dim lMyID, lMeetID
Dim Filepath
Dim file    
Dim TextStream		
Dim Line
Dim sSplit,  sField
Dim field1, field2, field3, field4, field5, field6, field7, field8, field9
Dim bFound

If Not Session("role") = "admin" Then Response.Redirect "/index.html"

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

lMeetID = Request.QueryString("meet_id")
If Session("meet_id") = vbNullString Then Session("meet_id") = lMeetID
sWhichFile = Request.QueryString("which_file")
sBatchFileName = "/ccmeet_admin/manage_meet/part_upload/batch_upload/" & sWhichFile

Response.Buffer = true		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Not sWhichFile = vbNullString Then
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	Filepath = Server.MapPath(sBatchFileName)

	If fs.FileExists(Filepath) Then
	    Set file = fs.GetFile(Filepath)
	    Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)
		
	    Do While Not TextStream.AtEndOfStream
	        Line = TextStream.readline
			sSplit =  Split(Line, vbTab)	

			field1 = Replace(Trim(sSplit(0)), "'", "''")		'first name
			field2 = Replace(Trim(sSplit(1)), "'", "''")		'last name	
            field3 = Trim(sSplit(2))							'gender
			field4 = Trim(sSplit(3))							'grade	
			field5 = Replace(Trim(sSplit(4)), "'", "''")		'teamid
			field6 = Replace(Trim(sSplit(5)), "'", "''")		'race	
            field7 = Trim(sSplit(6))							'bib
			field8 = Trim(sSplit(7))							'start	
			field9 = Trim(sSplit(8))							'gate	

			field1 = Replace(field1, chr(34), "")
			field2 = Replace(field2, chr(34), "")
			field3 = Replace(field3, chr(34), "")
            field4 = Replace(field4, chr(34), "")
			field5 = Replace(field5, chr(34), "")
			field6 = Replace(field6, chr(34), "")
			field7 = Replace(field7, chr(34), "")
            field8 = Replace(field8, chr(34), "")
			field9 = Replace(field9, chr(34), "")

            'see if they exist in the team's roster
            lMyID = 0
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT RosterID FROM Roster WHERE FirstName = '" & field1 & "' AND LastName = '" & field2 
			sql = sql & "' AND Gender = '" & field3 & "' AND TeamsID = " & field5
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then lMyID = rs(0).Value
            rs.Close
            Set rs = Nothing
        
            If CLng(lMyID) = 0 Then
                'insert team member
                sql = "INSERT INTO Roster (TeamsID, FirstName, LastName, Gender) VALUES (" & field5 & ", '" & field1 & "', '" & field2 & "', '" 
		        sql = sql & field3 & "')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
        
                'get roster id
                sql = "SELECT RosterID FROM Roster WHERE TeamsID = " & field5 & " AND FirstName = '" & field1 
				sql = sql & "' AND LastName = '" & field2 & "' AND Gender = '" & field3 & "' ORDER BY RosterID DESC"
                Set rs = conn.Execute(sql)
                lMyID = rs(0).Value
                Set rs = Nothing
 
		        'get year for roster grades
		        If Month(Date) <=5 Then
			        sGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
		        Else
			        sGradeYear = Right(CStr(Year(Date)), 2)	
		        End If
       
                'insert grade
                sql = "INSERT INTO Grades (RosterID, Grade" & sGradeYear & ") VALUES (" & lMyID & ", " & field4 & ")"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            End If

			'now enter into indrslts
			'first see if they are entered in the race in question
			If CInt(field6) > 0 Then 	'only enter them if a race is designated
				bFound = False
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql = "SELECT IndRsltsID FROM IndRslts WHERE RacesID = " & field6 & " AND RosterID = " & lMyID
				rs.Open sql, conn, 1, 2
				If rs.RecordCount > 0 Then bFound = True
				rs.Close
				Set rs = Nothing

				'if not entered in the race then enter them
				If bFound = False Then
					sql = "INSERT INTO IndRslts (MeetsID, RacesID, RosterID, Bib, IndDelay, Gate) VALUES ("
					sql = sql & Session("meet_id") & ", " & field6 & ", " & lMyID & ", " & field7 & ", " & field8 & ", " & field9 & ")"
					Set rs = conn.Execute(sql)
					Set rs = Nothing
				End If
			End If
        Loop
	Else
	    Response.Write sWhichFile & " can not be found."
	End If
	
'	If fs.FileExists("C:\inetpub\h51web\gopherstateevents\ccmeet_admin\manage_meet\part_upload\batch_upload\" & sWhichFile) = True Then
'		fs.DeleteFile("C:\inetpub\h51web\gopherstateevents\ccmeet_admin\manage_meet\part_upload\batch_upload\" & sWhichFile)
'	End If
	Set fs = nothing
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" 
"http://www.w3.org/TR/html4/strict.dtd">

<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Roster Upload</title>

<style type="text/css">
	li{
		margin:5px;
	}
</style>
</head>

<body">
<div class="container">
	<h4 class="h4">GSE Meet Participant Batch Upload Utility</h4>
	
	<p>
		This utility will upload participants to team rosters (if they don't already exist) and enter them in a race.
		Where appropriate it will also assign bibs, ind delay, and gates.  Please ensure that all fields listed below
		have a value and are in the correct order!
	</p>
	<ul class="list-group">
		<li class="list-group-item">First Name</li>
		<li class="list-group-item">Last Name</li>
		<li class="list-group-item">Gender</li>
		<li class="list-group-item">Grade</li>
		<li class="list-group-item">Team</li>
		<li class="list-group-item">Race</li>
		<li class="list-group-item">Bib (0 if na)</li>
		<li class="list-group-item">Start (seconds;0 if na)</li>
		<li class="list-group-item">Gate (1 if na)</li>
	</ul>
	
	<fieldset style="margin:5px;text-align:center;">
		<legend>&nbsp;Upload File:&nbsp;</legend>
		<form name="upload_batch" method="Post"  enctype="multipart/form-data" action="receive_part.asp">
		<input type="file" name="file1" id="file1" size="30">
		<br>
		<input type="hidden" name="submit_batch" id="submit_batch" value="submit_batch">
		<input type="submit" id="submit1" name="submit1" value="Upload File" style="color:#009933;">
		</form>
	</fieldset>
</div>
<!--#include file = "../../../includes/js.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>