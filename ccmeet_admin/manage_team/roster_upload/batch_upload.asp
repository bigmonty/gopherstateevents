<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim fs
Dim sBatchFileName, sErrMsg, sGradeYear, sWhichFile
Dim lMyID, lTeamID
Dim Filepath
Dim file    
Dim TextStream		
Dim Line
Dim sSplit,  sField
Dim field1, field2, field3, field4
Dim bPartFound

If Not (Session("role") = "admin" Or Session("role") = "coach") Then Response.Redirect "/default.asp?sign_out=y"

lTeamID = Request.QueryString("team_id")

If CStr(lTeamID) = vbNullString Then 
	lTeamID = Session("team_id") 
	Session("team_id") = vbNullString
Else
	Session("team_id") = lTeamID
End If

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

sWhichFile = Request.QueryString("which_file")
sBatchFileName = "/ccmeet_admin/manage_team/roster_upload/batch_upload/" & sWhichFile

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
            field3 = Left(Trim(sSplit(2)), 1)							'gender
			field4 = Trim(sSplit(3))							'grade	

			field1 = Replace(field1, chr(34), "")
			field2 = Replace(field2, chr(34), "")
			field3 = Replace(field3, chr(34), "")
            field4 = Replace(field4, chr(34), "")

            'see if they exist in the db
            bPartFound = False
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT RosterID FROM Roster WHERE FirstName = '" & field1 & "' AND LastName = '" & field2 & "' AND Gender = '" & field3
            sql = sql & "' AND TeamsID = " & lTeamID
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then bPartFound = True
            rs.Close
            Set rs = Nothing
        
            If bPartFound = False Then
                'insert team member
                sql = "INSERT INTO Roster (TeamsID, FirstName, LastName, Gender) VALUES (" & lTeamID & ", '" & field1 & "', '" & field2 & "', '" 
		        sql = sql & field3 & "')"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
        
                'get roster id
                sql = "SELECT RosterID FROM Roster WHERE TeamsID = " & lTeamID & " AND FirstName = '" & field1 & "' AND LastName = '"
                sql = sql & field2 & "' AND Gender = '" & field3 & "' ORDER BY RosterID DESC"
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
                sql = "INSERT INTO Grades (RosterID, Grade" & sGradeYear & ") VALUES (" & lMyID & ", " & field4 & ")"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
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
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Roster Upload</title>




<style type="text/css">
	li{
		margin:5px;
	}
</style>
</head>

<body style="background:none;">
<div style="margin:5px;text-align:left;">
	<h4 style="text-align:center;">GSE Roster Batch Upload Utility</h4>
	
	<h5 style="text-align:center;">PLEASE READ CAREFULLY!</h5>
	
	<p style="text-align:left;">In order for your data to upload seamlessly the following guidelines must be followed EXACTLY!  
        <span style="color: red;">We have just migrated to a new server.  If you experience any
    difficulty plase email your roster (formatted as indicated below) to 
    <a href="mailto:bob.schneider@gopherstateevents.com">bob.schneider@gopherstateevents.com</a> and we will upload it for you</span></p>
	
	<p style="text-align:left;">NOTE:  If a participant exists on your team with the exact same name they will not be uploaded.  If they are archived
    you can determin that gy going to the 'Archived Roster' page.  If they are, in fact, different people simply make a small change in the name 
    (ie: change Bob to Robert for instance) and enter them manually.</p>

	<%If Not sErrMsg = vbNullString Then%>
		<p><%=sErrMsg%></p>
	<%End If%>
	
	
	<ul style="padding:5px 5px 5px 20px;">
		<li>The file MUST BE a tab-delimited text file (*.txt)</li>
		<li>The file CAN NOT have a header row.</li>
		<li>The file CAN NOT have any trailing spaces or rows after the final line.</li>
		<li>The file MUST HAVE ONLY the following fields IN THIS ORDER.</li>
		<li style="list-style:none;">
			<ul>
				<li style="margin:0;">First Name</li>
				<li style="margin:0;">Last Name</li>
				<li style="margin:0;">Gender ("M" or "F")</li>
				<li style="margin:0;">Grade (numeric, not "Sophomore" for instance)</li>
			</ul>
		</li>
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
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>