<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim fs
Dim iLeapYrs, iAgeDays, iAge
Dim sBatchFileName, sErrMsg
Dim lPartID, lEventID
Dim Filepath
Dim file    
Dim TextStream		
Dim Line
Dim sSplit,  sField
Dim field1, field2, field3, field4, field5, field6, field7, field8, field9, field10, field11, field12, field13, field14, field15, field16, field17, field18
Dim Races()
Dim bPartExists
Dim dEVentDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

'On Error Resume Next

lEventID = Request.QueryString("event_id")

If CStr(lEventID) = vbNullString Then 
	lEventID = Session("event_id") 
	Session("event_id") = vbNullString
Else
	Session("event_id") = lEventID
End If

Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

sBatchFileName = Request.QueryString("which_file")

Response.Buffer = true		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
dEventDate = rs(0).Value
rs.Close
Set rs = Nothing

If Not CStr(sBatchFileName) = vbNullString Then
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
			field4 = sSplit(3)							'age
			field5 = sSplit(4)							'dob
			field6 = sSplit(5)							'phone
			If Not sSplit(6) & "" = "" Then field7 = Replace(sSplit(6), "'", "''")		'city
			If Not sSplit(7) & "" = "" Then field8 = Replace(sSplit(7), "'", "''")		'state	
			field9 = sSplit(8)						'email
			field10 = sSplit(9)						'size
			field11 = sSplit(10)					'bib
			field12 = sSplit(11)					'race id

			field1 = Replace(field1, chr(34), "")
			field2 = Replace(field2, chr(34), "")
			field3 = Replace(field3, chr(34), "")
			If Not field4 & "" = "" Then field4 = Replace(field4, chr(34), "")
			If Not field5 & "" = "" Then field5 = Replace(field5, chr(34), "")
			If Not field6 & "" = "" Then field6 = Replace(field6, chr(34), "")
			If Not field7 & "" = "" Then field7 = Replace(field7, chr(34), "")
			If Not field8 & "" = "" Then field8 = Replace(field8, chr(34), "")
			If Not field9 & "" = "" Then field9 = Replace(field9, chr(34), "")
			If Not field10 & "" = "" Then field10 = Replace(field10, chr(34), "")
			If Not field11 & "" = "" Then field11 = Replace(field11, chr(34), "")
            field12 = Replace(field12, chr(34), "")

            If CStr(field4) = vbNullString Then 
                'first get leap years
                iLeapYrs = 0
                For i = Year(CDate(dEventDate)) To Year(CDate(field5)) Step -1
                    If i / 4 = i \ 4 Then iLeapYrs = iLeapYrs + 1
                Next
        
                iAgeDAys = DateDiff("d", CDate(field5), CDate(dEventDate))
                iAgeDAys = iAgeDAys - iLeapYrs
                iAge = iAgeDAys \ 365

                field4 = iAge
            End If

            'check for a data match
            lPartID = 0
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT ParticipantID FROM Participant WHERE FirstName = '" & field1 & "' AND LastName = '" & field2
            sql = sql & "' AND City = '" & field7 & "' AND Email = '" & field9 & "' AND Gender = '" & field3 & "'"
            rs.Open sql, conn, 1, 2
            If rs.RecordCount > 0 Then lPartID = rs(0).Value
            rs.Close
            Set rs = Nothing
 
            bPartExists = False 'this flag is jsut so we don't put the same person in this race twice
                    
            'they are in the db...see if they are in this race
            If CLng(lPartID) = 0 Then
	            'insert into the partdata table
	            sql = "INSERT INTO Participant (FirstName, LastName, Gender, DOB, Phone, City, St, Email)"
	            sql = sql & " VALUES ('" & field1 & "', '"  & field2 & "', '" & field3 & "', '" & field5 & "', '" & field6 & "', '"
	            sql = sql & field7 & "', '"  & field8 & "', '" & field9 & "')"			
	            Set rs=conn.Execute(sql)
	            Set rs=Nothing
			
	            'get partid
	            sql = "SELECT ParticipantID FROM Participant WHERE FirstName='" & field1 & "' AND LastName='" & field2 & "' AND Gender = '" & field3 
	            sql = sql & "' AND City = '" & field7 & "' AND St = '" & field8 & "' AND Email = '" & field9 & "' ORDER BY ParticipantID DESC"
	            Set rs = conn.Execute(sql)
	            lPartID = rs(0).Value
	            Set rs=Nothing
            Else
                'check to see if in this race
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT ParticipantID FROM PartReg WHERE ParticipantID = " & lPartID & " AND RaceID = " & field12
                rs.Open sql, conn, 1, 2
                If rs.RecordCount > 0 Then bPartExists = True
                rs.Close
                Set rs = Nothing
            End If
 
            If bPartExists = False Then
			    'insert into part reg table
			    sql = "INSERT INTO PartReg (ParticipantID, RaceID, DateReg, WhereReg) VALUES (" & lPartID & ", " & field12 & ", '" & Date & "', 'na')"
			    Set rs=conn.Execute(sql)
			    Set rs=Nothing
		
			    'insert into part race table
			    sql = "INSERT INTO PartRace (ParticipantID, Age, RaceID, Bib, AgeGrp) VALUES (" & lPartID & ", "  & field4 & ", " & field12 & ", '" 
                sql = sql & field11 & "', '" & GetAgeGrp(field3, field4, field12) & "')"
			    Set rs=conn.Execute(sql)
			    Set rs=Nothing
            End If
		Loop
	    Set TextStream = nothing
		
		If err.number <> 0 Then
		  sErrMsg = "There was an error in the formatting of at least one field.  At least one participant was not added to this event."
		End if	
	Else
	    Response.Write sBatchFileName & " can not be found."
	End If
	
	If fs.FileExists("c:\Inetpub\h51web\gopherstateevents\admin\participants\batch_upload\" & sBatchFileName) = True Then
		fs.DeleteFile("c:\Inetpub\h51web\gopherstateevents\admin\participants\batch_upload\" & sBatchFileName)
	End If
	Set fs = nothing
End If

'get race ids for this event
i = 0
ReDim Races(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

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
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Batch Upload</title>
<!--#include file = "../../../includes/js.asp" -->
<style type="text/css">
	li{
		margin:5px;
	}
</style>
</head>

<body>
<img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
<div class="container">
	<%If Not sErrMsg = vbNullString Then%>
		<p><%=sErrMsg%></p>
	<%End If%>
	
	<h4 class="h4">Participant Batch Upload Utility</h4>
	
	<h5 class="h5">PLEASE READ CAREFULLY!</h5>
	
	<p>In order for your data to upload seamlessly the following guidelines must be followed EXACTLY!</p>
	
	<ul>
		<li>The file MUST BE a tab-delimited text file (*.txt)</li>
		<li>The file CAN NOT have a header row.</li>
		<li>The file CAN NOT have any trailing spaces or rows after the final line.</li>
		<li>The file MUST HAVE ONLY the following fields IN THIS ORDER.  Required fields are so noted and optional fields MUST exist at 
			least as an empty field.</li>
		<li style="list-style:none;">
			<ul>
				<li style="margin:0;">First Name (reqd)</li>
				<li style="margin:0;">Last Name (reqd)</li>
				<li style="margin:0;">Gender ("M" or "F") (reqd)</li>
                <li style="margin:0;">Age (DOB or Age reqd)</li>
				<li style="margin:0;">DOB (DOB or Age reqd)</li>
				<li style="margin:0;">Phone</li>
				<li style="margin:0;">City</li>
				<li style="margin:0;">State</li>
				<li style="margin:0;">Email</li>
				<li style="margin:0;">Shirt Size</li>
				<li style="margin:0;">Bib</li>
				<li style="margin:0;">Race ID # (reqd-see below)</li>
			</ul>
		</li>
	</ul>

	<h5 class="h5">Race ID #s</h5>
	<ul>
		<%For i = 0 To UBound(Races, 2)- 1%>
			<li><%=Races(1, i)%> - <%=Races(0, i)%></li>
		<%Next%>
	</ul>
	
	<fieldset style="margin:5px;text-align:center;">
		<legend>&nbsp;Upload File:&nbsp;</legend>
		<form name="upload_batch" method="Post"  enctype="multipart/form-data" action="receive_part.asp?event_id=<%=lEventID%>">
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