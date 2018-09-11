<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim i
Dim fs
Dim iLeapYrs, iAgeDays, iAge
Dim sBatchFileName, sErrMsg, sEventName
Dim lEventID
Dim Filepath
Dim file    
Dim TextStream		
Dim Line
Dim sSplit,  sField
Dim field1, field2, field3, field4, field5, field6, field7, field8
Dim Races()
Dim dEventDate

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
sql = "SELECT EventDate, EventName FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
dEventDate = rs(0).Value
sEventName = Replace(rs(1).Value, "''", "'")
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
			field4 = sSplit(3)							        'age
			field5 = sSplit(4)							        'dob
			field6 = sSplit(5)						            'email
			field7 = sSplit(6)					                'bib
			field8 = sSplit(7)					                'race id

            If Len(field6) > 0 Then                             'ensure email is not empty
                If CStr(field7) = vbNullString Then field7 = "0"

                If CStr(field4) = vbNullString Then                 'get age from DOB if no age
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
            End If
    	Loop
	    Set TextStream = nothing

        'inser into data table
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
<title>GSE Pre-Race Email From File Utility</title>
<!--#include file = "../../../includes/js.asp" -->
</head>

<body>
<img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive" style="margin-top: 15px;">
<div class="container">
	<%If Not sErrMsg = vbNullString Then%>
		<p><%=sErrMsg%></p>
	<%End If%>
	
	<h4 class="h4">Pre-Race Email From File Utility: <%=sEventName%> on <%=dEventDate%></h4>
	
    <div class="row">
        <div class="col-xs-6 bg-info">
	        <h5 class="h5">Field Order</h5>
	        <ul>
		        <li style="margin:0;">First Name (reqd)</li>
		        <li style="margin:0;">Last Name (reqd)</li>
		        <li style="margin:0;">Gender ("M" or "F") (reqd)</li>
                <li style="margin:0;">Age (DOB or Age reqd)</li>
		        <li style="margin:0;">DOB (DOB or Age reqd)</li>
		        <li style="margin:0;">Email (reqd)</li>
		        <li style="margin:0;">Bib</li>
		        <li style="margin:0;">Race ID # (reqd)</li>
	        </ul>
        </div>
        <div class="col-xs-6 bg-warning">
	        <h5 class="h5">Race ID #s</h5>
	        <ul>
		        <%For i = 0 To UBound(Races, 2)- 1%>
			        <li><%=Races(1, i)%> - <%=Races(0, i)%></li>
		        <%Next%>
	        </ul>
        </div>
    </div>
	
    <div class="row">
	    <div class="bg-success" style="padding:5px;">
		    <h4 class="h4">Upload File:</h4>
		    <form class="form" name="upload_batch" method="Post"  enctype="multipart/form-data" action="receive_file.asp?event_id=<%=lEventID%>">
            <h5 class="h5">Supplemental Message:</h5> 
            <textarea class="form-control" name="supp_msg_ind" id="supp_msg_ind" rows="3"></textarea>
		    <br>
		    <input class="form-control" type="file" name="file1" id="file1">
		    <input class="form-control" type="hidden" name="submit_batch" id="submit_batch" value="submit_batch">
		    <input class="form-control" type="submit" id="submit1" name="submit1" value="Upload File">
		    </form>
	    </div>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>