<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim fs
Dim sBatchFileName
Dim EmailList()
Dim Filepath
Dim File
Dim TextStream		
Dim Line
Dim sSplit,  sField
Dim field1
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

sBatchFileName = "dontsend.txt"

If Not CStr(sBatchFileName) = vbNullString Then
	Set fs = Server.CreateObject("Scripting.FileSystemObject")

	Filepath = Server.MapPath(sBatchFileName)

	If fs.FileExists(Filepath) Then
	    Set file = fs.GetFile(Filepath)
	    Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)
		
	    Do While Not TextStream.AtEndOfStream
	        Line = TextStream.readline
			sSplit =  Split(Line, vbTab)	

			field1 = Trim(sSplit(0))		'first name
   
            'check for a data match
'            sql = "INSERT INTO DontSend(Email, WhenEntered) VALUES ('" & field1 & "', '" & Now() & "')"
'            Set rs = conn.Execute(sql)
'            Set rs = Nothing
		Loop
	    Set TextStream = nothing
	Else
	    Response.Write sBatchFileName & " can not be found."
	End If
End If

%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE&copy; Admin Data Modify</title>

<!--#include file = "../includes/js.asp" -->
</head>

<<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->
	<div id="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE&copy; Modify Data</h4>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
