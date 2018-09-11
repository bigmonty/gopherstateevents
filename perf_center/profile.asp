<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, conn2, rs, sql, rs2, sql2
Dim i
Dim sFirstName, sLastName, sAddress, sCity, sState, sPostal, sEmail, sPhone, sComments, sMobile, sEmailErr, sDOBErr, sScreenName, sMyPix
Dim sUserName, sPassword, sConfirmPassword, sErrMsg
Dim iMonth, iDay, iYear
Dim dDOB

If CStr(Session("my_hist_id")) = vbNullString Then Response.Redirect "my_hist_login.asp"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_info") = "submit_info" Then
	If Not Request.Form.Item("first_name") & "" = "" Then sFirstName = Replace(Request.Form.Item("first_name"), "''", "'")
	If Not Request.Form.Item("last_name") & "" = "" Then sLastName =  Replace(Request.Form.Item("last_name"), "''", "'")
	If Not Request.Form.Item("address") & "" = "" Then sAddress =  Replace(Request.Form.Item("address"), "''", "'")
	If Not Request.Form.Item("city") & "" = "" Then sCity =  Replace(Request.Form.Item("city"), "''", "'")
	If Not Request.Form.Item("state") & "" = "" Then sState = Replace(Request.Form.Item("state"), "''", "'")
	If Not Request.Form.Item("postal") & "" = "" Then sPostal = Replace(Request.Form.Item("postal"), "''", "'")
	If Not Request.Form.Item("email") & "" = "" Then sEmail = Replace(Request.Form.Item("email"), "''", "'")
	If Not Request.Form.Item("phone") & "" = "" Then sPhone = Replace(Request.Form.Item("phone"), "''", "'")
	If Not Request.Form.Item("comments") & "" = "" Then sComments =  Replace(Request.Form.Item("comments"), "''", "'")
	If Not Request.Form.Item("mobile") & "" = "" Then sMobile = Replace(Request.Form.Item("mobile"), "''", "'")
	If Not Request.Form.Item("screen_name") & "" = "" Then sScreenName = Replace(Request.Form.Item("screen_name"), "''", "'")
    dDOB = Request.Form.Item("dob")

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, ScreenName, BirthDate, Email FROM PartData WHERE PartID = " & Session("etraxc_id")
	rs.Open sql, conn2, 1, 2
	
	If sFirstName & "" = "" Then
		rs(0).Value = rs(0).OriginalValue
	Else
		rs(0).Value = sFirstName
	End if
	
	If sLastName & "" = "" Then
		rs(1).Value = rs(1).OriginalValue
	Else
		rs(1).Value = sLastName
	End if

	rs(2).Value = sScreenName
	
	If sEmail & "" = "" Then
		s(4).Value = Null
	Else
        If ValidEmail(sEmail) = False Then sEmailErr = "Your email address does not appear to be in a valid format.  Please re-enter.  Some work was not done."

        If sEmailErr = vbNullString Then 
		    rs(4).Value = sEmail
        Else
            rs(4).Value = rs(4).OriginalValue
        End If
	End if

    rs.Update
    rs.Close
    Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT Address, City, St, Zip, Phone, CellPhone, Comments FROM PartProfile WHERE PartID = " & Session("etraxc_id")
	rs.Open sql, conn2, 1, 2
	rs(0).Value = sAddress
	rs(1).Value = sCity
	rs(2).Value = sState
	rs(3).Value = sPostal
	rs(4).Value = sPhone
	rs(5).Value = sMobile
	rs(6).Value = sComments
	rs.Update
	rs.Close
	Set rs = Nothing

    'update user name
	sUserName = Request.Form.Item("user_name")

    If Not sUserName & "" = "" Then
        If ValidUserName(sUserName) = False Then 
            sErrMsg = "Your user name is not valid.  It is either already in use or not between 5 and 12 characters in length.  "
            sErrMsg = sErrMsg & "Please adjust and re-enter."
        End If

        If sErrMsg = vbNullString Then
	        Set rs = Server.CreateObject("ADODB.Recordset")
	        sql = "SELECT UserName FROM MyHist WHERE MyHistID = " & Session("my_hist_id")
	        rs.Open sql, conn, 1, 2
	        rs(0).Value = sUserName
	        rs.Update
	        rs.Close
	        Set rs = Nothing
        End If
    End If

    If sErrMsg = vbNullString Then
        If Not sPassword & "" = "" Then
	        sPassword = Request.Form.Item("password")
	        sConfirmPassword = Request.Form.Item("confirm_password")

            If Not CStr(sPassword) = CStr(sConfirmPassword) Then sErrMsg = "Your passwords do not match.  Please adjust."

            If sErrMsg = vbNullString Then
                'check for password validity
                If ValidPassword(sUserName) = False Then 
                    sErrMsg = "Your password is not valid.  It is either already in use or not between 5 and 12 characters in length.  "
                    sErrMsg = sErrMsg & "Please adjust and re-enter."
                End If
            End If

            If sErrMsg = vbNullString Then
	            Set rs = Server.CreateObject("ADODB.Recordset")
	            sql = "SELECT Password FROM MyHist WHERE MyHistID = " & Session("my_hist_id")
	            rs.Open sql, conn, 1, 2
	            rs(0).Value = sPassword
	            rs.Update
	            rs.Close
	            Set rs = Nothing
            End If
        End If
    End If
End If

'get picture
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PixURL FROM IndPix WHERE PartID = " & Session("etraxc_id")
rs.Open sql, conn2, 1, 2
If rs.RecordCount > 0 Then 	
	If rs(0).Value & "" = "" Then
		sMyPix = "/graphics/photo_na.gif"
	Else
		If FileExists(rs(0).Value) = True Then 
			sMyPix = "http://www.etraxc.com/graphics/ind_pix/" & rs(0).value
		Else
			sMyPix = "/graphics/photo_na.gif"
		End If
	End If
Else
	sMyPix = "/graphics/photo_na.gif"
End If
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Email, BirthDate, ScreenName FROM PartData WHERE PartID = " & Session("etraxc_id")
rs.Open sql, conn2, 1, 2
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName = Replace(rs(1).Value, "''", "'")
sEmail = rs(2).Value
dDOB = rs(3).Value
If Not rs(4).Value & "" = "" Then sScreenName = Replace(rs(4).Value, "''", "'")
rs.Close
Set rs = Nothing

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT Address, City, St, Zip, Phone, CellPhone, Comments FROM PartProfile WHERE PartID = " & Session("etraxc_id")
rs.Open sql, conn2, 1, 2
If Not rs(0).Value & "" = "" Then sAddress = Replace(rs(0).Value, "''", "'")
If Not rs(1).Value & "" = "" Then sCity = Replace(rs(1).Value, "''", "'")
If Not rs(2).Value & "" = "" Then sState = Replace(rs(2).Value, "''", "'")
sPostal = rs(3).Value
sPhone = rs(4).Value
sMobile = rs(5).Value
If Not rs(6).Value & "" = "" Then sComments = Replace(rs(6).Value, "''", "'")
rs.Close
Set rs = Nothing

%>
<!--#include file = "../includes/valid_email.asp" -->
<%

Private Function FileExists(lThisPic)
	FileExists = False
	
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists("C:\Inetpub\h51web\eTRaXC\graphics\ind_pix\" & lThisPic) = True Then FileExists = True
	Set fso = Nothing
End Function

Function ValidUserName(sThisUserName) 
	ValidUserName = True

	If Len(sThisUserName) < 5 Or Len(sThisUserName) > 12 Then ValidUserName = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT UserName FROM MyHist WHERE UserName = '" & sThisUserName & "' AND MyHistID <> " & Session("my_hist_id")
    rs.open sql, conn, 1, 2
    If rs.RecordCount > 0 Then ValidUserName = False
    rs.Close
    Set rs = Nothing
End Function

Function ValidPassword(sThisPassword) 
	ValidPassword = True

	If Len(sThisPassword) < 5 Or Len(sThisPassword) > 12 Then ValidPassword = False

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Password FROM MyHist WHERE Password = '" & sThisPassword & "' AND MyHistID <> " & Session("my_hist_id")
    rs.open sql, conn, 1, 2
    If rs.RecordCount > 0 Then ValidPassword = False
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>My GSE History Profile</title>
<meta name="description" content="Manage my profile for My GSE History account.">
<!--#include file = "../includes/js.asp" -->
</head>
<body>
<div class="container">
    <img src="/graphics/html_header.png" class="img-responsive" alt="Individual Results">
    <h3 class="h3">My GSE History</h3>

    <!--#include file = "my_hist_nav.asp" -->
    <h4 class="h4">My GSE Profile</h4>
   
  	<div class="bg-info">
        NOTE: The information included on this site is considered private.  Under no circumstances will it ever be made available 
        to any third party for any reason without the written permission of the person that it represents.
    </div>

	<%If Not sEmailErr = vbNullString Then%>
		<p class="text-danger"><%=sEmailErr%></p>
	<%End If%>

	<%If Not sErrMsg = vbNullString Then%>
		<p class="text-danger"><%=sErrMsg%></p>
	<%End If%>
				
	<form role="form" class="form" name="event_dir_info" method="post" action="profile.asp">
    <div class="row">
        <div class="col-sm-3">
		    <div class="form_group">
				<label for="first_name">First:</label>
				<input type="text" class="form-control" name="first_name" id="first_name" maxlength = "10" value="<%=sFirstName%>" tabindex="1">
            </div>
            <div class="form_group">
				<label for="city">City:</label>
				<input type="text" class="form-control" name="city" id="city" maxlength = "25" value="<%=sCity%>" tabindex="4">
            </div>
            <div class="form_group">
				<label for="phone">Phone:</label>
				<input type="text" class="form-control" name="phone" id="phone" maxlength = "20" value="<%=sPhone%>" tabindex="7">
			</div>
            <div class="form_group">
				<label for="dob">DOB:</label>
				<input type="text" class="form-control" name="dob" id="dob" value="<%=dDOB%>" tabindex="10">
			</div>
            <hr>
			<div class="form-group">
				<label for="user_name">User Name:</label>
				<input type="text" class="form-control" name="user_name" id="user_name" maxLength="12" tabindex="12">
			</div>
        </div>
        <div class="col-sm-3">
            <div class="form_group">
				<label for="last_name">Last:</label>
				<input type="text" class="form-control" name="last_name" id="last_name" maxlength = "15" value="<%=sLastName%>" tabindex="2">
            </div>
                <div class="form_group">
				<label for="state">State/Prov:</label>
				<input type="text" class="form-control" name="state" id="state" maxlength = "2"  value="<%=sState%>" tabindex="5">
            </div>
            <div class="form_group">
				<label for="mobile">Mobile:</label>
				<input type="text" class="form-control" name="mobile" id="mobile" maxlength = "50" value="<%=sMobile%>" tabindex="8">
			</div>
            <div class="form_group">
                <label for="screen_name">Screen Name:</label>
                <input type="text" class="form-control" name="screen_name" id="screen_name" maxlength = "15" value="<%=sScreenName%>" tabindex="11">
			</div>
            <hr>
			<div class="form-group">
				<label for="password">Password:</label>
				<input type="password" class="form-control" name="password" id="password" maxLength="12" tabindex="13">
			</div>
        </div>
        <div class="col-sm-3">
            <div class="form_group">
				<label for="address">Address:</label>
				<input type="text" class="form-control" name="address" id="address" maxlength = "50" value="<%=sAddress%>" tabindex="3">
            </div>
            <div class="form_group">
				<label for="postal">Postal:</label>
				<input type="text" class="form-control" name="postal" id="postal" maxlength = "8"  value="<%=sPostal%>" tabindex="6">
            </div>
            <div class="form_group">
				<label for="email">Email:</label>
				<input type="text" class="form-control" name="email" id="email" maxlength = "50" value="<%=sEmail%>" tabindex="9">
			</div>
			<br><br><br>
            <hr>
			<div class="form-group">
				<label for="confirm_password">Confirm:</label>
				<input type="password" class="form-control" name="confirm_password" id="confirm_password" maxLength="12" tabindex="14">
			</div>
        </div>
        <div class="col-sm-3" style="padding:10px;">
            <img class="img-responsive center-block" src="<%=sMyPix%>" alt="My Profile">
            <div style="text-align: center;">
                <a href="javascript:pop('part_pix.asp',600,400)">Upload Profile Piture</a>
                <br>
                (Your profile picture will appear here, on results pages of races you run, and in your My-eTRaXC account.)
            </div>
        </div>
    </div>
    <div class="form_group">
		<label for="comments">Comments:</label>
        <textarea class="form-control" name="comments" id="comments" rows="2" tabindex="10"><%=sComments%></textarea>
    </div>
    <div class="form_group">
        <br>
		<input type="hidden" class="form-control" name="submit_info" id="submit_info" value="submit_info">
		<input type="submit" class="form-control" name="submit1" id="submit1" value="Save Changes" tabindex="15">
    </div>
	</form>
</div>
</body>
</html>
<%
conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>