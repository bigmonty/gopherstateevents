<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim lCellProvider
Dim i
Dim sFirstName, sLastName, sGender, sCellPhone, sEmail, sScreenName, sUserName, sPassword, sNewUserName
Dim sResultsNotif, sNewPassword, sConfirmPassword, sErrMsg, sEmailErr, sImage
Dim iMonth, iDay, iYear
Dim dDOB
Dim CellProviders
Dim sFileName, fs

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_remove_pix") = "submit_remove_pix" Then
    'delete image
	Set fs=Server.CreateObject("Scripting.FileSystemObject") 
	If fs.FileExists("c:\inetpub\h51web\gopherstateevents\cc_meet\perf_tracker\images\" & sFileName) = True Then
		fs.DeleteFile("c:\inetpub\h51web\gopherstateevents\cc_meet\perf_tracker\images\" & sFileName)
	End If
	Set fs=Nothing

    'remove it from db
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Image FROM PerfTrkr WHERE PerfTrkrID = " & Session("perf_trkr_id")
    rs.Open sql, conn, 1, 2
    rs(0).Value = ""
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
	sCellPhone = Trim(Request.Form.Item("cell_phone"))
    lCellProvider = Request.Form.Item("cell_provider")
    sScreenName = Trim(Request.Form.Item("screen_name"))
	sEmail = Trim(Request.Form.Item("email"))
	sUserName = Trim(Request.Form.Item("old_user_name"))
	sPassword = Trim(Request.Form.Item("old_password"))
	sNewUserName = Trim(Request.Form.Item("new_user_name"))
	sNewPassword = Trim(Request.Form.Item("new_password"))
	sConfirmPassword = Trim(Request.Form.Item("confirm_password"))
    sResultsNotif = Request.Form.Item("results_notif")

    If sErrMsg = vbNullString Then
        'check for email uniqueness
        If UniqueEmail(sEmail) = False Then 
            sErrMsg = "Your email address is already in our system.  If you believe you are the only human using this address, "
            sErrMsg = sErrMsg & "please log in to your existing account.  If a friend or family member may be using this email "
            sErrMsg = sErrMsg & "address please use another one. "
        End If
    End If

    If sErrMsg = vbNullString Then
        'check for email validity
        If ValidEmail(sEmail) = False Then sErrMsg = "Your email address does not appear to be in a valid format.  Please re-enter."
    End If

    'if everything checks out
    If sErrMsg = vbNullString Then
	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT CellPhone, CellProvider, Email, ScreenName, ResultsNotif FROM PerfTrkr WHERE PerfTrkrID = " 
        sql = sql & Session("perf_trkr_id")
	    rs.Open sql, conn, 1, 2
	    rs(0).Value = sCellPhone
	    rs(1).Value = lCellProvider

	    If sEmail & "" = "" Then
		    rs(2).Value = rs(2).OriginalValue
	    Else
            If ValidEmail(sEmail) = False Then sEmailErr = "Your email address does not appear to be in a valid format.  Please re-enter.  Some work was not done."

            If sEmailErr = vbNullString Then
                If UniqueEmail(sEmail) = False Then sEmailErr = "Your new email address does not appear to be unique.  Please select another one.  Some work was not done."
            End If

            If sEmailErr = vbNullString Then 
		        rs(2).Value = sEmail
            Else
                rs(2).Value = rs(2).OriginalValue
            End If
	    End if

	    rs(3).Value = sScreenName
        rs(4).Value = sResultsNotif
        rs.Update
        rs.Close
        Set rs = Nothing
    End If

    If Not (sUserName & "" = "" Or sPassword & "" = "") Then
        'check for correct user name and password entry
	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT UserName, Password FROM PerfTrkr WHERE PerfTrkrID = " & Session("perf_trkr_id") & " AND UserName = '" & sUserName
        sql = sql & "' AND Password = '" & sPassword & "'"
        rs.Open sql, conn, 1, 2      
        If rs.RecordCount > 0 Then
            'check for user name validity
            If ValidUserName(sUserName) = False Then 
                sErrMsg = "Your user name is not valid.  It is either already in use or not between 5 and 12 characters in length.  "
                sErrMsg = sErrMsg & "Please adjust and re-enter."
            End If

            If sErrMsg = vbNullString Then
                'check for password validity
                If ValidPassword(sUserName) = False Then 
                    If Not CStr(sNewPassword) = CStr(sConfirmPassword) Then sErrMsg = "Your passwords do not match.  Please adjust."
                End If
            End If

            If sErrMsg = vbNullString Then
                'check for password validity
                If ValidPassword(sUserName) = False Then 
                    sErrMsg = "Your password is not valid.  It is either already in use or not between 5 and 12 characters in length.  "
                    sErrMsg = sErrMsg & "Please adjust and re-enter."
                End If
            End If

            If sErrMsg = vbNullString Then
                rs(0).Value = sNewUserName
                rs(1).Value = sNewPassword
                rs.Update
            End If
        Else
            sErrMsg = "The credentials you supplied do not match our records for this account.  Please re-enter."
        End If

        rs.Close
        Set rs = Nothing
    End If
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.FirstName, r.LastName, r.Gender, p.CellPhone, p.Email, p.ScreenName, p.CellProvider, p.Image, "
sql = sql & "p.UserName, p.ResultsNotif FROM PerfTrkr p INNER JOIN Roster r "
sql = sql & "ON p.RosterID = r.RosterID WHERE p.PerfTrkrID = " & Session("perf_trkr_id")
rs.Open sql, conn, 1, 2
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName = Replace(rs(1).Value, "''", "'")
sGender = rs(2).Value
sCellPhone = rs(3).Value
sEmail = rs(4).Value
If Not rs(5).Value & "" = "" Then sScreenName = Replace(rs(5).Value, "''", "'")
lCellProvider = rs(6).Value
sImage = rs(7).Value
sUserName = rs(8).Value
sResultsNotif = rs(9).Value
rs.Close
Set rs = Nothing

If sImage & "" = "" Then 
    sImage = "images/pna.png"
Else
    sImage = "images/" & sImage
End If

Function ValidUserName(sThisUserName) 
	ValidUserName = True

	If Len(sThisUserName) < 5 Or Len(sThisUserName) > 12 Then ValidUserName = False

    If ValidPassword = True Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT UserName FROM PerfTrkr WHERE UserName = '" & sThisUserName & "' AND PerfTrkrID <> " & Session("perf_trkr_id")
        rs.open sql, conn, 1, 2
        If rs.RecordCount > 0 Then ValidUserName = False
        rs.Close
        Set rs = Nothing
    End If
End Function

Function ValidPassword(sThisPassword) 
	ValidPassword = True

	If Len(sThisPassword) < 5 Or Len(sThisPassword) > 12 Then ValidPassword = False

    If ValidPassword = True Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT Password FROM PerfTrkr WHERE Password = '" & sThisPassword & "' AND PerfTrkrID <> " & Session("perf_trkr_id")
        rs.open sql, conn, 1, 2
        If rs.RecordCount > 0 Then ValidPassword = False
        rs.Close
        Set rs = Nothing
    End If
End Function

%>
<!--#include file = "../../includes/valid_email.asp" -->
<%

Function UniqueEmail(sThisEmail) 
	UniqueEmail = True

    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Email FROM PerfTrkr WHERE Email = '" & sThisEmail & "' AND PerfTrkrID <> " & Session("perf_trkr_id")
    rs2.open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then UniqueEmail = False
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>My GSE Performance Tracker Profile</title>

<script>
function chkFlds() {
if (document.edit_accnt.email.value == '')

{
 	alert('First name, last name, email, and dob fields must be populated!');
 	return false
 	}
else
 	return true;
}
</script>
</head>

<body>
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->
    <!--#include file = "perf_trkr_nav.asp" -->

    <div class="row">
        <h3 class="h3">GSE Performance Tracker Profile For <%=sFirstName%>&nbsp;<%=sLastName%> (<%=sGender%>)</h3>
   
  	    <div>
            NOTE: The information included on this site is considered private.  Under no circumstances will it ever be made available to any third party 
            for any reason without the written permission of the person that it represents.
        </div>

	    <%If Not sEmailErr = vbNullString Then%>
		    <p><%=sEmailErr%></p>
	    <%End If%>
    </div>

    <br>

    <div class="row">
        <div class="col-sm-8">
  	        <form class="form-horizontal" name="edit_accnt" method="Post" action="profile.asp" onSubmit="return chkFlds();">
            <div class="form-group row">
                <label for="screen_name" class="control-label col-sm-2">Screen Name:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="screen_name" id="screen_name" maxLength="12" value="<%=sScreenName%>">
                </div>
                <label for="email" class="control-label col-sm-2">Email:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="email" id="email" maxLength="50" value="<%=sEmail%>">
                </div>
            </div>
            <div class="form-group row">
                <label for="cell_phone" class="control-label col-sm-2">Cell Phone:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="cell_phone" id="cell_phone" maxLength="50" value="<%=sCellPhone%>">
                </div>
                <label for="cell_provider" class="control-label col-sm-2">Provider:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="cell_provider" id="cell_provider">
                        <option value="">&nbsp;</option>
                        <%For i = 0 To UBound(CellProviders, 2)%>
                            <%If CLng(lCellProvider) = CLng(CellProviders(0, i)) Then%>
                                <option value="<%=CellProviders(0, i)%>" selected><%=CellProviders(1, i)%></option>
                            <%Else%>
                                <option value="<%=CellProviders(0, i)%>"><%=CellProviders(1, i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </div>
            </div>

            <div class="form-group row">
                <label for="old_user_name" class="control-label col-sm-2">User Name:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="old_user_name" id="old_user_name" maxLength="12" value="<%=sUserName%>">
                </div>
                <label for="new_user_name" class="control-label col-sm-2">New:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="new_user_name" id="new_user_name" maxLength="12">
                </div>
            </div>

            <div class="form-group row">
                <label for="old_password" class="control-label col-sm-2">Password:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="password" name="old_password" id="old_password" maxLength="12">
                </div>
                <label for="password" class="control-label col-sm-2">New Pwd:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="password" name="password" id="password" maxLength="12">
                </div>
            </div>
            <div class="form-group row">
                <label for="confirm_password" class="control-label col-sm-2">Confirm:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="password" name="confirm_password" id="confirm_password" maxLength="12">
                </div>
                <label for="results_notif" class="control-label col-sm-2">Results Notif:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="results_notif" id="results_notif">
                        <%If sResultsNotif = "y" Then%>
                            <option value="n">No</option>
                            <option value="y" selected>Yes</option>
                        <%Else%>
                            <option value="n">No</option>
                            <option value="y">Yes</option>
                        <%End If%>
                    </select>
                </div>
            </div>
	        <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
	        <input class="form-control" type="submit" name="submit1" id="submit1" value="Save Changes">
	        </form>
        </div>
        <div class="col-sm-2">
            <a href="javascript:pop('upload_pix.asp?this_user=<%=Session("perf_trkr_id")%>',600,400)">Upload New Image</a>
                
            <img src="<%=sImage%>" alt="My Pix" class="img-responsive" style="margin-left:10px;">
             
			<form class="form" name="remove_image" method="Post" action="profile.asp">
			<input type="hidden" name="submit_remove_pix" id="submit_remove_pix" value="submit_remove_pix">
			<input class="form-control" type="submit" name="submit2" id="submit2" value="Remove Picture">
			</form>
        </div>
        <div class="col-sm-2">
            <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
            <!-- GSE Vertical ad -->
            <ins class="adsbygoogle"
                    style="display:block"
                    data-ad-client="ca-pub-1381996757332572"
                    data-ad-slot="6120632641"
                    data-ad-format="auto"></ins>
            <script>
            (adsbygoogle = window.adsbygoogle || []).push({});
            </script>
        </div>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.close
Set conn = Nothing

conn2.close
Set conn2 = Nothing
%>
</html>
