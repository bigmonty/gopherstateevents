<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2, conn2
Dim i
Dim lEventDirID
Dim sUserName, sPassword, sErrMsg
Dim bValidateInput, bNotFound
Dim MyEvents()
Dim fs, fname, sFileName

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

Dim sRandPic
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, PixName FROM RacePix ORDER BY NEWID()"
rs.Open sql, conn, 1, 2
sRandPic = "/gallery/" & rs(0).Value & "/" & Replace(rs(1).Value, "''", "'")
rs.Close
Set rs = Nothing

bNotFound = False

If Request.Form.Item("submit_login") = "submit_login" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'contract' ORDER BY AuthAccessID DESC"
	rs.Open sql, conn, 1, 2
	If rs.RecordCount > 0 Then Session("access_contract") = "y"
	rs.Close
	Set rs = Nothing

	If Session("access_contract") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
		
		sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
		If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))
		
		If sHackMsg = vbNullString Then
            bNotFound = False
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT EventDirID FROM EventDir WHERE UserID = '" & sUserName & "' AND Password = '" & sPassword & "'"
			rs.Open sql, conn, 1, 2
			If rs.RecordCount > 0 Then 
                lEventDirID = rs(0).Value
            Else
                bNotFound = True
            End If
			rs.Close
			Set rs = Nothing

            'log this login and then redirect
            If bNotFound = True Then
                sErrMsg = "We are sorry but those login credentials were not found.  Please ensure that you have selected the "
                sErrMsg = sErrMsg & "correct role and try again or <a href=mailto:bob.schneider@gopherstateevents.com>contact</a> H51Software, LLC "
                sErrMsg = sErrMsg & "for assistance."
			End If
		End If
	End If
End If

If CStr(lEVentDirID) & "" = "" Then lEventDirID = 0

'log this user if they are just entering the site
If Session("access_login") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'contract')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'contract'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_contract")
End If

'get contracts if signed in
ReDim MyEvents(3, 0)
If Not CLng(lEventDirID) = 0 Then
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventName, EventDate FROM Events WHERE EventDirID = " & lEventDirID & " ORDER BY EventDate DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        MyEvents(0, i) = Replace(rs(1).Value, "''", "'")
        MyEvents(1, i) = rs(2).Value
        If ChkCntrct(rs(0).Value, rs(2).Value) = True Then 
            MyEvents(2, i) = "View"
            MyEvents(3, i) = "/contracts/" & Year(rs(2).Value) & "/" & rs(0).Value & ".pdf"
        Else
            MyEvents(2, i) = "N/A"
            MyEvents(3, i) = ""
        End If
        i = i + 1
        ReDim Preserve MyEvents(3, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

Private Function ChkCntrct(lThisEvent, dEventDate)
    Set fs=Server.CreateObject("Scripting.FileSystemObject")
    sFileName = "C:\Inetpub\h51web\gopherstateevents\contracts\" & Year(dEventDate) & "\" & lThisEvent & ".pdf"
    ChkCntrct = fs.FileExists(sFileName)
    Set fs = Nothing
End Function

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<!--#include file = "../../includes/clean_input.asp" -->
<%
Set cdoConfig = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title>GSE Contract Viewer</title>

<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">
<meta name="description" content="Sign in to Gopher State Events (GSE) contract page.">





 

<script>
function chkFlds() {
if (document.site_login.user_name.value == '' || 
    document.site_login.password.value == '') 
{
 	alert('All fields are required!');
 	return false
 	}
else
 	return true;
}
</script>

<style type="text/css">
    th,td{
        padding: 2px 0 2px 5px;
    }
</style>
</head>

<body onload="document.site_login.user_name.focus();">
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../includes/cmng_evnts.asp" -->
		<div class="col-md-10">
             <!--#include file = "../includes/banner_ad.asp" -->

			<h1 style="margin:5px;padding:5px;background-color:#ececd8;font-size:1.1em;">GSE Contract Viewer</h1>

			<a href="<%=sRandPic%>" onclick="openThis(this.href,1024,768);return false;"><img src="<%=sRandPic%>" alt="<%=sRandPic%>" 
			style="width:300px;margin:5px 10px 5px 5px;float:right;"></a>
            
			<%If Not sHackMsg = vbNullString Then%>
				<p><%=sHackMsg%></p>
			<%Else%>
                <%If CLng(lEventDirID) = 0 Then%>
				    <%If Not sErrMsg = vbNullString Then%>
					    <p><%=sErrMsg%></p>
				    <%End If%>
				
				    <form name="site_login" method="Post" action="contract.asp" onSubmit="return chkFlds();">
				    <table style="margin:50px;">
					    <tr>
						    <th>User Name:</th>
						    <td><input type="text" name="user_name" id="user_name" size="12"  maxlength="12" value="<%=sUserName%>"></td>
					    </tr>
					    <tr>
						    <th>Password:</th>
						    <td><input type="password" name="password" id="password" size="12"  maxlength="12" value="<%=sPassword%>"></td>
					    </tr>
					    <tr>
						    <td style="text-align:center;" colspan="2">
							    <input type="hidden" name="submit_login" id="submit_login" value="submit_login">
							    <input type="submit" name="submit1" id="submit1" value="Login">
						    </td>
					    </tr>
					    <tr>
						    <td style="background-color:#efefef;text-align:center;" colspan="2">
							    <a href="javascript:pop('forgot_contract_login.asp',600,750)" style="font-size:0.85em;">Forgot Sign In?</a>
						    </td>
					    </tr>
				    </table>
				    </form>
                <%Else%>
                    <div style="margin: 10px;">
                        <h4 class="h4">Select Contract To View</h4>

                        <table>
                            <tr>
                                <th>No.</th>
                                <th>Event</th>
                                <th>Date</th>
                                <th>Contract</th>
                            </tr>
                            <%For i = 0 To UBound(MyEvents, 2) - 1%>
                                <%If i mod 2 = 0 Then%>
                                    <tr>
                                        <td class="alt"><%=i + 1%>)</td>
                                        <td class="alt"><%=MyEvents(0, i)%></td>
                                        <td class="alt"><%=MyEvents(1, i)%></td>
                                        <td class="alt" style="text-align: center;">
                                            <%If MyEvents(2, i) ="View" Then%>
                                                <a href="javascript:pop('<%=MyEvents(3, i)%>',800,600)"><%=MyEvents(2, i)%></a>
                                            <%Else%>
                                                <%=MyEvents(2, i)%>
                                            <%End If%>
                                        </td>
                                    </tr>
                                <%Else%>
                                    <tr>
                                        <td><%=i + 1%>)</td>
                                        <td><%=MyEvents(0, i)%></td>
                                        <td><%=MyEvents(1, i)%></td>
                                        <td style="text-align: center;">
                                            <%If MyEvents(2, i) ="View" Then%>
                                                <a href="javascript:pop('<%=MyEvents(3, i)%>',800,600)"><%=MyEvents(2, i)%></a>
                                            <%Else%>
                                                <%=MyEvents(2, i)%>
                                            <%End If%>
                                        </td>
                                    </tr>
                                <%End If%>
                            <%Next%>
                        </table>
                    </div>
                <%End If%>
			<%End If%>
		</div>
		<!--#include file = "../includes/vira_sponsors.asp" -->
	</div>
	<!--#include file = "../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%>