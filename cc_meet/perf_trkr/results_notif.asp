<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lFollowerProvider
Dim i, j
Dim sFirstName, sLastName
Dim sFollowerCell, sFollowerName, sRelationship, sNotifyFollower, sFollowerEmail, sScreenName, sImage, sErrMsg
Dim sMsg
Dim CellProviders, Followers(), Relationships(2)
Dim cdoMessage, cdoConfig

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get relatiosnhips
Relationships(0) = "Family"
Relationships(1) = "Friend"
Relationships(2) = "Other"

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_follower") = "submit_follower" Then
	sFollowerCell = Trim(Request.Form.Item("follower_cell"))
    lFollowerProvider = Request.Form.Item("follower_provider")
	sFollowerEmail = Trim(Request.Form.Item("follower_email"))
	sNotifyFollower = Trim(Request.Form.Item("notify_follower"))
	sFollowerName = Trim(Request.Form.Item("follower_name"))
	sRelationship = Trim(Request.Form.Item("relationship"))

    If sErrMsg = vbNullString Then
        'check for email uniqueness
        If UniqueEmail(sFollowerEmail) = False Then 
            sErrMsg = "This email address is already in our system for you or one of your followers.  If you believe "
            sErrMsg = sErrMsg & "this is not the case, please contact bob.schneider@gopherstateevents.com.  "
        End If
    End If

    If sErrMsg = vbNullString Then
        'check for email validity
        If ValidEmail(sFollowerEmail) = False Then sErrMsg = "Your email address does not appear to be in a valid format.  Please re-enter."
    End If

    'if everything checks out
    If sErrMsg = vbNullString Then
        sql = "INSERT INTO PTFollowers (PerfTrkrID, FollowerName, Relationship, Email, CellPhone, CellProvider, ResultsNotif) VALUES ("
        sql = sql & Session("perf_trkr_id") & ", '" & sFollowerName & "', '" & sRelationship & "', '" & sFollowerEmail & "', '" & sFollowerCell
        sql = sql & "', " & lFollowerProvider & ", '" & sNotifyFollower & "')"
        SEt rs = conn.Execute(sql)
        Set rs = Nothing

        'email participant, follower, and gse
	    sMsg = "Hello.  This is a notification that " & Session("my_name") & " has invited you to follow their cross-country and/or Nordic Ski "
        sMsg = sMsg & "racing.  They will do this using GSE's Performance Tracker (www.gopherstateevents.com)!  This results sharing can come in the "
        sMsg = sMsg & "form of emails, text messages, or both. The purpose of Performance Tracker is to assist athletes in tracking and sharing their "
        sMsg = sMsg & "race results with those who might be interested in their successes but may not be able to attend all of their meets. " & vbCrLf & vbCrLf
        
        sMsg = sMsg & "PLEASE NOTE:  We at GSE did not add you to this listing.  The person mentioned above did.  We assume you know them well.  "
        sMsg = sMsg & "Your information will never be used for any purpose other than informing you of these results.  If you are uneasy about this, "
        sMsg = sMsg & "or if you do not know this person, please indicate that by replying to this email and we will delete you from their list of followers. " & vbCrLf & vbCrLf
	
	    sMsg = sMsg & "Sincerely~ " & vbCrLf
	    sMsg = sMsg & "Bob Schneider " & vbCrLf
	    sMsg = sMsg & "Owner: Gopher State Events, LLC " & vbCrLf
        sMsg = sMsg & "bob.schneider@gopherstateevents.com " & vbCrLf
        sMsg = sMsg & "612.720.8427 " & vbCrLf  & vbCrLf

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	    Set cdoMessage = CreateObject("CDO.Message")
	    With cdoMessage
		    Set .Configuration = cdoConfig
		    .To = sFollowerEmail
		    .From = "bob.schneider@gopherstateevents.com"
	        .CC = "bob.schneider@gopherstateevents.com;" & Session("my_email")
	        .Subject = "Invitation From " & Session("my_name")
		    .TextBody = sMsg
		    .Send
	    End With
	    Set cdoMessage = Nothing
	    Set cdoConfig = Nothing

        sFollowerCell = vbNullString
        lFollowerProvider = vbNullString
        sFollowerEmail = vbNullString
        sNotifyFollower = vbNullString
        sFollowerName = vbNullString
        sRelationship = vbNullString
    End If
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT r.FirstName, r.LastName, p.Image FROM PerfTrkr p INNER JOIN Roster r ON p.RosterID = r.RosterID "
sql = sql & "WHERE p.PerfTrkrID = " & Session("perf_trkr_id")
rs.Open sql, conn, 1, 2
sFirstName = Replace(rs(0).Value, "''", "'")
sLastName = Replace(rs(1).Value, "''", "'")
sImage = rs(2).Value
rs.Close
Set rs = Nothing

If sImage & "" = "" Then 
    sImage = "images/pna.png"
Else
    sImage = "images/" & sImage
End If

%>
<!--#include file = "../../includes/valid_email.asp" -->
<%

i = 0
ReDim Followers(6, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PTFollowersID, FollowerName, Relationship, Email, CellPhone, CellProvider, ResultsNotif FROM PTFollowers WHERE PerfTrkrID = " 
sql = sql & Session("perf_trkr_id") & " ORDER BY FollowerName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    For j = 0 To 6
        Followers(j, i) = rs(j).Value
    Next
    i = i + 1
    ReDim Preserve Followers(6, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If CStr(lFollowerProvider) = vbNullString Then lFollowerProvider = "0"

Function UniqueEmail(sThisEmail) 
	UniqueEmail = True

    'check account holder's email
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Email FROM PerfTrkr WHERE Email = '" & sThisEmail & "' AND PerfTrkrID = " & Session("perf_trkr_id")
    rs2.open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then UniqueEmail = False
    rs2.Close
    Set rs2 = Nothing

    'make sure no other followers for this person has this email
    'it is possible for followers of other pt participants to have this email
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Email FROM PTFollowers WHERE Email = '" & sThisEmail & "' AND PerfTrkrID = " & Session("perf_trkr_id")
    rs2.open sql2, conn, 1, 2
    If rs2.RecordCount > 0 Then UniqueEmail = False
    rs2.Close
    Set rs2 = Nothing
End Function

Private Function GetProvider(lProviderID)
    sql = "SELECT Provider FROM CellProviders WHERE CellProvidersID = " & lProviderID
    Set rs = conn.Execute(sql)
    GetProvider = rs(0).Value
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Performance Tracker Results Notification</title>

<script>
function chkFlds() {
if (document.add_follower.follower_email.value == '' ||
    document.add_follower.follower_name.value == '' ||
    document.add_follower.relationship.value == '')

{
 	alert('All fields are required when adding followers except cell phone and cell provider!');
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
        <h4 class="h4">GSE Results Notification for <%=sFirstName%>&nbsp;<%=sLastName%> </h4>

        <div>
            <p>
                This utility is designed to notify those who you would like notified, either by email or text message, 
                your results in any GSE-timed cross-country or Nordic Ski event.  List parents, friends, and others who might 
                want to know how your race went.  Notifications are usually sent at the conclusion of each race or 
                sooner.
            </p>

            <p>
                NOTE:  Do NOT provide email or mobile information for any persons without their knowledge and permission.  THIS INFORMATION WILL
                NEVER BE USED FOR ANY PURPOSE OTHER THAN RESULTS NOTIFICATIONS.  PERIOD!
            </p>
        </div>
        <br>
    </div>
    <div class="row">
        <div class="col-sm-10">
            <h5 class="h5 text-warning" style="padding:2px;">Add New Follower</h5>

            <p style="font-size:0.8em;" class="text-warning">
                All fields are required but if, for some reason, you would like to add this follower and not have them receive results 
                notification, just set "Notify" to "No".
            </p>

            <%If Not sErrMsg = vbNullString Then%>
                <p class="bg-danger"><%=sErrMsg%></p>
            <%End If%>

            <form class="form-horizontal" name="add_follower" method="Post" action="results_notif.asp" onSubmit="return chkFlds();">
            <div class="form-group row">
                <label for="follower_name" class="control-label col-sm-2">Name:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="follower_name" id="follower_name" maxLength="50" value="<%=sFollowerName%>">
                </div>
                <label for="follower_email" class="control-label col-sm-2">Email:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="follower_email" id="follower_email" maxLength="50" value="<%=sFollowerEmail%>">
                </div>
            </div>
            <div class="form-group row">
                <label for="relationship" class="control-label col-sm-2">Relationship:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="relationship" id="relationship">
                        <option value="">&nbsp;</option>
                        <%For i = 0 To UBound(Relationships)%>
                            <%If Relationships(i) = sRelationship Then%>
                                <option value="<%=Relationships(i)%>" selected><%=Relationships(i)%></option>
                            <%Else%>
                                <option value="<%=Relationships(i)%>"><%=Relationships(i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </div>
                <label for="follower_cell" class="control-label col-sm-2">Cell Phone:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="follower_cell" id="follower_cell" maxLength="50" value="<%=sFollowerCell%>">
                </div>
            </div>
            <div class="form-group row">
                <label for="follower_provider" class="control-label col-sm-2">Provider:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="follower_provider" id="follower_provider">
                        <option value="">&nbsp;</option>
                        <%For i = 0 To UBound(CellProviders, 2)%>
                            <%If CLng(CellProviders(0, i)) = CLng(lFollowerProvider) Then%>
                                <option value="<%=CellProviders(0, i)%>" selected><%=CellProviders(1, i)%></option>
                            <%Else%>
                                <option value="<%=CellProviders(0, i)%>"><%=CellProviders(1, i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </div>
                <label for="notify_follower" class="control-label col-sm-2">Notify:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="notify_follower" id="notify_follower">
                        <option value="n">No</option>
                        <option value="y" selected>Yes</option>
                    </select>
                </div>
            </div>
            <input type="hidden" name="submit_follower" id="submit_follower" value="submit_follower">
            <input class="form-control" type="submit" name="submit2" id="submit2" value="Add Follower">
            </form>
            <hr>
            <h5 class="h5">Existing Followers</h5>
            <a href="results_notif.asp">Refresh Page</a>
            <form class="form" name="followers" method="Post" action="results_notif.asp">
            <table class="table table-striped">
                <tr>
                    <th>No.</th>
                    <th>Name (click to edit)</th>
                    <th>Relationship</th>
                    <th>Email</th>
                    <th>Cell Phone</th>
                    <th>Provider</th>
                    <th>Notify?</th>
                </tr>
                <%For i = 0 To UBound(Followers, 2) - 1%>
                    <tr>
                        <td><%=i + 1%></td>
                        <td><a href="javascript:pop('edit_follower.asp?follower_id=<%=Followers(0,i)%>',1000,500)"><%=Followers(1, i)%></a></td>
                        <td><%=Followers(2, i)%></td>
                        <td><%=Followers(3, i)%></td>
                        <td><%=Followers(4, i)%></td>
                        <td><%=GetProvider(Followers(5, i))%></td>
                        <td><%=Followers(6, i)%></td>
                    </tr>
                <%Next%>
            </table>
        </div>
        <div class="col-sm-2">
            <img src="<%=sImage%>" alt="My Pix" class="img-responsive" style="width:150px;">
            <br>
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
%>
</html>
