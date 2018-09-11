<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql
Dim lCellProvider
Dim i
Dim sFollowerCell, sFollowerName, sRelationship, sNotifyFollower, sFollowerEmail
Dim CellProviders

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_changes") = "submit_changes" Then
	sFollowerCell = Trim(Request.Form.Item("follower_cell"))
    lFollowerProvider = Request.Form.Item("follower_provider")
	sFollowerEmail = Trim(Request.Form.Item("follower_email"))
	sNotifyFollower = Trim(Request.Form.Item("notify_follower"))
	sFollowerName = Trim(Request.Form.Item("follower_name"))
	sRelationship = Trim(Request.Form.Item("relationship"))

    If sErrMsg = vbNullString Then
        'check for email validity
        If ValidEmail(sFollowerEmail) = False Then sErrMsg = "Your email address does not appear to be in a valid format.  Please re-enter."
    End If

    'if everything checks out, save changes
    If sErrMsg = vbNullString Then
    End If
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
<title>GSE Performance Tracker Results Notification</title>

<script>
function chkFlds() {
if (document.edit_follower.email.value == '' ||
    document.edit_follower.name.value == '' ||
    document.edit_follower.relationship.value == '')

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
    <h5 class="h5 bg-warning" style="padding:2px;">Edit Follower</h5>

    <form class="form-horizontal" name="edit_follower" method="Post" action="edit_follower.asp" onSubmit="return chkFlds();">
    <div class="form-group row">
        <label for="follower_name" class="control-label col-sm-3">Name:</label>
        <div class="col-sm-9">
            <input class="form-control" type="text" name="follower_name" id="follower_name" maxLength="50">
        </div>
    </div>
    <div class="form-group row">
        <label for="follower_email" class="control-label col-sm-3">Email:</label>
        <div class="col-sm-9">
            <input class="form-control" type="text" name="follower_email" id="follower_email" maxLength="50">
        </div>
    </div>
    <div class="form-group row">
        <label for="relationship" class="control-label col-sm-3">Relationship:</label>
        <div class="col-sm-9">
            <select class="form-control" name="relationship" id="relationship">
                <option value="">&nbsp;</option>
                <option value="Family">Family</option>
                <option value="Friend">Friend</option>
                <option value="Other">Other</option>
            </select>
        </div>
    </div>
    <div class="form-group row">
        <label for="follower_cell" class="control-label col-sm-3">Cell Phone:</label>
        <div class="col-sm-9">
            <input class="form-control" type="text" name="follower_cell" id="follower_cell" maxLength="50">
        </div>
    </div>
    <div class="form-group row">
        <label for="follower_provider" class="control-label col-sm-3">Provider:</label>
        <div class="col-sm-9">
            <select class="form-control" name="follower_provider" id="follower_provider">
                <option value="">&nbsp;</option>
                <%For i = 0 To UBound(CellProviders, 2)%>
                    <option value="<%=CellProviders(0, i)%>"><%=CellProviders(1, i)%></option>
                <%Next%>
            </select>
        </div>
    </div>
    <div class="form-group row">
        <label for="notify_follower" class="control-label col-sm-3">Notify:</label>
        <div class="col-sm-9">
            <select class="form-control" name="notify_follower" id="notify_follower">
                <option value="n">No</option>
                <option value="y" selected>Yes</option>
            </select>
        </div>
    </div>
    <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
    <input class="form-control" type="submit" name="submit2" id="submit2" value="Add Follower">
    </form>
</div>

<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.close
Set conn = Nothing
%>
</html>
