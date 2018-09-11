<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i
Dim lPacksID
Dim sPackName, sGender, sSport, sAuthor, sMsg

Dim cdoMessage, cdoConfig

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_pack") = "submit_pack" Then
    sPackName = Replace(Request.Form.Item("pack_name"), "'", "''")
    sSport = Request.Form.Item("sport")
    sGender = Request.Form.Item("gender")

    sql = "INSERT INTO PerfTrkrPacks (PackName, Gender, Sport, PerfTrkrID, WhenCreated) VALUES ('" & sPackName & "', '" & sGender & "', '" & sSport 
    sql = sql & "', " & Session("perf_trkr_id") & ", '" & Now() & "')"
    Set rs = conn2.Execute(sql)
    Set rs = Nothing

    sql = "SELECT r.FirstName, r.LastName FROM Roster r INNER JOIN PerfTrkr p ON r.RosterID = p.RosterID WHERE p.PerfTrkrID = " & Session("perf_trkr_id")
    Set rs = conn2.Execute(sql)
    sAuthor = Replace(rs(0).Value, "''", "'") & " " & Replace(rs(1).Value, "''", "'")
    Set rs = Nothing

    'get pack id
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PerfTrkrPacksID FROM PerfTrkrPacks WHERE PackName = '" & sPackName & "' AND Sport = '" & sSport & "' AND Gender = '" & sGender
    sql = sql & "' AND PerfTrkrID = " & Session("perf_trkr_id") & " ORDER BY PerfTrkrPacksID DESC"
    rs.Open sql, conn2, 1, 2
    lPacksID = rs(0).Value
    rs.Close
    Set rs = Nothing

    'send email to bob and I
    sMsg = sMsg & "A new Performance Tracker Pack has been created.  The details are below."  & vbCrLf & vbCrLf
	sMsg = sMsg & "Pack Name: "  & sPackName & vbCrLf
	sMsg = sMsg & "Gender: "  & sGender  & vbCrLf
	sMsg = sMsg & "Sport: "  & sSport  & vbCrLf
    sMsg = sMsg & "Author: " & sAuthor

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.From = "bob.schneider@gopherstateevents.com"
	    .To = "bob.schneider@gopherstateevents.com"
	    .Subject = "New GSE Performance Tracker Pack"
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing

    'redirect to my packs with this one selected
    Response.Redirect "my_packs.asp?pack_id=" & lPacksID
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Create GSE Performance Tracker Pack</title>

<script>
function chkFlds() {
if (document.new_pack.pack_name.value == '' || 
    document.new_pack.gender.value == '' || 
    document.new_pack.sport.value == '')
{
 	alert('All fields are required!');
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
        <h4 class="h4">Create a Pack</h4>

        <div>
            Performance Tracker follows performance results via "packs".  A pack can be a single participant or a group of participants that are of the same 
            gender and of the same sport.  Please note that this service does not divulge any information that is not already "public" via results lists already 
            online.  What it does is makes those results more personal and formats them in a manner that is more informational.
        </div>
    </div>

    <br>

    <div class="row">
        <div class="col-sm-10">
            <form class="form-inline" name="new_pack" method="post" action="create_pack.asp" onsubmit="return chkFlds();">
            <div class="form-group">
                <label for="pack_name">Pack Name:</label>&nbsp;
                <input class="form-control" type="text" name="pack_name" id="pack_name">&nbsp;&nbsp;
            </div>
            <div class="form-group">
                <label for="sport">Sport:</label>&nbsp;
 		        <select class="form-control" name="sport" id="sport">
                    <option value="">&nbsp;</option>
			        <option value="Nordic Ski">Nordic Ski</option>
			        <option value="Cross-Country">Cross-Country</option>
		        </select>&nbsp;&nbsp;
            </div>
            <div class="form-group">
                <label for="gender">Gender:</label>&nbsp;
		        <select class="form-control" name="gender" id="gender">
                    <option value="">&nbsp;</option>
			        <option value="M">M</option>
			        <option value="F">F</option>
		        </select>&nbsp;&nbsp;
            </div>
            <div class="form-group">
		        <input type="hidden" name="submit_pack" id="submit_pack" value="submit_pack">
		        <input class="form-control" type="submit" name="submit1" id="submit1" value="Create Pack">
            </div>
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
