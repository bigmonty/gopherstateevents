<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sSeriesName, sSport, sComments, sRankBy
Dim iYear, iMaxPts

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then 
    If Month(Date) <= 4 Then
        iYear = Year(Date) - 1
    Else
        iYear = Year(Date)
    End If
End If
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")

Response.Buffer = False		'Turn buffering on
Response.Expires = -1		'Page expires immediately
								
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.item("submit_new") = "submit_new" Then
	sSeriesName = Replace(Request.Form.Item("series_name"), "''", "'")
	sSport =  Request.Form.Item("sport")
	iMaxPts =  Request.Form.Item("max_pts")
	sComments =  Replace(Request.Form.Item("comments"), "''", "'")
	iYear = Request.Form.Item("year")
    sRankBy =  Request.Form.Item("rank_by")

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "INSERT INTO CCSeries (SeriesName, Sport, MaxPts, Comments, SeriesYear, RankBy) VALUES ('" & sSeriesName & "', '" & sSport 
    sql = sql & "', " & iMaxPts & ", '" & sComments & "', " & iYear & ", '" & sRankBy & "')"
	rs = conn.Execute(sql)
	Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE CC/Nordic Series Manager: New Series</title>

<script>
function chkFlds() {
if (document.create_new_series.series_name.value == '' || 
    document.create_new_series.sport.value == '' ||
    document.create_new_series.max_pts.value == '') 
{
 	alert('All fields except comments are required!');
 	return false
 	}
else
 	return true;
}
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../includes/header.asp" -->
	
	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h2 style="margin-left:10px;">Create New CC/Nordic Series</h2>

            <!--#include file = "cc_series_nav.asp" -->

            <h4 class="h4">Create Series</h4>

            <form role="form" class="form-horizontal" name="create_new_series" method="post" action="new_series.asp" onsubmit="return chkFlds();">
            <div class="form-group row">
                <label class="col-sm-2 col-form-label" for="series_name">Series Name:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="series_name" id="series_name">
                </div>
                <label class="col-sm-2 col-form-label" for="sport">Sport:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="sport" id="sport">
                        <option value="">&nbsp;</option>
                        <option value="Nordic Ski">Nordic Ski</option>
                        <option value="Cross-Country">Cross-Country</option>
                    </select>
                </div>
            </div>
            <div class="form-group row">
                <label class="col-sm-2 col-form-label" for="year">Year Beginning:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="year" id="year">
                        <%For i = 2010 To Year(Date)%>
                            <%If CInt(iYear) = CInt(i) Then%>
                                <option value="<%=i%>" selected><%=i%></option>
                            <%Else%>
                                <option value="<%=i%>"><%=i%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                </div>
                 <label class="col-sm-2 col-form-label" for="max_pts">Max Pts:</label>
                <div class="col-sm-4">
                    <input class="form-control" type="text" name="max_pts" id="max_pts" size="3" maxlength="4">
                </div>
            </div>
            <div class="form-group row">
               <label class="col-sm-2 col-form-label" for="">Rank By:</label>
                <div class="col-sm-4">
                    <select class="form-control" name="rank_by" id="rank_by">
                        <option value="points">Points</option>
                        <option value="pctle">Pctle</option>
                    </select>
                </div>
                <label class="col-sm-2 col-form-label" for="comments">Comments:</label>
                <div class="col-sm-4">
                    <textarea class="form-control" name="comments" id="comments" rows="3"></textarea>
                </div>
            </div>
            <div class="form-group">
                <input type="hidden" name="submit_new" id="submit_new" value="submit_new">
                <input class="form-control" type="submit" name="submit1" id="submit1" value="Create Series">
            </div>
            </form>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
