<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lSeriesID
Dim Series()

lSeriesID = Request.QueryString("series_id")
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CCSeriesID, SeriesName, SeriesYear FROM CCSeries ORDER BY SeriesYear DESC, SeriesName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    Series(0, i) = rs(0).Value
	Series(1, i) = Replace(rs(1).Value, "''", "'") & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve Series(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE (Gopher State Events) CC/Nordic Series Results</title>
<meta name="description" content="Gopher State Events (GSE) Cross-Country/Nordic Ski Series Results.">
 
<!--Data Table references-->   
<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css">
<script src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
                              
<script>
     $(document).ready(function () {
        var series = 0;

        $('#series').on('change', function () {
            series = $('#series').val();
            getData(series)
        });
 
        var dt = null;

       function getData(series)  {
            var url = '/series/cc_nordic/results_array.asp?series_id=' + series;
            if(dt) { 
                dt.fnSettings().sAjaxSource = url;
                dt.dataTable().fnDraw();
            }
            else {
                var settings = {
                    bServerSide:true, 
                    sAjaxSource:url, 
                    pagingType: "full_numbers", 
                    fnServerData: function(src, data, cb) {
                           $.post(src, data, cb, "json");
                    }
              };
              dt = $('#standings').dataTable(settings);
            }
       };       

    });
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

    <div class="row">
        <ul class="nav">
            <li class="nav-item"><a class="nav-link" href="race-by-race.asp">Race-by-Race</a></li>
            <li class="nav-item"><a class="nav-link" href="javascript:pop('how_it_works.asp',600,650)">How It Works</a></li>
            <li class="nav-item"><a class="nav-link" href="javascript:window.print();">Print Page</a></li>
        </ul>

        <span style="font-weight: bold;">Series:</span>
        <select name="series" id="series">
            <option value="">&nbsp;</option>
            <%For i = 0 To UBound(Series, 2) - 1%>
                <option value="<%=Series(0, i)%>"><%=Series(1, i)%></option>
            <%Next%>
        </select>

        <table id="standings" class="display" cellspacing ="0" style="width:750px;font-size: 0.8em;">
            <thead>
                <tr>
                    <th>Pl</th>
                    <th>Name</th>
                    <th>School</th>
                    <th>Pts</th>
                </tr>
            </thead>
            <tfoot>
                <tr>
                    <th>PL</th>
                    <th>Name</th>
                    <th>School</th>
                    <th>Pts</th>
                </tr>
            </tfoot>                        
        </table>
    </div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
