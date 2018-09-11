<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lSeriesID
Dim iYear
Dim Series()

If CStr(lSeriesID) = vbNullString Then lSeriesID = 0
If Not IsNumeric(lSeriesID) Then Response.Redirect "http://www.google.com"
If CLng(lSeriesID) < 0 Then Response.Redirect "http://www.google.com"

iYear = Request.QueryString("year")
If CStr(iYear) = vbNullString Then iYear = Year(date)
If IsNumeric(iYear) = False Then Response.Redirect("http://www.google.com")
If CLng(iYear) < 0 Then Response.Redirect "http://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Series(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID, SeriesName, SeriesYear FROM Series ORDER BY SeriesYear DESC, SeriesName"
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
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE (Gopher State Events) Series Results: Age-Neutral</title>
<meta name="description" content="Gopher State Events (GSE) race series results.">
 
<link href="//cdn.datatables.net/1.10.2/css/jquery.dataTables.css" rel="stylesheet" type="text/css">
    
<script type="text/javascript" src="//code.jquery.com/jquery-2.1.4.min.js"></script>
<script type="text/javascript" src="//cdn.datatables.net/1.10.2/js/jquery.dataTables.min.js"></script>

<!-- bootstrap JavaScript & CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css">
<script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
  
<style type="text/css">
    div#standings_info, div#standings_paginate{
        display: none;
    }
</style>  

<script>
    $(document).ready(function () {
        $('body').delegate('#series,#gender,#standings_length select', 'change', function () {
            getData()
        });

        $('body').delegate('table#standings th[data-sort], div#standings_wrapper a', 'click', function (e) {
            e.preventDefault();
            e.stopPropagation();
             e.stopImmediatePropagation();
            $("th").removeClass("sorting_asc").removeClass("sorting_desc");
            if ($(this).attr("aria-sort") === undefined) {
                $(this).attr("aria-sort", "ascending");
                $(this).addClass("sorting_asc");
            } else {
                if ($(this).attr("aria-sort") === "descending") {
                    $(this).attr("aria-sort", "ascending");
                    $(this).addClass("sorting_asc");
                } else {
                    $(this).attr("aria-sort", "descending");
                    $(this).addClass("sorting_desc");
                }
            }
            getData();
        });

        $('body').delegate('#standings_filter input', 'input', function () {
            getData()
        })

        var dt = null;

        function getUrlVars() {
            var vars = [], hash;
            var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
            for (var i = 0; i < hashes.length; i++) {
                hash = hashes[i].split('=');
                vars.push(hash[0]);
                vars[hash[0]] = hash[1];
            }
            return vars;
        }

        function initByRequest() {
            var params = getUrlVars();
            var _series = params["series_id"];
            if (_series != null && _series != "") {
                $("#series").val(_series);
                getData()
            }
        }

        initByRequest();

        function getData() {
            var gender = $('#gender').val() ? $('#gender').val() : '',
            series = $('#series').val() ? $('#series').val() : '',
            standings_length = $('#standings_length select').val() ? $('#standings_length select').val() : 10,
            standings_filter = $('#standings_filter input').val() ? $('#standings_filter input').val() : '',
            standings_sort = $('th[data-sort].sorting_asc, th[data-sort].sorting_desc').attr('data-sort') ? $('th[data-sort].sorting_asc, th[data-sort].sorting_desc').attr('data-sort') : 'Pts',
            standings_sort_direction = ($('th[data-sort].sorting_asc').length > 0) ? 'ASC' : 'DESC',
            standings_page = $('a.paginate_button.current').attr('data-dt-idx') ? $('a.paginate_button.current').attr('data-dt-idx') : 2;

            var url = '/series/results_array2.asp?series_id=' + series + '&gender=' + gender + '&standings_filter=' + standings_filter 
                    + '&standings_length=' + standings_length + '&standings_sort_direction=' + standings_sort_direction + '&standings_sort=' 
                    + standings_sort;

            if (dt) {
                dt.fnSettings().sAjaxSource = url;
                dt.dataTable().fnDraw();
            }
            else {
                var settings = {
                    bServerSide: true,
                    sAjaxSource: url,
                    pagingType: "full_numbers",
                    "lengthMenu": [[10, 25, 50, 100, -1], [10, 25, 50, 100, 'All']],
                    fnServerData: function (src, data, cb) {
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
    <div class="row">
        <div class="col-xs-6">
            <img src="/graphics/html_header.png" alt="Series Header" class="img-responsive">
        </div>
        <div class="col-xs-6">
            <h2 class="h2">Series Standings: Age Neutral</h2>
        </div>
    </div>

    <!--#include file = "series_nav.asp" -->

    <div class="row bg-warning">
        <div class="col-xs-1"><label for="series">Series:</label></div>
        <div class="col-xs-3">
            <select class="form-control" name="series" id="series">
                <option value="">&nbsp;</option>
                <%For i = 0 To UBound(Series, 2) - 1%>
                    <option value="<%=Series(0, i)%>"><%=Series(1, i)%></option>
                <%Next%>
            </select>
        </div>

        <div class="col-xs-1"><label for="gender">Gender:</label></div>
        <div class="col-xs-2">
            <select class="form-control" name="gender" id="gender">
                <option value="M">Male</option>
                <option value="F">Female</option>
            </select>
        </div>

        <div class="col-xs-5">
            &nbsp;
        </div>
    </div>
    <br>
    <table id="standings" class="display" cellspacing ="0">
        <thead>
            <tr>
                <th>Pl</th>
                <th data-sort="PartName">Name</th>
                <th data-sort="Age">Age</th>
                <th data-sort="AgeScore">Pts</th>
            </tr>
        </thead>
        <tfoot>
            <tr>
                <th>PL</th>
                <th>Name</th>
                <th>Age</th>
                <th>Pts</th>
            </tr>
        </tfoot>                        
    </table>
</div>
</body>
</html>
