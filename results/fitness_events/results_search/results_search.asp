<%@ Language=VBScript%>

<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lRaceID
Dim sEventName, sRaceName
Dim dEventDate

lRaceID = Request.QueryString("race_id")
If CStr(lRaceID) = vbNullString Then lRaceID = 0
If Not IsNumeric(lRaceID) Then Response.Redirect "http://www.google.com"
If CLng(lRaceID) < 0 Then Response.Redirect "http://www.google.com"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT e.EventName, e.EventDate, rd.RaceName FROM Events e INNER JOIN RaceData rd ON e.EventID = rd.EventID WHERE rd.RaceID = " & lRaceID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
sRaceName = Replace(rs(2).Value, "''", "'")
rs.Close
Set rs = Nothing

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE (Gopher State Events) Results Search Utility</title>
     
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.0/jquery.min.js"></script>

<!-- bootstrap JavaScript & CSS -->
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
  
<!--Data Table references-->   
<link rel="stylesheet" href="https://cdn.datatables.net/1.10.15/css/dataTables.bootstrap.min.css">
<script src="https://cdn.datatables.net/1.10.15/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.10.15/js/dataTables.bootstrap.min.js"></script>

<script>
    $(document).ready(function () {
        $('body').delegate('#bib, #first_name, #last_name, #age, #gender select', 'change', function () {
            getData()
        });

        function GetURLParameter(sParam)
        {
            var sPageURL = window.location.search.substring(1);
            var sURLVariables = sPageURL.split('&');
            for (var i = 0; i < sURLVariables.length; i++)
            {
                var sParameterName = sURLVariables[i].split('=');
                if (sParameterName[0] == sParam)
                {
                    return sParameterName[1];
                }
            }
        }​

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

        initByRequest();

        $(document).ready(function(){
            $('#results').DataTable();
        });

        function getData() {
            var race_id  = GetURLParameter('race_id');gender = $('#gender').val() ? $('#gender').val() : '', age = $('#age').val() ? $('#age').val() : '', 
            first_name = $('#first_name').val ? $('#first_name').val() : '', last_name = $('#last_name').val ? $('#last_name').val() : '',
            results_length = $('#results_length select').val() ? $('#results_length select').val() : 10,
            results_sort = $('th[data-sort].sorting_asc, th[data-sort].sorting_desc').attr('data-sort') ? $('th[data-sort].sorting_asc, th[data-sort].sorting_desc').attr('data-sort') : 'Pts',
            results_sort_direction = ($('th[data-sort].sorting_asc').length > 0) ? 'ASC' : 'DESC',
            results_page = $('a.paginate_button.current').attr('data-dt-idx') ? $('a.paginate_button.current').attr('data-dt-idx') : 2;

            var url = 'results_array.asp?race_id=' + race_id + '&gender=' + gender + '&age=' + age + '&first_name=' + first_name
                    + '&last_name=' + last_name + 'results_length=' + results_length
                    + '&results_sort_direction=' + results_sort_direction + '&results_sort=' + results_sort;

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
                dt = $('#results').dataTable(settings);
            }
            };
     });
</script>
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-xs-6">
            <img src="/graphics/html_header.png" alt="Results Header" class="img-responsive">
        </div>
        <div class="col-xs-6">
            <h2 class="h2">Gopher State Events Results Search:  <%=sEventName%> on <%=dEventDate%></h2>
            <h3 class="h3">Race: <%=sRaceName%></h3>
        </div>
    </div>

    <div class="row bg-warning">
        <h4 class="h4">Search Criteria</h4>
        <div class="col-xs-1"><label for="first_name">First Name:</label></div>
        <div class="col-xs-2">
            <input type="text" class="form-control" name="first_name" id="first_name">
        </div>

        <div class="col-xs-1"><label for="last_name">Last Name:</label></div>
        <div class="col-xs-2">
            <input type="text" class="form-control" name="last_name" id="last_name">
        </div>

        <div class="col-xs-1"><label for="gender">Gender:</label></div>
        <div class="col-xs-2">
            <select class="form-control" name="gender" id="gender">
                <option value="B">Combined</option>
                <option value="M">Male</option>
                <option value="F">Female</option>
            </select>
        </div>

        <div class="col-xs-1"><label for="age">Age:</label></div>
        <div class="col-xs-2">
            <select class="form-control" name="age" id="age">
                <%For i = 0 to 99%>
                    <option value="<%=i%>"><%=i%></option>
                <%Next%>
            </select>
        </div>
    </div>

    <div class="row">
        <h4 class="h4 bg-success">Search Results</h4>
        <table id="results" class="display" cellspacing ="0">
            <thead>
                <tr>
                    <th data-sort="Pl">Pl</th>
                    <th data-sort="Bib">Bib</th>
                    <th data-sort="PartName">Name</th>
                    <th data-sort="Gender">Gndr</th>
                    <th data-sort="Age">Age</th>
                    <th data-sort="Chip">Chip Time</th>
                    <th data-sort="Gun">Gun Time</th>
                    <th data-sort="Start">Start Time</th>
                    <th data-sort="From">From</th>
                </tr>
            </thead>
            <tfoot>
                <tr>
                    <th>Pl</th>
                    <th>Bib</th>
                    <th>Name</th>
                    <th>Gndr</th>
                    <th>Age</th>
                    <th>Chip Time</th>
                    <th>Gun Time</th>
                    <th>Start Time</th>
                    <th>From</th>
                </tr>
            </tfoot>                        
        </table>
    </div>
</div>
</body>
</html>
