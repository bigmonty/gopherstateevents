<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID
Dim sEventName
Dim dEventDate

lEventID = Request.QueryString("event_id")
If CStr(lEventID) = vbNullString Then lEventID = 0
If Not IsNumeric(lEventID) Then Response.Redirect("http://www.google.com")
If CLng(lEventID) < 0 Then Response.Redirect("http://www.google.com")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = Replace(rs(0).Value, "''", "'")
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html>
<head>
<title>Scrolling Results For <%=sEventName%></title>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
<meta name="description" content="Scrolling Results from Gopher State Events">

<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.min.css">
<link rel="alternate" href="http://gopherstateevents.com" hreflang="en-us" />
<link rel="stylesheet" href="https://code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css">
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap-theme.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-submenu/3.0.1/css/bootstrap-submenu.min.css">


<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
<script src="https://cdn.jsdelivr.net/jquery.marquee/1.3.1/jquery.marquee.min.js"></script>
<script src="https://code.jquery.com/jquery-2.1.4.min.js"></script>
<script src="https://code.jquery.com/ui/1.11.4/jquery-ui.js"></script>

<!--Data Table references-->   
<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css">
<script src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>

<script>
$(document).ready(function() {
    $('#results').DataTable( {
        "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "All"]],
        "order": [[ 0, 'asc' ]],
        "columnDefs": [
	        {
		        "targets": [6,7,8],
		        "orderable": false
	       }],
         "ajax": {"url":"results_scroll_source.asp?event_id=<%=lEventID%>"}
    } );
} );

var JSONdata = [
{
    "Pl": "1","Bib": "106","First Name":"Chase","Last Name": "Cayo","MFX": "M","Age": "22","Chip Time": "14:46.3","Gun Time": "14:46.9","Start Time":"0:00.6","City": "Otsego","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    },
    {
    "Pl": "2","Bib": "66","First Name":"Brendan","Last Name": "Sage","MFX": "M","Age": "22","Chip Time": "15:56.6","Gun Time": "15:56.8","Start Time":"0:00.2","City": "Saint Michael","St": "MN"
    },
    {
    "Pl": "3","Bib": "36","First Name":"Isaac","Last Name": "Basten","MFX": "M","Age": "17","Chip Time": "16:19.2","Gun Time": "16:21.5","Start Time":"0:02.3","City": "Buffalo","St": "MN"
    },
    {
    "Pl": "4","Bib": "304","First Name":"nick","Last Name": "oak","MFX": "M","Age": "17","Chip Time": "16:21.7","Gun Time": "16:22.5","Start Time":"0:00.8","City": "buffalo","St": "MN"
    }
];

$('#results').DataTable( {
    data: JSONdata,
columns:[    
{ data: "Pl",        title: "Pl" },         
{ data: "Bib",       title: "Bib" },        
{ data: "First Name",title: "First Name" },  
{ data: "Last Name", title: "Last Name" },  
{ data: "MFX",       title: "MFX" },        
{ data: "Age",       title: "Age" },        
{ data: "Chip Time", title: "Chip Time" },  
{ data: "Gun Time",  title: "Gun Time" },   
{ data: "Start Time",title: "Start Time" },  
{ data: "City",      title: "City" },       
{ data: "St",        title: "St" }          
],
    paging: false
 });

var $el = $(".table-responsive");
function anim() {
  var st = $el.scrollTop();
  var sb = $el.prop("scrollHeight")-$el.innerHeight();
  $el.animate({scrollTop: st<sb/2 ? sb : 0}, 5000, anim);
}
function stop(){
  $el.stop();
}
anim();
$el.hover(stop, anim);
</script>
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-sm-6">
            <img class="img-responsive" src="/graphics/html_header.png" alt="Scrolling Results">
        </div>
        <div class="col-sm-6" style="text-align:center;">
            <h3 class="h3">Scrolling Results For <%=sEventName%></h3>
            <h4 class="h4"><%=dEventDate%></h4>
        </div>
    </div>

    <div style="padding:5px 0 5px 0;">
        <a href="/results/fitness_events/results_scrollx2.asp?event_id=<%=lEventID%>">Refresh Results</a>
    </div>

    <div class="table-responsive" style="width:100%;height: 700px; overflow: auto;">
        <table id="results" class="display" cellspacing ="0" width="100%">
            <thead>
                <tr>
                    <th>Pl</th>
                    <th>Bib</th>
                    <th>First Name</th>
                    <th>Last Name</th>
                    <th>MFX</th>
                    <th>Age</th>
                    <th>Chip Time</th>
                    <th>Gun Time</th>
                    <th>Start Time</th>
                    <th>City</th>
                    <th>St</th>
                </tr>
            </thead>
            <tfoot>
                <tr>
                    <th>Pl</th>
                    <th>Bib</th>
                    <th>First Name</th>
                    <th>Last Name</th>
                    <th>MFX</th>
                    <th>Age</th>
                    <th>Chip Time</th>
                    <th>Gun Time</th>
                    <th>Start Time</th>
                    <th>City</th>
                    <th>St</th>
                </tr>
            </tfoot>                        
        </table>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>