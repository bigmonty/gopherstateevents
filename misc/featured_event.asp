<%@ Language=VBScript %>
<%
Option Explicit

Dim rs, sql, conn
Dim sClickPage
Dim FeaturedEvent(6)
Dim bHasFeatured

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FeaturedEventsID, EventName, EventDate, Location, WebURL, Descr, BlockImage, Views, "
sql = sql & "RAND(CAST(NEWID() AS VARBINARY)) * ( DateDiff( day, getDate(), EventDate)) AS Weight "
sql = sql & "FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date & "' AND '" & Date + 360 & "') AND "
sql = sql & "Active = 'y' ORDER BY Weight ASC"
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    FeaturedEvent(0) = rs(0).Value
    FeaturedEvent(1) = Replace(rs(1).Value, "''", "'")
    FeaturedEvent(2) = rs(2).Value
    FeaturedEvent(3) = Replace(rs(3).Value, "''", "'")
    FeaturedEvent(4) = rs(4).Value
    FeaturedEvent(5) = Replace(rs(5).Value, "''", "'")
    FeaturedEvent(6) = rs(6).Value
    rs(7).Value = CLng(rs(7).Value) + 1
    rs.Update
    bHasFeatured = True
Else
    bHasFeatured = False
End If
rs.Close
Set rs = Nothing

FeaturedEvent(4) = Replace(FeaturedEvent(4), "http://", "")
FeaturedEvent(4) = "http://" & FeaturedEvent(4)
%>
<!--#include file = "../includes/clean_input.asp" -->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Gopher State Events Featured Event</title>
<meta name="description" content="A featured event by Gopher State Events.">
<!--#include file = "../includes/js.asp" --> 

<script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
<script>
  (adsbygoogle = window.adsbygoogle || []).push({
    google_ad_client: "ca-pub-1381996757332572",
    enable_page_level_ads: true
  });
</script>
</head>

<body>
<div class="container">
    <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
    <!-- GSE Banner Ad -->
    <ins class="adsbygoogle"
            style="display:inline-block;width:728px;height:90px"
            data-ad-client="ca-pub-1381996757332572"
            data-ad-slot="1411231449"></ins>
    <script>
    (adsbygoogle = window.adsbygoogle || []).push({});
    </script>

    <%If bHasFeatured = True Then%>
    <h2 class="h2 bg-success">Gopher State Events Featured Event (<a href="/misc/featured_events.asp">Add Your Event</a>)</h2>
        <div class="col-md-6">
            <ul>
                <li><%=FeaturedEvent(1)%></li>
                <li><%=FeaturedEvent(2)%></li>
                <li><%=FeaturedEvent(3)%></li>
                <li><a class="text-danger" href="/featured_events/featured_clicks.asp?featured_events_id=<%=FeaturedEvent(0)%>&amp;click_page=<%=sClickPage%>" onclick="openThis(this.href,1024,768);return false;">Website</a></li>
            </ul>
            <p><%=FeaturedEvent(5)%></p>
        </div>
        <div class="col-md-6">
            <img class="img-responsive" src="/featured_events/images/<%=FeaturedEvent(6)%>" alt="<%=FeaturedEvent(1)%>">
        </div>
    <%End If%>
 </div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
