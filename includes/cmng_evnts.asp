<%
Dim FeaturedEvents()
Dim sClickPage

sClickPage = Request.ServerVariables("URL")

i = 0
ReDim FeaturedEvents(3, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TOP 5 FeaturedEventsID, BlockImage, WebURL, Views, EventName FROM FeaturedEvents WHERE (EventDate BETWEEN '" & Date & "' AND '" & Date + 360 
sql = sql & "') AND Active = 'y' ORDER BY NewID()"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    FeaturedEvents(0, i) = rs(0).Value
    FeaturedEvents(1, i) = rs(1).Value
    FeaturedEvents(2, i) = rs(2).Value
    FeaturedEvents(3, i) = Replace(rs(4).Value, "''", "'")
    rs(3).Value = CLng(rs(3).Value) + 1
    i = i + 1
    ReDim Preserve FeaturedEvents(3, i)
    rs.Update
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

%>
		<div class="col-sm-2">
            <%If UBound(FeaturedEvents, 2) > 0 Then%>
                <h5 class="h5">Your Next Race?</h5>
                <%For i = 0 To UBound(FeaturedEvents, 2) - 1%>
                    <a href="/featured_events/featured_clicks.asp?featured_events_id=<%=FeaturedEvents(0, i)%>&amp;click_page=<%=sClickPage%>" onclick="openThis(this.href,1024,768);return false;">
                        <img src="/featured_events/images/<%=FeaturedEvents(1, i)%>" 
                        alt="<%=FeaturedEvents(3, i)%>" class="img-responsive" style="width:150px;">
                    </a>
                    <hr>
                <%Next%>
            <%End If%>
		</div>
