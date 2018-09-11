<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html>
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>How a GSE&copy; Series Works</title>
<meta name="description" content="How the Gopher State Events (GSE) series works.">
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<img src="/graphics/html_header.png" alt="Series Header">
<div class="container">
    <h2 class="h2">How a GSE Series Works</h2>

    <p>GSE series is just a virtual affiliation between 
    multiple events that we manage.  It can be comprised of events hosted by a specific organization,
    events that are in the same general geographic location, events with the same theme, etc.</p>

    <p>Some series have no swag for participants, some have a little something special for those who participate in all events in the series, and some
    reward overall and age group achievement by gender and age.  These types of "extras" are determined solely by the event director(s).  All GSE does 
    is compile the data and post the results.</p>

    <p>From a participant standpoint, there is nothing for you to do except register, show up and finish.  If you feel the data regarding your 
    participation in a series is incorrect, please <a href="mailto:bob.schneider@gopherstateevents.com">contact us</a> and we will look into it.</p>

    <p><span style="font-weight: bold;">About Our Algorithm:</span>  First place in each gender or age category is awarded 100 points.  Points for subsequent finishers are awarded based on
    their place as compared to the size of the field in that gender/category.  Doing this allows races of different distances, modalities, and sizes
    to be in the same series.</p>

    <p>An interesting side effect of this is that you could be ahead of someone in your age group in the overall standings but behind them in the actual 
    age group standings simply because there are fewer people in the age groups from race to race so the point differentials are bigger. </p>
     
    <p>Series statistics are kept for overall male and female and in the following age groupings for both genders (NOTE-your age grouping
    is determined by your age on the date of the first race of the series):</p>

    <ul style="margin-left: 15px;">
        <li>14 and Under</li>
        <li>15 - 19</li>
        <li>20 - 24</li>
        <li>25 - 29</li>
        <li>30 - 34</li>
        <li>35 - 39</li>
        <li>40 - 44</li>
        <li>55 - 59</li>
        <li>60 - 64</li>
        <li>65 - 69</li>
        <li>70 & Over</li>
    </ul>

    <p>Final Note:  Since we at GSE take race-by-race registrations independently of different events, you may have to use your "My GSE History" account
    to "claim" race results that you participated in.</p>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
