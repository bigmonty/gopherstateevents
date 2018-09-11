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
<title>GSE&copy; Broken Marathon</title>
<meta name="description" content="A virtual race combining two separate race times using age-grading.">
<!--#include file = "../includes/js.asp" -->
</head>

<body>
<div class="container">
    <div class="row">
        <div class="col-md-8">
	        <img src="/graphics/broken_marathon_2017.png" alt="Broken Marathon" class="img-responsive">
        </div>
        <div class="col-md-4">
		    <h1 class="h1">Gopher State Events "Broken Marathon"</h1>

            <p class="bg-success">
                The Gopher State Events Broken Marathon is a virtual race where age graded performances in the fastest two times of three separate events 
                are combined for overall order of finish.  In essence, it is a "marathon" that was done over two different days in two different races.
                See guidelines below:
            </p>

            <h4 class="h4"><a href="http://www.tempotickets.com/BrokenMarathon17">Register Here</a></h4>
            <p class="bg-danger text-danger"  style="font-weight: bold;font-style: italic;">A PRIZE MONEY RACE!  Prizes are determined by AGE GRADED PERFORMANCES!</p>
        </div>
    </div>
	
    <ul class="list-group">
        <li class="list-group-item list-group-item-warning">
            All participants’ scores would be age graded.  Age grading is a world-accepted method of comparing the “quality” of 
            road race performances across gender and age.  Note that placing is based on a scale factor, not an adjusted time.  Here are the age-graded
            performances of last year's New Prague, Mora and Gandy Dancer races:
            <ul>
                <li><a href="http://www.gopherstateevents.com/results/fitness_events/age_graded.asp?event_id=498&race_id=831">New Prague</a></li>
                <li><a href="http://www.gopherstateevents.com/results/fitness_events/age_graded.asp?event_id=528&race_id=898">Mora</a></li>
                <li><a href="http://www.gopherstateevents.com/results/fitness_events/age_graded.asp?event_id=546&race_id=937">Gandy Dancer</a></li>
            </ul>
        </li>
        <li class="list-group-item list-group-item-danger">
            Prize money will be awarded to the top 10% of the finishers up to a maximum of 10 awards.  First prize is $50 and it goes down from there 
            proportionally to the size of the field.  For instance, a 50 runner field would award 5 cash prizes ($50, $40, $30, $20, and $10).   As 
            per age grading principles, these are mixed gender.
        </li>
        <li class="list-group-item list-group-item-info">
            All registered finishers OF ALL THREE RACES would receive a Gopher State Events shoulder bag.  These would be handed out at the conclusion of 
            the second event.  It would be up to the participant to pick theirs up…none will be mailed.
        </li>
        <li class="list-group-item list-group-item-success">
            There is a $5 entry fee to participate in this virtual event.  This fee is completely separate from the entry fee for the races  
            involved in this virtual event.  You may register <a href="http://www.tempotickets.com/BrokenMarathon17"><b>here</b></a>.
        </li>
        <li class="list-group-item list-group-item-warning">
            The organizing bodys of Run New Prague, The Mora Half-Marathon and The Gandy Dancer Half Marathon are in no way responsible for this event.  
            They have graciously consented to let Gopher State Events coordinate with them on this unique experience.
        </li>
        <li class="list-group-item list-group-item-danger">
            You may register for the three included races below:
            <ul>
                <li><a href="http://www.runnewprague.com/">New Prague</a></li>
                <li><a href="http://www.morahalfmarathon.com/">Mora</a></li>
                <li><a href="http://www.gandymarathon.com/">Gandy Dancer</a></li>
            </ul>
        </li>
    </ul>
	<!--#include file = "../includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
