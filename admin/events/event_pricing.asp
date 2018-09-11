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
<html lang="en">
<head>
<title>GSE&copy; Fitness Event Costs & Benefits</title>

<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="postinfo" content="/scripts/postinfo.asp">
<meta name="resource" content="document">
<meta name="distribution" content="global">

<meta name="description" content="Gopher State Events (GSE) is a conventional timing service for fitness events, cross-country, and nordic skiing offererd by H51 Software, LLC in Minnetonka, MN.">




<style type="text/css">
    td{
        text-align: left;
        padding-left: 5px;
    }
    
    th{
        text-align: right;
        padding-right: 5px;
    }
    
	p{
		font-size:0.8em;
        text-align: left;
	}
	
	ul{
		font-size:0.8em;
		margin-left:15px;
	}
	
	li{
		padding-top:5px;
        text-align: left;
	}
</style>
</head>

<body>
<div style="width: 750px;">
	<h2 style="margin:10px;padding:5px;background-color:#ececd8;">GSE Race Timing Fees & Services Provided</h2>
		
    <div style="float: left;width: 350px;">	
	    <h4 style="margin-left:10px;margin-top: 0;">What we bring to the table...</h4>

 		<ul style="margin:0;padding:0 0 0 30px;font-size: 0.9em;">
            <li><span>Pre-race reminder emails</span> sent to all participants (optional).</li>
            <li><span>Individual results emails</span> sent within minutes of finishing.</li>
            <li><span>Results Online</span> as the event unfolds.</li>
            <li>Accurate, awards-ready <span>results on demand.</span>.</li>
            <li><span>Finish Line Pix</span> online that day.</li>
            <li><span>Finish Line Videos</span> online within a week.</li>
            <li><span>Electronic results display</span> during the race.</li>
            <li><span>Electronic Announcer's Portal.</span></li>
            <li>Automatically scored <span>Race Series</span></li>
                <li><span>Team Scoring</span>.</li>
                <li><span>Race Splits</span>.</li>
        </ul>

        <h4 style="margin-top: 10px;">Pictures-Only Option</h4>

        <p>
            Don't want to divert funds from your smaller race to timing costs?  We have a plan for that.  We will bring a visible clock and a 
            camera and try to take a finish line picture of every finisher with the clock in the background indicating their time.  We put the images
            online and send you a link that you can share with your participants.  The fee for this utility is $300 + mileage.
        </p>
    </div>
    <div style="margin-left: 360px;">
	    <h4 style="margin:10px 0 0 10px;">Pricing:</h4>

        <table style="font-size: 0.9em;margin-left: 30px;">
            <tr>
                <th style="text-align: left;border-bottom: 1px sold #ccc;">
                    Race Distance
                    <br>
                    <span style="font-size: 0.8em;font-weight: normal;">(longest race not to exceed)</span>
                </th>
                <th style="white-space: nowrap;">Base Fee</th>
            </tr>
            <tr>
                <td class="alt">5K</td>
                <td class="alt">$500</td>
            </tr>
            <tr>
                <td>10K</td>
                <td>$550</td>
            </tr>
            <tr>
                <td class="alt" style="white-space: nowrap;">Half-Mar (& Multi-Sport)</td>
                <td class="alt">$750</td>
            </tr>
            <tr>
                <td>Marathon</td>
                <td>$850</td>
            </tr>
        </table>

	    <ul style="padding-left:5px;margin:10px 0 10px 30px;font-size: 0.8em;">
            <li>Per Participant: (custom bibs are extra)
                <ul style="font-size: 1.0em;">
                    <li>Single Disposable Chip on Tyvek Bib: $1.50</li>
                    <li>Single Disposable Chip on Handlebar Bib: $2.00</li>
                    <li>Nordic Ski Bibs (Event Staff Apply): $2.00</li>
                    <li>Nordic Ski Bibs (GSE Staff Apply): $2.50</li>
                    <li>Permanent Ankle Straps: $1.50 ($15 replacement fee)</li>
                </ul>
            </li>
		    <li>Mileage: $0.60/mile round trip from Minnetonka, MN.</li>
		    <li>Lodging: $150 (for events more than 150 miles from Excelsior, MN)</li>
            <li>Note: A non-refundable deposit must be submitted prior to reserving your event date.  This deposit will be
                applied to your invoice upon the event's conclusion.</li>
            <li>Optional Features:
                <ul style="font-size: 1.0em;">
                    <li>Digital Results Display: $150</li>
                    <li>Announcer Portal: $150</li>
                    <li>Team Scoring: $25</li>
                    <li>Splits: $25 + $150 for each additional timing station</li>
                    <li>Series: $10/race in the series</li>
                </ul>
            </li>
	    </ul>

        <p>
            <span style="font-weight: bold;">Small Race Pricing:</span> To increase the viability of smaller races we offer a $200 base fee plus 
            $5/participant (plus mileage).  This is usually more economical until a race reaches around 90 participants.  At that point, 
            our conventional pricing structure will be applied.
        </p>
    </div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
