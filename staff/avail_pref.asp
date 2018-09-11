<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim sSummSatAM, sSummSatPM, sSummSunAM, sSummSunPM, sSummWkdyAM, sSummWkdyPM, sSummCmnts
Dim sSchlSatAM, sSchlSatPM, sSchlSunAM, sSchlSunPM, sSchlWkdyAM, sSchlWkdyPM, sSchlCmnts
Dim bFound

If Not Session("role") = "staff" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

bFound = False
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT StaffID FROM StaffAvailPref WHERE StaffID = " & Session("staff_id")
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then bFound = True
rs.Close
Set rs = Nothing

If bFound = False Then
    sql = "INSERT INTO StaffAvailPref (StaffID) VALUES (" & Session("staff_id") & ")"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If

If Request.Form.Item("submit_this") = "submit_this" Then
    If Request.Form.Item("summ_sat_am") = "on" Then sSummSatAM = "y"
    If Request.Form.Item("summ_sat_pm") = "on" Then sSummSatPM = "y"
    If Request.Form.Item("summ_sun_am") = "on" Then sSummSunAM = "y"
    If Request.Form.Item("summ_sun_pm") = "on" Then sSummSunPM = "y"
    If Request.Form.Item("summ_wkdy_am") = "on" Then sSummWkdyAM = "y"
    If Request.Form.Item("summ_wkdy_pm") = "on" Then sSummWkdyPM = "y"

    If Request.Form.Item("summ_sat_am_na") = "on" Then sSummSatAM = "n"
    If Request.Form.Item("summ_sat_pm_na") = "on" Then sSummSatPM = "n"
    If Request.Form.Item("summ_sun_am_na") = "on" Then sSummSunAM = "n"
    If Request.Form.Item("summ_sun_pm_na") = "on" Then sSummSunPM = "n"
    If Request.Form.Item("summ_wkdy_am_na") = "on" Then sSummWkdyAM = "n"
    If Request.Form.Item("summ_wkdy_pm_na") = "on" Then sSummWkdyPM = "n"

    If Request.Form.Item("schl_sat_am") = "on" Then sSchlSatAM = "y"
    If Request.Form.Item("schl_sat_pm") = "on" Then sSchlSatPM = "y"
    If Request.Form.Item("schl_sun_am") = "on" Then sSchlSunAM = "y"
    If Request.Form.Item("schl_sun_pm") = "on" Then sSchlSunPM = "y"
    If Request.Form.Item("schl_wkdy_am") = "on" Then sSchlWkdyAM = "y"
    If Request.Form.Item("schl_wkdy_pm") = "on" Then sSchlWkdyPM = "y"

    If Request.Form.Item("schl_sat_am_na") = "on" Then sSchlSatAM = "n"
    If Request.Form.Item("schl_sat_pm_na") = "on" Then sSchlSatPM = "n"
    If Request.Form.Item("schl_sun_am_na") = "on" Then sSchlSunAM = "n"
    If Request.Form.Item("schl_sun_pm_na") = "on" Then sSchlSunPM = "n"
    If Request.Form.Item("schl_wkdy_am_na") = "on" Then sSchlWkdyAM = "n"
    If Request.Form.Item("schl_wkdy_pm_na") = "on" Then sSchlWkdyPM = "n"

    If Not Request.Form.Item("summ_cmnts") = vbNullString Then sSummCmnts = Replace(Request.Form.Item("summ_cmnts"), "'", "''")
    If Not Request.Form.Item("schl_cmnts") = vbNullString Then sSchlCmnts = Replace(Request.Form.Item("schl_cmnts"), "'", "''")

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT SummSatAM, SummSatPM, SummSunAM, SummSunPM, SummWkdyAM, SummWkdyPM, SchlSatAM, SchlSatPM, SchlSunAM, SchlSunPM, SchlWkdyAM, SchlWkdyPM, "
    sql = sql & "SummCmnts, SchlCmnts FROM StaffAvailPref WHERE StaffID = " & Session("staff_id")
    rs.Open sql, conn, 1, 2
    rs(0).Value = sSummSatAM
    rs(1).Value = sSummSatPM
    rs(2).Value = sSummSunAM
    rs(3).Value = sSummSunPM
    rs(4).Value = sSummWkdyAM
    rs(5).Value = sSummWkdyPM
    rs(6).Value = sSchlSatAM
    rs(7).Value = sSchlSatPM
    rs(8).Value = sSchlSunAM
    rs(9).Value = sSchlSunPM
    rs(10).Value = sSchlWkdyAM
    rs(11).Value = sSchlWkdyPM
    rs(12).Value = sSummCmnts
    rs(13).Value = sSchlCmnts
    rs.Update
    rs.Close
    Set rs = Nothing
End If

sSummSatAM = vbNullString
sSummSatPM = vbNullString
sSummSunAM = vbNullString
sSummSunPM = vbNullString
sSummWkdyAM = vbNullString
sSummWkdyPM = vbNullString
sSchlSatAM = vbNullString
sSchlSatPM = vbNullString
sSchlSunAM = vbNullString
sSchlSunPM = vbNullString
sSchlWkdyAM = vbNullString
sSchlWkdyPM = vbNullString
sSummCmnts = vbNullString
sSchlCmnts = vbNullString

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SummSatAM, SummSatPM, SummSunAM, SummSunPM, SummWkdyAM, SummWkdyPM, SchlSatAM, SchlSatPM, SchlSunAM, SchlSunPM, SchlWkdyAM, SchlWkdyPM, "
sql = sql & "SummCmnts, SchlCmnts FROM StaffAvailPref WHERE StaffID = " & Session("staff_id")
rs.Open sql, conn, 1, 2
sSummSatAM = rs(0).Value
sSummSatPM = rs(1).Value
sSummSunAM = rs(2).Value
sSummSunPM = rs(3).Value
sSummWkdyAM = rs(4).Value
sSummWkdyPM = rs(5).Value
sSchlSatAM = rs(6).Value
sSchlSatPM = rs(7).Value
sSchlSunAM = rs(8).Value
sSchlSunPM = rs(9).Value
sSchlWkdyAM = rs(10).Value
sSchlWkdyPM = rs(11).Value
If Not rs(12).Value & "" = "" Then sSummCmnts = Replace(rs(12).Value, "''", "'")
If Not rs(13).Value & "" = "" Then sSchlCmnts = Replace(rs(13).Value, "''", "'")
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Staff Availability</title>
<meta name="description" content="Gopher State Events staff availability page.">
<!--#include file = "../includes/js.asp" --> 

<style type="text/css">
    p{text-indent: 0;margin: 5px 0 0 0;padding: 0;}
    td,th{padding-top: 5px;}
    h4{margin-top: 0;}
    hr{margin: 10px 0 0 0;}
    textarea{font-size: 1.25em;}
</style>
</head>

<body>
<div class="container">
	<!--#include file = "../includes/header.asp" -->
	
  	<div id="row">
		<!--#include file = "staff_menu.asp" -->
		<div class="col-md-10">
			<h3 class="h3">GSE Staff Availability Page</h3>
			
			<p>This page allows you to help us determine who to assign to which events.  Ultimately we would like staff members to assign themselves to
            available races but as we take on more opportunities we want to ensure that we have staff available.  Please check all that apply to 
            you from the options below.</p>

            <hr>

            <h4>Summer Availability</h4>
 
            <p>Note: A typical weekend am event in the summer requires us to be onsite by 6:30 or 7:00 and can last until 12:00 or 1:00 PM.  These times 
            do not include driving time, which can vary greatly.  The technician usually drives and is compensated for gas.</p>
            <form name="My_avail" method="post" action="avail_pref.asp">
            <table>
                <tr>
                    <td style="text-align: center;" colspan="2">
                        <input type="hidden" name="submit_this" id="submit_this" value="submit_this">
                        <input type="submit" name="submit1" id="submit1" value="Save Preferences">
                    </td>
                </tr>                
                <tr>
                    <th valign="top" colspan="2">I am often available (check all that apply):</th>
                </tr>
                <tr>
                    <td valign="top" colspan="2">
                        <%If sSummSatAM = "y" Then%>
                            <input type="checkbox" name="summ_sat_am" id="summ_sat_am" checked>&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sat_am" id="summ_sat_am">&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummSatPM = "y" Then%>
                            <input type="checkbox" name="summ_sat_pm" id="summ_sat_pm" checked>&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sat_pm" id="summ_sat_pm">&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummSunAM = "y" Then%>
                            <input type="checkbox" name="summ_sun_am" id="summ_sun_am" checked>&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sun_am" id="summ_sun_am">&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummSunPM = "y" Then%>
                            <input type="checkbox" name="summ_sun_pm" id="summ_sun_pm" checked>&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sun_pm" id="summ_sun_pm">&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummWkdyAM = "y" Then%>
                            <input type="checkbox" name="summ_wkdy_am" id="summ_wkdy_am" checked>&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_wkdy_am" id="summ_wkdy_am">&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummWkdyPM = "y" Then%>
                            <input type="checkbox" name="summ_wkdy_pm" id="summ_wkdy_pm" checked>&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_wkdy_pm" id="summ_wkdy_pm">&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                    </td>
                </tr>
                 <tr>
                    <th valign="top" colspan="2">I am NEVER available (check all that apply):</th>
                </tr>
                <tr>
                    <td valign="top" colspan="2">
                        <%If sSummSatAM = "n" Then%>
                            <input type="checkbox" name="summ_sat_am_na" id="summ_sat_am_na" checked>&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sat_am_na" id="summ_sat_am_na">&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummSatPM = "n" Then%>
                            <input type="checkbox" name="summ_sat_pm_na" id="summ_sat_pm_na" checked>&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sat_pm_na" id="summ_sat_pm_na">&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummSunAM = "n" Then%>
                            <input type="checkbox" name="summ_sun_am_na" id="summ_sun_am_na" checked>&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sun_am_na" id="summ_sun_am_na">&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummSunPM = "n" Then%>
                            <input type="checkbox" name="summ_sun_p_nam" id="summ_sun_pm_na" checked>&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_sun_pm_na" id="summ_sun_pm_na">&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummWkdyAM = "n" Then%>
                            <input type="checkbox" name="summ_wkdy_am_na" id="summ_wkdy_am_na" checked>&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_wkdy_am_na" id="summ_wkdy_am_na">&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSummWkdyPM = "n" Then%>
                            <input type="checkbox" name="summ_wkdy_pm_na" id="summ_wkdy_pm_na" checked>&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="summ_wkdy_pm_na" id="summ_wkdy_pm_na">&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                    </td>
                </tr>
                <tr>
                    <th valign="top">Comments:</th>                   
                    <td valign="top"><textarea name="summ_cmnts" id="summ_cmnts" rows="3" cols="100"><%=sSummCmnts%></textarea></td>                 
                </tr>
            </table>

            <hr>

            <h4>School Year Availability</h4>

            <p>Note: A typical weekday event in the school year requires us to be onsite by 2:30 or 3:00 and can last until 5:30 PM (Nordic Ski) or 6:30
            (Cross-Country Running).  A typical weekend am event in the school year requires us to be onsite by 8:30 or 9:00 and can last until 1:00 or 
            2:00 PM.  These times do not include driving time, which can vary greatly.  The technician usually drives and is compensated for gas.</p>

           <table>
                <tr>
                    <th valign="top" colspan="2">I am often available (check all that apply):</th>
                </tr>
               <tr>
                    <td valign="top" colspan="2">
                        <%If sSchlSatAM = "y" Then%>
                            <input type="checkbox" name="schl_sat_am" id="schl_sat_am" checked>&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sat_am" id="schl_sat_am">&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlSatPM = "y" Then%>
                            <input type="checkbox" name="schl_sat_pm" id="schl_sat_pm" checked>&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sat_pm" id="schl_sat_pm">&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlSunAM = "y" Then%>
                            <input type="checkbox" name="schl_sun_am" id="schl_sun_am" checked>&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sun_am" id="schl_sun_am">&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlSunPM = "y" Then%>
                            <input type="checkbox" name="schl_sun_pm" id="schl_sun_pm" checked>&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sun_pm" id="schl_sun_pm">&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlWkdyAM = "y" Then%>
                            <input type="checkbox" name="schl_wkdy_am" id="schl_wkdy_am" checked>&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_wkdy_am" id="schl_wkdy_am">&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlWkdyPM = "y" Then%>
                            <input type="checkbox" name="schl_wkdy_pm" id="schl_wkdy_pm" checked>&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_wkdy_pm" id="schl_wkdy_pm">&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                    </td>
                </tr>
                 <tr>
                    <th valign="top" colspan="2">I am NEVER available (check all that apply):</th>
                </tr>
                <tr>
                    <td valign="top" colspan="2">
                        <%If sSchlSatAM = "n" Then%>
                            <input type="checkbox" name="schl_sat_am_na" id="schl_sat_am_na" checked>&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sat_am_na" id="schl_sat_am_na">&nbsp;Sat AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlSatPM = "n" Then%>
                            <input type="checkbox" name="schl_sat_pm_na" id="schl_sat_pm_na" checked>&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sat_pm_na" id="schl_sat_pm_na">&nbsp;Sat PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlSunAM = "n" Then%>
                            <input type="checkbox" name="schl_sun_am_na" id="schl_sun_am_na" checked>&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sun_am_na" id="schl_sun_am_na">&nbsp;Sun AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlSunPM = "n" Then%>
                            <input type="checkbox" name="schl_sun_pm_na" id="schl_sun_pm_na" checked>&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_sun_pm_na" id="schl_sun_pm_na">&nbsp;Sun PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlWkdyAM = "n" Then%>
                            <input type="checkbox" name="schl_wkdy_am_na" id="schl_wkdy_am_na" checked>&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_wkdy_am_na" id="schl_wkdy_am_na">&nbsp;Weekday AM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                        <%If sSchlWkdyPM = "n" Then%>
                            <input type="checkbox" name="schl_wkdy_pm_na" id="schl_wkdy_pm_na" checked>&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%Else%>
                            <input type="checkbox" name="schl_wkdy_pm_na" id="schl_wkdy_pm_na">&nbsp;Weekday PM&nbsp;&nbsp;&nbsp;&nbsp;
                        <%End If%>
                    </td>
                </tr>
                <tr>
                    <th valign="top">Comments:</th>                   
                    <td valign="top"><textarea name="schl_cmnts" id="schl_cmnts" rows="3" cols="100"><%=sSchlCmnts%></textarea></td>                 
                </tr>
            </table>
            </form>
		</div>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>