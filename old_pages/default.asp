<%@ Language=VBScript %>
<%
Option Explicit

Dim rs, sql, conn, conn2, rs2, sql2
Dim i
Dim iEventType
Dim sErrMsg, sUserName, sPassword, sRole, sSignOut, sLoginErr, sClickPage
Dim TodayEvents(), Testim(2), Series(), RaceReport(), FeaturedEvent(6)
Dim bNotFound, bHasFeatured

Response.Redirect "index.html"

sSignOut=Request.QueryString("sign_out")
If sSignOut = vbNullString Then sSignOut = "n"

If sSignOut = "y" Then Session.Contents.RemoveAll()

sClickPage = Request.ServerVariables("URL")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
				
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

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

i = 0
ReDim Series(1, 0)
sql = "SELECT SeriesID, SeriesName FROM Series WHERE SeriesYear = " & Year(Date) & " ORDER BY SeriesName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
    Series(0, i) = rs(0).Value
	Series(1, i) = Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Series(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

sql = "SELECT Testimonial, Author, Role FROM Testimonials ORDER BY NewID()"
Set rs = conn.Execute(sql)
Testim(0) = rs(0).Value
Testim(1) = rs(1).Value
Testim(2) = rs(2).Value
Set rs = Nothing

If Request.Form.Item("submit_login") = "submit_login" Then
	'see if this user has entered from the form correctly within the past 20 minutes
	sql = "SELECT AuthAccessID FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "' AND WhenHit >= '" & Now() - CSng(1/72) & "' AND Page = 'sign_in' ORDER BY AuthAccessID DESC"
	Set rs = conn.Execute(sql)
	If rs.RecordCount > 0 Then Session("access_sign_in") = "y"
    If rs.BOF and rs.EOF Then
        '--
    Else
        Session("access_sign_in") = "y"
    End If
	Set rs = Nothing

	If Session("access_sign_in") = "y" Then	'if they are an authorized user allow them to proceed
		Dim sHackMsg, sMsgText
		
		sUserName = CleanInput(Trim(Request.Form.Item("user_name")))
		If sHackMsg = vbNullString Then sPassword = CleanInput(Trim(Request.Form.Item("password")))
		
		If sHackMsg = vbNullString Then
            sRole = Request.Form.Item("role")

            Select Case sRole
                Case "staff"
				    sql = "SELECT StaffID, FirstName, LastName, Email FROM Staff WHERE UserID = '" & sUserName & "' AND Password = '" & sPassword & "' "
                    sql = sql & " AND Active = 'y'"
				    Set rs = conn.Execute(sql)
                    If rs.BOF and rs.EOF Then
                        bNotFound = True
                    Else
                        Session("role") = "staff"
                        Session("staff_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                        bNotFound = False
                    End If
				    Set rs = Nothing

                    'log this login and then redirect
                    If bNotFound = False Then
                        sql = "INSERT INTO StaffLogin (StaffID, WhenVisit, IPAddress, Browser) VALUES (" & Session("staff_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn.Execute(sql)
                        Set rs=Nothing

                        Response.Redirect "/staff/profile.asp"
                    Else
                        sLoginErr = "I am sorry.  Those credentials were not found for that role."
                    End If
                Case "admin"
			        If sUserName = "bobbabuoy" And sPassword = "Zeroto@123" Then 
                        Session("role") = "admin"
                        Session("my_name") = "Bob Schneider"
                        Session("my_email") = "bob.schneider@gopherstateevents.com"
                        bNotFound = False
			        ElseIf sUserName = "solveig" And sPassword = "colaianni" Then
                        Session("role") = "admin"
                        Session("my_name") = "Solveig Colianni"
                        Session("my_email") = "solveigkc@gmail.com"
                        bNotFound = False
                    Else
                        bNotFound = True
			        End If

                    If bNotFound = False Then
                        sql = "INSERT INTO AdminLogin (AdminName, WhenVisit, IPAddress, Browser) VALUES ('" & Session("my_name") & "', '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn.Execute(sql)
                        Set rs=Nothing

                        Response.Redirect "/admin/events/event_mgr.asp" 
                    Else
                        sLoginErr ="I am sorry.  Those credentials were not found for that role."
                    End If
                Case "coach"
                    Set rs = Server.CreateObject("ADODB.Recordset")
		            sql = "SELECT CoachesID, FirstName, LastName, Email FROM Coaches WHERE UserID = '" & sUserName & "' AND Password = '" & sPassword & "'"
				    rs.Open sql, conn2, 1, 2
                    If rs.RecordCount > 0 Then
			            Session("role") = "coach"
			            Session("my_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                        bNotFound = False
                    Else
                        bNotFound = True
                    End If
                    rs.Close
		            Set rs = Nothing

                    If bNotFound = False Then
                        sql = "INSERT INTO CoachLogin (CoachesID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn2.Execute(sql)
                        Set rs=Nothing
                       
                        Response.Redirect "/cc_meet/coach/meets/lineup_mgr.asp"
                    Else
                        'check for team staff login
		                sql = "SELECT TeamStaffID,  FirstName, LastName, Email, CoachesID FROM TeamStaff WHERE UserName = '" & sUserName 
                        sql = sql & "' AND Password = '" & sPassword & "'"
				        Set rs = conn2.Execute(sql)
                        If rs.BOF and rs.EOF Then
                            bNotFound = True
                        Else
			                Session("role") = "team_staff"
			                Session("my_id") = rs(0).Value
                            Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                            Session("my_email") = rs(3).Value
                            Session("team_coach_id") = rs(4).Value
                            bNotFound = False
                        End If
		                Set rs = Nothing

                        If bNotFound = False Then
                            sql = "INSERT INTO TeamStaffLogin (TeamStaffID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                            sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                            Set rs=conn2.Execute(sql)
                            Set rs=Nothing
                       
                            Response.Redirect "/cc_meet/coach/meets/lineup_mgr.asp"
                        End If
                    End If

                    If bNotFound = True Then sLoginErr ="I am sorry.  Those credentials were not found for that role."
                Case "meet_dir"
		            sql = "SELECT MeetDirID,  FirstName, LastName, Email FROM MeetDir WHERE UserID = '" & sUserName & "' AND Password = '" 
                    sql = sql & sPassword & "'"
                    Set rs = conn2.Execute(sql)
                    If rs.BOF and rs.EOF Then
                        bNotFound = True
                    Else
			            Session("role") = "meet_dir"
			            Session("my_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                    End If
		            Set rs = Nothing

                    If bNotFound = False Then
                        sql = "INSERT INTO MeetDirLogin (MeetDirID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn2.Execute(sql)
                        Set rs=Nothing

                        Response.Redirect "/cc_meet/meet_dir/meet_dir_home.asp"
                    Else
                        sLoginErr ="I am sorry.  Those credentials were not found for that role."
                    End If
                Case "event_dir"
		            sql = "SELECT EventDirID,  FirstName, LastName, Email FROM EventDir WHERE UserID = '" & sUserName & "' AND Password = '" & sPassword & "'"
                    Set rs = conn.Execute(sql)
                    If rs.BOF and rs.EOF Then
                        bNotFound = True
                    Else
			            Session("role") = "event_dir"
			            Session("my_id") = rs(0).Value
                        Session("my_name") = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
                        Session("my_email") = rs(3).Value
                    End If
		            Set rs = Nothing
 					
                    If bNotFound = False Then
                        sql = "INSERT INTO EventDirLogin (EventDirID, WhenVisit, IPAddress, Browser) VALUES (" & Session("my_id") & ", '" & Now() 
                        sql = sql & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Request.ServerVariables("HTTP_USER_AGENT") & "')"
                        Set rs=conn.Execute(sql)
                        Set rs=Nothing
                       
                        Response.Redirect "/events/event_dir/event_dir_home.asp"
                    Else
                        sLoginErr ="I am sorry.  Those credentials were not found for that role."
                    End If
            End Select
		End If
	End If
End If

'log this user if they are just entering the site
If Session("access_login") = vbNullString Then 
	sql = "INSERT INTO AuthAccess(WhenHit, IPAddress, Page) VALUES ('" & Now() & "', '" & Request.ServerVariables("REMOTE_ADDR") 
	sql = sql & "', 'sign_in')"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
Else
	sql = "DELETE FROM AuthAccess WHERE IPAddress = '" & Request.ServerVariables("REMOTE_ADDR")  & "' AND Page  = 'sign_in'"
	Set rs = conn.Execute(sql)
	Set rs = Nothing

	Session.Contents.Remove("access_sign_in")
End If

Dim conn3
Dim sThinkAbout						
Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"
sql = "SELECT Quote FROM Quotes ORDER BY NEWID()"
Set rs = conn3.Execute(sql)
sThinkAbout = Replace(rs(0).Value, "''", "'")
Set rs = Nothing
conn3.close
Set conn3 = Nothing

i = 0
ReDim TodayEvents(6, 0)
sql = "SELECT EventID, EventName, Logo, EventType, Website FROM Events WHERE EventDate = '" & Date & "'"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	TodayEvents(0, i) = rs(0).Value
	TodayEvents(1, i) = Replace(rs(1).Value, "''", "'") 
    TodayEvents(2, i) = rs(2).Value
    TodayEvents(3, i) = "fitness"
    TodayEvents(4, i) = rs(3).Value
    TodayEvents(5, i) = rs(4).Value
	i = i + 1
	ReDim Preserve TodayEvents(6, i)
	rs.MoveNext
Loop
Set rs = Nothing
	
'now get cc/nordic	
sql = "SELECT MeetsID, MeetName, Sport, Logo, Website, StartList FROM Meets WHERE ShowOnline = 'y' AND MeetDate = '" & Date & "'"
Set rs = conn2.Execute(sql)
Do While Not rs.EOF
	TodayEvents(0, i) = rs(0).Value
	TodayEvents(1, i) = rs(1).Value
    TodayEvents(2, i) = rs(3).Value
    TodayEvents(3, i) = rs(2).Value
    TodayEvents(5, i) = rs(4).Value
    TodayEvents(6, i) = rs(5).Value
	i = i + 1
	ReDim Preserve TodayEvents(6, i)
	rs.MoveNext
Loop
Set rs = Nothing
	
'now get cc/nordic	
i = 0
ReDim RaceReport(7, 0)
sql = "SELECT e.EventID, e.EventName, e.EventDate, e.Location, r.Weather, r.RaceReport, r.Gallery, e.EventType FROM Events e INNER JOIN RaceReport r "
sql = sql & "ON e.EventID = r.EventID WHERE e.EventDate < '" & Date & "' AND e.EventDate > '" & Date - 30 & "' ORDER BY EventDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	RaceReport(0, i) = rs(0).Value
	RaceReport(1, i) = Replace(rs(1).Value, "''", "'")
    RaceReport(2, i) = rs(2).Value
    RaceReport(3, i) = Replace(rs(3).Value, "''", "'")
    RaceReport(4, i) = rs(4).Value
    RaceReport(5, i) = rs(5).Value
    RaceReport(6, i) = rs(6).Value
    RaceReport(7, i) = rs(7).Value

	i = i + 1
	ReDim Preserve RaceReport(7, i)

    if i = 5 Then Exit Do

	rs.MoveNext
Loop
Set rs = Nothing

Private Function GetEmbedLink(lThisEvent)
    sql2 = "SELECT EmbedLink FROM RaceGallery WHERE EventID = " & lThisEvent
    Set rs2 = conn.Execute(sql2)
    If Not rs2.EOF = rs2.BOF Then GetEmbedLink = rs2(0).Value
    Set rs2 = Nothing
End Function
%>
<!--#include file = "includes/clean_input.asp" -->
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "includes/meta2.asp" -->
<title>Minnesota Race Timing, Chip Timing, RFID Timing incl 5k by Gopher State Events</title>
<meta name="description" content="Gopher State Events chip timing, rfid timing, and other race timing for road races, fat tire, multi, cross-country, snow shoe, and nordic skiing in Minnesota including 5k.">
<!--#include file = "includes/js.asp" --> 

<!--Twitter Summary Card-->
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:site" content="@gsetiming">
<meta name="twitter:creator" content="@gsetiming">
<meta name="twitter:title" content="Minnesota Chip Timing / RFID Timing by Gopher State Events, LLC including 5k">
<meta name="twitter:description" content="Gopher State Events chip timing, rfid timing, and other race timing for events of all types including 5k.">
<meta name="twitter:image" content="http://www.gopherstateevents.com/graphics/banner_sample_001.jpg">

<meta property="og:title" content="Gopher State Events" />
<meta property="og:type" content="website" />
<meta property="og:url" content="http://www.gopherstateevents.com" />
<meta property="og:image" content="http://www.gopherstateevents.com/graphics/vira_crest.gif" />
<meta property="og:site_name" content="Gopher State Events" />
<meta property="fb:admins" content="509779604" />

<script id="twitter-wjs" src="/graphics/widgets.js"></script>
<script id="facebook-jssdk" src="/graphics/all.js"></script>
<script async="" src="/graphics/analytics.js"></script>

<script>
    function chkFlds() {
        if (document.site_login.user_name.value == '' ||
        document.site_login.password.value == '' ||
        document.site_login.role.value == '') {
            alert('All fields are required!');
            return false;
        }
        else
            return true;
    }
</script>

<style type="text/css">.fb_hidden{position:absolute;top:-10000px;z-index:10001}.fb_invisible{display:none}.fb_reset{background:none;border:0;border-spacing:0;color:#000;cursor:auto;direction:ltr;font-family:"lucida grande", tahoma, verdana, arial, sans-serif;font-size:11px;font-style:normal;font-variant:normal;font-weight:normal;letter-spacing:normal;line-height:1;margin:0;overflow:visible;padding:0;text-align:left;text-decoration:none;text-indent:0;text-shadow:none;text-transform:none;visibility:visible;white-space:normal;word-spacing:normal}.fb_reset>div{overflow:hidden}.fb_link img{border:none}
.fb_dialog{background:rgba(82, 82, 82, .7);position:absolute;top:-10000px;z-index:10001}.fb_reset .fb_dialog_legacy{overflow:visible}.fb_dialog_advanced{padding:10px;-moz-border-radius:8px;-webkit-border-radius:8px;border-radius:8px}.fb_dialog_content{background:#fff;color:#333}.fb_dialog_close_icon{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/yq/r/IE9JII6Z1Ys.png) no-repeat scroll 0 0 transparent;_background-image:url(http://static.ak.fbcdn.net/rsrc.php/v2/yL/r/s816eWC-2sl.gif);cursor:pointer;display:block;height:15px;position:absolute;right:18px;top:17px;width:15px}.fb_dialog_mobile .fb_dialog_close_icon{top:5px;left:5px;right:auto}.fb_dialog_padding{background-color:transparent;position:absolute;width:1px;z-index:-1}.fb_dialog_close_icon:hover{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/yq/r/IE9JII6Z1Ys.png) no-repeat scroll 0 -15px transparent;_background-image:url(http://static.ak.fbcdn.net/rsrc.php/v2/yL/r/s816eWC-2sl.gif)}.fb_dialog_close_icon:active{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/yq/r/IE9JII6Z1Ys.png) no-repeat scroll 0 -30px transparent;_background-image:url(http://static.ak.fbcdn.net/rsrc.php/v2/yL/r/s816eWC-2sl.gif)}.fb_dialog_loader{background-color:#f6f7f8;border:1px solid #606060;font-size:24px;padding:20px}.fb_dialog_top_left,.fb_dialog_top_right,.fb_dialog_bottom_left,.fb_dialog_bottom_right{height:10px;width:10px;overflow:hidden;position:absolute}.fb_dialog_top_left{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/ye/r/8YeTNIlTZjm.png) no-repeat 0 0;left:-10px;top:-10px}.fb_dialog_top_right{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/ye/r/8YeTNIlTZjm.png) no-repeat 0 -10px;right:-10px;top:-10px}.fb_dialog_bottom_left{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/ye/r/8YeTNIlTZjm.png) no-repeat 0 -20px;bottom:-10px;left:-10px}.fb_dialog_bottom_right{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/ye/r/8YeTNIlTZjm.png) no-repeat 0 -30px;right:-10px;bottom:-10px}.fb_dialog_vert_left,.fb_dialog_vert_right,.fb_dialog_horiz_top,.fb_dialog_horiz_bottom{position:absolute;background:#525252;filter:alpha(opacity=70);opacity:.7}.fb_dialog_vert_left,.fb_dialog_vert_right{width:10px;height:100%}.fb_dialog_vert_left{margin-left:-10px}.fb_dialog_vert_right{right:0;margin-right:-10px}.fb_dialog_horiz_top,.fb_dialog_horiz_bottom{width:100%;height:10px}.fb_dialog_horiz_top{margin-top:-10px}.fb_dialog_horiz_bottom{bottom:0;margin-bottom:-10px}.fb_dialog_iframe{line-height:0}.fb_dialog_content .dialog_title{background:#6d84b4;border:1px solid #3a5795;color:#fff;font-size:14px;font-weight:bold;margin:0}.fb_dialog_content .dialog_title>span{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/yd/r/Cou7n-nqK52.gif) no-repeat 5px 50%;float:left;padding:5px 0 7px 26px}body.fb_hidden{-webkit-transform:none;height:100%;margin:0;overflow:visible;position:absolute;top:-10000px;left:0;width:100%}.fb_dialog.fb_dialog_mobile.loading{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/ya/r/3rhSv5V8j3o.gif) white no-repeat 50% 50%;min-height:100%;min-width:100%;overflow:hidden;position:absolute;top:0;z-index:10001}.fb_dialog.fb_dialog_mobile.loading.centered{max-height:590px;min-height:590px;max-width:500px;min-width:500px}#fb-root #fb_dialog_ipad_overlay{background:rgba(0, 0, 0, .45);position:absolute;left:0;top:0;width:100%;min-height:100%;z-index:10000}#fb-root #fb_dialog_ipad_overlay.hidden{display:none}.fb_dialog.fb_dialog_mobile.loading iframe{visibility:hidden}.fb_dialog_content .dialog_header{-webkit-box-shadow:white 0 1px 1px -1px inset;background:-webkit-gradient(linear, 0% 0%, 0% 100%, from(#738ABA), to(#2C4987));border-bottom:1px solid;border-color:#1d4088;color:#fff;font:14px Helvetica, sans-serif;font-weight:bold;text-overflow:ellipsis;text-shadow:rgba(0, 30, 84, .296875) 0 -1px 0;vertical-align:middle;white-space:nowrap}.fb_dialog_content .dialog_header table{-webkit-font-smoothing:subpixel-antialiased;height:43px;width:100%}.fb_dialog_content .dialog_header td.header_left{font-size:12px;padding-left:5px;vertical-align:middle;width:60px}.fb_dialog_content .dialog_header td.header_right{font-size:12px;padding-right:5px;vertical-align:middle;width:60px}.fb_dialog_content .touchable_button{background:-webkit-gradient(linear, 0% 0%, 0% 100%, from(#4966A6), color-stop(.5, #355492), to(#2A4887));border:1px solid #2f477a;-webkit-background-clip:padding-box;-webkit-border-radius:3px;-webkit-box-shadow:rgba(0, 0, 0, .117188) 0 1px 1px inset, rgba(255, 255, 255, .167969) 0 1px 0;display:inline-block;margin-top:3px;max-width:85px;line-height:18px;padding:4px 12px;position:relative}.fb_dialog_content .dialog_header .touchable_button input{border:none;background:none;color:#fff;font:12px Helvetica, sans-serif;font-weight:bold;margin:2px -12px;padding:2px 6px 3px 6px;text-shadow:rgba(0, 30, 84, .296875) 0 -1px 0}.fb_dialog_content .dialog_header .header_center{color:#fff;font-size:16px;font-weight:bold;line-height:18px;text-align:center;vertical-align:middle}.fb_dialog_content .dialog_content{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/y9/r/jKEcVPZFk-2.gif) no-repeat 50% 50%;border:1px solid #555;border-bottom:0;border-top:0;height:150px}.fb_dialog_content .dialog_footer{background:#f6f7f8;border:1px solid #555;border-top-color:#ccc;height:40px}#fb_dialog_loader_close{float:left}.fb_dialog.fb_dialog_mobile .fb_dialog_close_button{text-shadow:rgba(0, 30, 84, .296875) 0 -1px 0}.fb_dialog.fb_dialog_mobile .fb_dialog_close_icon{visibility:hidden}
.fb_iframe_widget{display:inline-block;position:relative}.fb_iframe_widget span{display:inline-block;position:relative;text-align:justify}.fb_iframe_widget iframe{position:absolute}.fb_iframe_widget_lift{z-index:1}.fb_hide_iframes iframe{position:relative;left:-10000px}.fb_iframe_widget_loader{position:relative;display:inline-block}.fb_iframe_widget_fluid{display:inline}.fb_iframe_widget_fluid span{width:100%}.fb_iframe_widget_loader iframe{min-height:32px;z-index:2;zoom:1}.fb_iframe_widget_loader .FB_Loader{background:url(http://static.ak.fbcdn.net/rsrc.php/v2/y9/r/jKEcVPZFk-2.gif) no-repeat;height:32px;width:32px;margin-left:-16px;position:absolute;left:50%;z-index:4}</style>

<script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
<script>
  (adsbygoogle = window.adsbygoogle || []).push({
    google_ad_client: "ca-pub-1381996757332572",
    enable_page_level_ads: true
  });
</script>
</head>

<body data-twttr-rendered="true">
<div class="container">
  	<!--#include file = "includes/header.asp" -->

    <div class="row">
 		<div class="col-md-4">
            <img class="img-responsive" src="/graphics/new_logo_g.jpeg" alt="GSE Logo" style="margin: 10px 0 0 0;">
            <%If UBound(TodayEvents, 2) > 0 Then%>
                 <div>
                    <h4 class="h4 text-danger">Today's Events</h4>
                    <table class="table-condensed">
                        <%For i = 0 To UBound(TodayEvents, 2) - 1%>
                            <tr>
                                <td style="text-align: center;padding: 0;" valign="middle">
                                    <h5 class="h5 text-danger bg-danger"><%=TodayEvents(1, i)%></h5>
                                    <%If TodayEvents(3, i) = "fitness" Then%>
								        <%If TodayEvents(2, i) & "" = "" Then%>
                                            <%If TodayEvents(5, i) & "" = "" Then%>
									            <a href="/events/raceware_events.asp?event_id=<%=TodayEvents(0, i)%>"  
                                                onclick="openThis(this.href,1024,768);return false;"><img src="/graphics/info.jpg" alt="Info" 
                                                    style="height: 30px;"></a>
                                            <%Else%>
									            <a href="<%=TodayEvents(5, i)%>"  
                                                onclick="openThis(this.href,1024,768);return false;"><img src="/graphics/info.jpg" alt="Info" 
                                                    style="height: 30px;"></a>
                                            <%End If%>
                                        <%Else%>
                                            <%If TodayEvents(5, i) & "" = "" Then%>
                                                <a href="/events/raceware_events.asp?event_type=<%=TodayEvents(4, i)%>&event_id=<%=TodayEvents(0, i)%>" onclick="openThis(this.href,1024,768);return false;">
                                                    <img src="/events/logos/<%=TodayEvents(2, i)%>" alt="<%=TodayEvents(1, i)%>" style="height:30px;margin: 0;">
                                                </a>
                                            <%Else%>
                                                <a href="<%=TodayEvents(5, i)%>" onclick="openThis(this.href,1024,768);return false;">
                                                    <img src="/events/logos/<%=TodayEvents(2, i)%>" alt="<%=TodayEvents(1, i)%>" style="height:30px;margin: 0;">
                                                </a>
                                            <%End If%>
                                       <%End If%>
                                        &nbsp;&nbsp; &nbsp;&nbsp;
                                        <a href="/results/fitness_events/results.asp?event_type=<%=TodayEvents(4, i)%>&event_id=<%=TodayEvents(0, i)%>&first_rcd=1">
                                            <img src="/graphics/race_results.jpg" alt="Race Results" style="height: 30px;">
                                        </a>
                                    <%Else%>
								        <%If TodayEvents(2, i) & "" = "" Then%>
                                            <%If TodayEvents(5, i) & "" = "" Then%>
                                                <a href="/events/ccmeet_info.asp?meet_id=<%=TodayEvents(0, i)%>"  
                                                    onclick="openThis(this.href,1024,768);return false;"><img src="/graphics/info.jpg" alt="Info" 
                                                    style="height: 30px;"></a>
                                            <%Else%>
                                                <a href="<%=TodayEvents(5, i)%>"  
                                                    onclick="openThis(this.href,1024,768);return false;"><img src="/graphics/info.jpg" alt="Info" 
                                                    style="height: 30px;"></a>
                                            <%End If%>
                                            &nbsp;&nbsp; &nbsp;&nbsp;
					                        <%If Not TodayEvents(6, i) = vbNullString Then%>
						                        <a href="/ccmeet_admin/manage_meet/run_order/<%=TodayEvents(6, i)%>" target="_blank">
                                                    <img src="http://www.gopherstateevents.com/graphics/social_media/list.png" alt="View" style="height: 30px;">
                                                </a>
					                        <%End If%>
                                        <%Else%>
                                            <%If TodayEvents(5, i) & "" = "" Then%>
                                                <a href="/events/ccmeet_info.asp?meet_id=<%=TodayEvents(0, i)%>" onclick="openThis(this.href,1024,768);return false;">
                                                    <img src="/events/logos/<%=TodayEvents(2, i)%>" alt="<%=TodayEvents(1, i)%>" style="height:50px;margin: 0;">
                                                </a>
                                            <%Else%>
                                                <a href="<%=TodayEvents(5, i)%>" onclick="openThis(this.href,1024,768);return false;">
                                                    <img src="/events/logos/<%=TodayEvents(2, i)%>" alt="<%=TodayEvents(1, i)%>" style="height:50px;margin: 0;">
                                                </a>
                                            <%End If%>
                                            &nbsp;&nbsp; &nbsp;&nbsp;
					                        <%If Not TodayEvents(6, i) = vbNullString Then%>
						                        <a href="/ccmeet_admin/manage_meet/run_order/<%=TodayEvents(6, i)%>" target="_blank">
                                                    <img src="http://www.gopherstateevents.com/graphics/social_media/list.png" alt="View" style="height: 30px;">
                                                </a>
					                        <%End If%>
                                       <%End If%>
                                        &nbsp;&nbsp; &nbsp;&nbsp;
                                        <a href="/results/cc_rslts/cc_rslts.asp?meet_id=<%=TodayEvents(0, i)%>&sport=<%=TodayEvents(3, i)%>&rslts_page=overall_rslts.asp">
                                            <img src="/graphics/race_results.jpg" alt="Race Results" style="height: 50px;">
                                        </a>
                                    <%End If%>
                                    <hr style="margin: 0;">
                                </td>
                            </tr>
                        <%Next%>
                    </table>
                </div>
            <%End If%>
            <div>
		        <h4 class="h4">What Others Have Said</h4>
                <img class="img-responsive" src="/graphics/testimonials.jpg" style="width: 100px;float: right;">
                <p><em><%=Testim(0)%></em></p>
                <p><%=Testim(1) %><br><%=Testim(2)%>&nbsp;&nbsp;<a href="http://www.gopherstateevents.com/about/testim.asp">See more</a></p>
			</div>
            <div> 
                <h4 class="h4">Sample Finish Line Picture</h4>
                <img src="/graphics/banner_sample_001.jpg" class="img-responsive" alt="Image Sample">
                <p>...Free of charge to all finishers.</p>
			</div>
             <div>
                <h4 class="h4">No Timing Necessary?</h4>
                <p>
                    <img class="img-responsive" src="/graphics/gopher.jpg" style="float:left;height: 100px;" alt="Gopher">
                    ..."just for fun"...to raise money for a good cause...little emphasis on times, places and awards?<br><br>
                    We will bring a clock, a camera, and a friendly face.  Pix online later that day 
                    FOR FREE DOWNLOAD with the clock in the background.  No timing equipment, no chips, no data entry, and far less cost.  Call us
                    at <a href="tel:+1-612-720-8427">612-720-8427</a> or <a href="mailto:bob.schneider@gopherstateevents.com" style="font-weight: bold;">email us.</a>
                </p>
			</div>
            <div class="bg-success">
                <h4 class="h4">The GSE Mission</h4>
                <p>
                    To make fitness event management, including related school sports, a more enjoyable, healthful, and social experience for all participants,
                    to support the worthy causes that these events contribute to, to make the work easier for the event director, and to do this at a fee that 
                    is affordable and margin-based, rather than driven by what we think the market will bear.
                </p>
			</div>
		</div>
        <div class="col-md-4">
             <div class="bg-success">
                <h4 class="h4"><a href="http://www.gopherstateevents.com/misc/value_adds.asp" onclick="openThis(this.href,1024,768);return false;"
                   style="text-decoration: none;">Value-added services.</a></h4>
            </div>
            <%If bHasFeatured = True Then%>
                <div>
                    <h4 class="h4 text-danger">Featured Event (<a href="/misc/featured_events.asp">Add Your Event</a>)</h4>
                    <img class="img-responsive" src="/featured_events/images/<%=FeaturedEvent(6)%>" alt="<%=FeaturedEvent(1)%>" style="float: right;">
                    <ul class="text-danger">
                        <li><%=FeaturedEvent(1)%></li>
                        <li><%=FeaturedEvent(2)%></li>
                        <li><%=FeaturedEvent(3)%></li>
                        <li><a class="text-danger" href="/featured_events/featured_clicks.asp?featured_events_id=<%=FeaturedEvent(0)%>&amp;click_page=<%=sClickPage%>" onclick="openThis(this.href,1024,768);return false;">Website</a></li>
                    </ul>
                    <div class="text-danger"><%=FeaturedEvent(5)%></div>
                </div>
            <%End If%>
            <div class="bg-warning">
                <h4 class="h4">Nordic Ski Technology</h4>
                <p>Our Nordic Ski process uses a gantry structure with four or more overhead antennae spanning 18'- 26' or more as the situation requires.  
                    Participants wear vest-type bibs with an RFID tag adhered to the inside front and back of the bib.  Below is a video indicating how 
                    to attach the Nordic Ski adhesive tags.</p>
                <iframe class="embed-responsive-item" src="https://www.youtube.com/embed/e3VRgq48k8A" frameborder="0" allowfullscreen></iframe>            
            </div>
            <div class="bg-danger">
                <h4 class="h4">Our Timing System...</h4>
                 <p>
                    <a href="http://www.rfidtiming.com/" onclick="openThis(this.href,1024,768);return false;">
			        <img class="img-responsive" src="/graphics/rfid.jpg" alt="RFID Timing" style="width:150px;float: left;"></a>
                    <span style="font-style: italic;">RFID (chip) Timing:</span>
                    <a href="http://www.rfidtiming.com/" onclick="openThis(this.href,1024,768);return false;" style="font-weight:bold;">Ultra</a>
                    - The ultimate rifd timing/chip timing system.  We chose the Ultra system because it is one of the leading technologies in RFID 
                    timing and chip timing.  It allows us the option of being able to source UHF tags from anywhere, meaning we keep our
                    prices competitive with rfid timing or chip timing providers without compromising quality.
                </p>
            </div>
        </div>
         <div class="col-md-4">
           <div class="bg-primary">
                <h4 class="h4">Think About This...</h4>
                <p><%=sThinkAbout%></p>
            </div>
            <div>
                <h4 class="h4">User Login</h4>
                <%If Not sLoginErr = vbNullString Then%>
                    <p class="bg-danger text-danger"><%=sLoginErr%></p>
                <%End If%>
				<form role="form" class="form-horizontal" name="site_login" method="Post" action="http://www.gopherstateevents.com/default.asp?sign_out=n" 
                onsubmit="return chkFlds();">
			    <div class="form-group">
				    <label for="user_name" class="control-label col-xs-4">User &nbsp;&nbsp;&nbsp; Name:</label>
				    <div class="col-xs-8">
                        <input type="text" class="form-control" name="user_name" id="user_name" value="<%=sUserName%>">
                    </div>
			    </div>
                <div class="form-group">
					<label for="password" class="control-label col-xs-4">Password:</label>
				    <div class="col-xs-8">
                        <input type="password" class="form-control" name="password" id="password"maxlength="12" value="">
                    </div>
                </div>
                <div class="form-group">
					<label for="role" class="control-label col-xs-4">Role:</label>
				    <div class="col-xs-8">
					    <select class="form-control" name="role" id="role">
                            <option value="">&nbsp;</option>
                            <option value="admin">GSE Administrator</option>
                            <option value="coach">CC/Nordic Coach</option>
                            <option value="staff">GSE Staff</option>
                            <option value="meet_dir">CC/Nordic Meet Director</option>
                            <option value="event_dir">Fitness Event Director</option>
                            <option value="perf_trkr">Performance Tracker</option>
                        </select>

                    </div>
                </div>
                <div class="form-group">
					<input class="form-control" type="hidden" name="submit_login" id="submit_login" value="submit_login">
					<input class="form-control" type="submit" name="submit1" id="submit1" value="Login">
                    &nbsp;&nbsp;
					<a href="javascript:pop('/misc/forgot_login.asp',600,750)" style="font-size:0.85em;">Forgot Sign In?</a>
                </div>
				</form>
            </div>
             <!--
            <div>
                <h4 class="h4 bg-info">Check This Out!</h4>
                <iframe class="embed-responsive-item" src="https://www.youtube.com/embed/8pZQckevn70" frameborder="0" allowfullscreen></iframe>
            </div>
             -->
            <div class="row">
                <div class="col-sm-7 bg-success">
		            <h4 class="h4">At The Races</h4>
                    <%For i = 0 To UBound(RaceReport, 2) - 1%>
                        <div>
                            <h5 class="text-warning"><%=RaceReport(1, i)%></h5>
                            <h5 class="text-warning"><%=RaceReport(2, i)%></h5>
                            <h5 class="text-warning"><%=RaceReport(3, i)%></h5>
                            <ul class="text-warning">
                                <%If Not RaceReport(4, i) & "" = "" Then%>
                                    <li style="padding: 2px;"><span style="font-weight:bold;">Weather:</span>&nbsp;<%=RaceReport(4, i)%></li>
                                <%End If%>
                                <%If Not RaceReport(5, i) & "" = "" Then%>
                                    <li style="padding: 2px;"><span style="font-weight:bold;">Race Report:</span>&nbsp;<%=RaceReport(5, i)%></li>
                                <%End If%>
                                <%If Not RaceReport(6, i) & "" = "" Then%>
                                    <li style="padding: 2px;"><span style="font-weight:bold;">Gallery:</span>&nbsp;<%=RaceReport(6, i)%></li>
                                <%End If%>
                                <li style="padding: 2px;">
                                    <a href="/results/fitness_events/results.asp?event_type=<%=RaceReport(7, i)%>&event_id=<%=RaceReport(0, i)%>" 
                                    style="font-weight:bold;">Results/Pix</a>
                                </li>
                            </ul>
                        </div>
                        <hr>
                    <%Next%>
                </div>
                <div class="col-sm-5">
                    <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
                    <!-- GSE Vertical ad -->
                    <ins class="adsbygoogle"
                            style="display:block"
                            data-ad-client="ca-pub-1381996757332572"
                            data-ad-slot="6120632641"
                            data-ad-format="auto"></ins>
                    <script>
                    (adsbygoogle = window.adsbygoogle || []).push({});
                    </script>
                </div>
            </div>
			<!--
            <div>
                <h4 class="h4">Recent Press Releases</h4>
                <img class="img-responsive" src="/graphics/newspaper.jpg" alt="Press Release" style="border:none;width:100px;float: right;">
                <ul style="font-size: 0.9em;margin-top: 0;padding-top: 0;">
                                        <li style="list-style-type: circle;padding-top: 5px;"><a href="http://www.gopherstateevents.com/press_releases/randy_murph.pdf" onclick="openThis(this.href,1024,768);return false;">GSE Supports Military Events</a></li>
                    <li style="list-style-type: circle;padding-top: 5px;"><a href="http://www.gopherstateevents.com/press_releases/bauman_rovn.pdf" onclick="openThis(this.href,1024,768);return false;">GSE Back With Bauman-Rovn</a></li>
                    <li style="list-style-type: circle;padding-top: 5px;"><a href="http://www.gopherstateevents.com/press_releases/runnin_with_law.pdf" onclick="openThis(this.href,1024,768);return false;">GSE To Run With The Law</a></li>
                    <li style="list-style-type: circle;padding-top: 5px;"><a href="http://www.gopherstateevents.com/press_releases/nordic_pair.pdf" onclick="openThis(this.href,1024,768);return false;">GSE Signs Up Nordic Pair</a></li>
                    <li style="list-style-type: circle;padding-top: 5px;"><a href="http://www.gopherstateevents.com/press_releases/gray_ghost.pdf" onclick="openThis(this.href,1024,768);return false;">GSE Adds Gray Ghost 5k</a></li>
                    <li style="list-style-type: circle;padding-top: 5px;"><a href="http://www.gopherstateevents.com/press_releases/new_prague.pdf" onclick="openThis(this.href,1024,768);return false;">GSE To Time Run New Prague</a></li>
                    <li style="list-style-type: circle;padding-top: 5px;"><a href="http://www.gopherstateevents.com/press_releases/meg_meet.pdf" onclick="openThis(this.href,1024,768);return false;">GSE Welcomes Mega Meet</a></li>
                </ul>
            </div>
            -->
        </div>
	</div>
	<!--#include file = "includes/footer.asp" -->
</div>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>
</body>
</html>
