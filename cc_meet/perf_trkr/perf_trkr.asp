<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, conn2
Dim i
Dim dExpiration

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "login.asp"

dExpiration = Date + 2

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Cross=Country-Nordic Ski Performance Tracker Home</title>
</head>

<body>
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->

    <div class="row">
        <h3 class="h3">Performance Tracker Home</h3>
        <div class="col-md-10">
            <%If CDate(dExpiration) >= Date Then%> 
                <%If CDate(dExpiration) < CDate("1/1/2030") Then%>
                    <div class="row">
                        <div class="bg-warning col-md-9">
                            GSE's Performance Tracker carries a one-time fee of $5 to cover administrative costs after a free trial of 7 days.  This includes 
                            cross-country running and nordic skiing.  Your trial expires on <%=dExpiration%>.  If you believe you are getting this message in error, 
                            please <a style="color:#fff;" href="mailto:bob.schneider@gopherstateevents.com">email us</a> and we will look into it.  
                            <span style="font-weight: bold;">Please include your first name, last name, grade, gender, and school/team</span> with all 
                            correspondence.
                        </div>

                        <div class="col-md-3" style="text-align: center;">
                            <form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_top">
                            <input type="hidden" name="cmd" value="_s-xclick">
                            <input type="hidden" name="hosted_button_id" value="8TJ9A86Y9942G">
                            <input type="image" src="https://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif" border="0" name="submit" alt="PayPal - The safer, easier way to pay online!">
                            <img alt="" border="0" src="https://www.paypalobjects.com/en_US/i/scr/pixel.gif" width="1" height="1">
                            </form>
                        </div>
                    </div>
                    <br>
                <%End If%>

                <div class="row">
                    <div class="col-md-3 bg-warning">
                        <a style="color:#fff;" href="profile.asp">
                        <h4 class="h4">My Profile</h4>
                            <p>Settings and contact info</p>
                        </a>
                    </div>
                    <div class="col-md-3 bg-danger">
                        <a style="color:#fff;" href="create_pack.asp">
                            <h4 class="h4">Create Pack</h4>
                            <p>Create a new group to follow.</p>
                        </a>
                    </div>
                    <div class="col-md-3 bg-info">
                        <a style="color:#fff;" href="my_packs.asp">
                            <h4 class="h4">My Packs</h4>
                            <p>Add to/edit my existing groups.</p>
                        </a>
                    </div>
                    <div class="col-md-3 bg-warning">
                        <a style="color:#fff;" href="perf_lists.asp">
                            <h4 class="h4">Performance Lists</h4>
                            <p>My packs' performances lists.</p>
                        </a>
                    </div>
                </div>
    
                <div class="row">
                    <div class="col-md-3 bg-danger">
                        <a style="color:#fff;" href="perf_graphs.asp">
                            <h4 class="h4">Performance Graphs</h4>
                            <p>Performance graphs by my packs.</p>
                        </a>
                    </div>
                    <div class="col-md-3 bg-info">
                        <a style="color:#fff;" href="social.asp">
                            <h4 style="color"#fff;" class="h4">Social Connectivity</h4>
                            <p style="color"#fff;">Post fb & twitter accounts.</p>
                        </a>
                    </div>
                    <div class="col-md-3 bg-warning">
                        <a style="color:#fff;" href="http://www.my-etraxc.com/" onclick="openThis(this.href,1024,760);return false;">
                            <h4 style="color"#fff;" class="h4">My-eTRaXC</h4>
                            <p style="color"#fff;">Personal training account.</p>
                        </a>
                    </div>
                    <div class="col-md-3 bg-danger">
                        <a style="color:#fff;" href="results_notif.asp">
                            <h4 class="h4">Results Notification</h4>
                            <p>Get your results via email/sms..</p>
                        </a>
                    </div>
                </div>
    
                <div class="row">
                    <div class="col-md-3 bg-info">
                        <a style="color:#fff;" href="perf_graphs.asp">
                            <h4 class="h4">My History</h4>
                            <p>Looking at my GSE performances.</p>
                        </a>
                    </div>
                    <div class="col-md-3 bg-warning">
                        <h4 style="color:#fff;" class="h4">Space reserved...</h4>
                        <p>...for coming features
                    </div>
                    <div class="col-md-3 bg-danger">
                        <h4 style="color:#fff;" class="h4">Space reserved...</h4>
                        <p>...for coming features
                    </div>
                    <div class="col-md-3 bg-info">
                        <h4 style="color:#fff;" class="h4">Space reserved...</h4>
                        <p>...for coming features
                    </div>
                </div>
            <%Else%>
                <div class="row">
                    <div class="bg-danger text-danger col-md-9">
                        I am sorry.  Your trial period has expired.  You can restore access to your account by using the "Buy Now" button above.  The cost is a 
                        one-time fee of $5.  Please allow 24 hours for your account to be restored.  You will be notified via email once it is open.  If you
                        believe you are getting this message in error, please <a href="mailto:bob.schneider@gopherstateevents.com">email us</a> and we will look into it.
                        <span style="font-weight: bold;">Please include your first name, last name, grade, gender, and school/team</span> with all correspondence.
                    </div>

                    <div class="col-md-3">
                        <form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_top">
                        <input type="hidden" name="cmd" value="_s-xclick">
                        <input type="hidden" name="hosted_button_id" value="8TJ9A86Y9942G">
                        <input type="image" src="https://www.paypalobjects.com/en_US/i/btn/btn_buynowCC_LG.gif" border="0" name="submit" alt="PayPal - The safer, easier way to pay online!">
                        <img alt="" border="0" src="https://www.paypalobjects.com/en_US/i/scr/pixel.gif" width="1" height="1">
                        </form>
                    </div>
                </div>
            <%End If%>
        </div>
        <div class="col-md-2">
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

</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.close
Set conn = Nothing

conn2.close
Set conn2 = Nothing
%>
</html>
