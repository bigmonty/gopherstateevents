<div class="navbar navbar-default navbar-fixed-top no-print" role="navigation">
    <div class="container">
        <div class="navbar-header">
            <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
                <span class="sr-only">Toggle navigation</span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
            </button>
        </div>
        <div class="collapse navbar-collapse">
            <ul class="nav navbar-nav">
                <li class="active"><a href="/default.asp">Home</a></li>
                <li>
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">Open Results<b class="caret"></b></a>
                    <ul class="dropdown-menu multi-level">
                        <li><a href="/results/fitness_events/results.asp?event_type=5"><span>Road Race</span></a></li>
				        <li><a href="/results/fitness_events/results.asp?event_type=46"><span>Nordic-Snowshoe</span></a></li>
                        <li><a href="/results/fitness_events/results.asp?event_type=3"><span>Off-Road Bike</span></a></li>
                        <li><a href="/results/fitness_events/results.asp?event_type=910"><span>Multi-Sport</span></a></li>
                        <li><a href="/results/fitness_events/results.asp?event_type=7"><span>Mud/Obstacle/Trail</span></a></li>
                        <li><a href="/results/fitness_events/results.asp?event_type=2"><span>Specialty</span></a></li>
                        <li><a href="/results/fitness_events/digital_results.asp" onclick="openThis(this.href,1024,768);return false;"><span>Kiosk Version</span></a></li>
                    </ul>
                </li>
                <li>
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">School Results<b class="caret"></b></a>
                    <ul class="dropdown-menu multi-level">
                        <li><a href="/results/cc_rslts/cc_rslts.asp?sport=cc"><span>CC Running</span></a></li>
				        <li><a href="/results/cc_rslts/cc_rslts.asp?sport=nordic"><span>Nordic Ski</span></a></li>
                    </ul>
                </li>
                <li><a href="/calendar/calendar.asp">Calendar</a></li>
                <li>
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">Series<b class="caret"></b></a>
                    <ul class="dropdown-menu multi-level">
                        <li><a href="/series/series_info.asp?year=<%=Year(Date)%>"><span>Open Series</span></a></li>
				        <li><a href="/series/cc_nordic/series_info.asp?year=<%=Year(Date)%>"><span>School Series</span></a></li>
                    </ul>
                </li>
                <li>
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">Performance<b class="caret"></b></a>
                    <ul class="dropdown-menu multi-level">
                        <li><a href="/misc/honor_roll.asp"><span>Honor Roll</span></a></li>
                        <li><a href="http://www.runmdra.org/grandprix/" onclick="openThis(this.href,1024,768);return false;"><span>Grand Prix</span></a></li>                 
                         <li><a href="/cc_meet/perf_trkr/login.asp" onclick="openThis(this.href,1024,768);return false;"><span>Performance Tracker</span></a></li>                 
                    </ul>
                </li>
                <li>
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">Resources<b class="caret"></b></a>
                    <ul class="dropdown-menu multi-level">
                        <li><a href="/misc/sms_kiosk.asp" onclick="openThis(this.href,1024,768);return false;"><span>SMS Set-Up</span></a></li>
                        <li><a href="/results/fitness_events/digital_results.asp" onclick="openThis(this.href,1024,768);return false;"><span>Kiosk Version</span></a></li>
                        <li><a href="http://www.gseannouncer.com" onclick="openThis(this.href,1024,768);return false;"><span>Announcer</span></a></li>
  			            <li><a href="http://www.usatfne.org/road/checklist.html" onclick="openThis(this.href,1024,768);return false;"><span>Event Checklist</span></a></li>
                        <li><a href="http://us11.campaign-archive1.com/?u=1be1fa83d91e63dd91b86f7dc&id=0ebf0363e6&e=[UNIQID]" onclick="openThis(this.href,1024,768);return false;"><span>Fitness Newsletter</span></a></li>
                        <li><a href="http://us11.campaign-archive2.com/?u=1be1fa83d91e63dd91b86f7dc&id=13bba33def&e=[UNIQID]" onclick="openThis(this.href,1024,768);return false;"><span>CC Newsletter</span></a></li>
                        <li><a href="/misc/sample_rslts.asp" onclick="openThis(this.href,800,600);return false;"><span>Results Email</span></a></li>
                         <li class="divider"></li>
                        <li><a href="http://www.etraxc.com" onclick="openThis(this.href,1024,768);return false;"><span>eTRaXC</span></a></li>
                        <li><a href="http://www.my-etraxc.com" onclick="openThis(this.href,1024,768);return false;"><span>My-eTRaXC</span></a></li>
                         <li class="divider"></li>
                        <li><a href="/race_vids/race_vids.asp"><span>Finish Line Videos</span></a></li>
                        <li><a href="/gallery/finish_pix.asp"><span>Finish Line Pix</span></a></li>                
                    </ul>
                </li>
                <li>
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">About<b class="caret"></b></a>
                    <ul class="dropdown-menu multi-level">
                    <li><a href="/about/about_gse.asp"><span>About Us</span></a></li>
                    <li><a href="/about/gse_staff.asp"><span>Our Staff</span></a></li>
                    <li><a href="/about/offerings.asp"><span>Our Services</span></a></li>
                    <li><a href="/about/privacy.asp"><span>Privacy Policy</span></a></li>
                    <li><a href="/about/testim.asp"><span>Testimonials</span></a></li>
                    </ul>
                </li>
                <li><a href="/sponsors/sponsors.asp"><span>Partners</span></a></li>
                <li><a href="/misc/teams.asp" onclick="openThis(this.href,800,600);return false;"><span>Teams</span></a></li>
                <li><a href="/misc/contact.asp"><span>Contact</span></a></li>
            </ul>
        </div><!--/.nav-collapse -->
    </div>
</div>
<div id="fb-root"></div>
<div class="container no-print" style="margin-top: 50px;">
    <div class="row">
        <div class="col-xs-9">
            <div  style="padding: 3px 0 0 3px;background: none;float: left;">
                <a href="https://www.facebook.com/GopherStateEvents">
                <img src="/graphics/social_media/fb.png" alt="Facebook" height="20"></a>
            </div>
            <div  style="padding: 3px 0 0 3px;background: none;float: left;">
                <a href="https://www.instagram.com/gopherstateevents/">
                <img src="/graphics/social_media/instagram.jpg" alt="Instagram" height="20"></a>
            </div>
           <div  style="padding: 3px 0 0 3px;background: none;float: left;">
                <a href="http://www.youtube.com/channel/UCs09DthS7jEZy5srWZEDJQw">
                <img src="/graphics/social_media/youtube.png" alt="YouTube" height="20"></a>
            </div>
            <div style="padding: 3px 0 0 3px;background: none;float: left;">
                <a href="http://plus.google.com/100097568010679842973?prsrc=3" rel="publisher" style="text-decoration:none;">
                <img src="/graphics/social_media/GooglePlus-512-Red.png" alt="Google+" height="20"></a>
            </div>
            <div style="padding: 3px 0 0 3px;background: none;float: left;">
                <a href="http://www.linkedin.com/pub/bob-schneider/8/96a/876">
                    <img src="/graphics/social_media/LinkedIn-Logo.png" height="20" alt="View Bob Schneider's profile on LinkedIn">
                </a>     
            </div>
            <div style="padding: 3px 0 0 3px;background: none;float: left;">
                <a href="https://twitter.com/gsetiming" class="twitter-follow-button" data-show-count="false">
                    <img src="/graphics/social_media/Twitter.png" alt="Follow @gsetiming" height="20">
                </a>
                <script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>
            </div>
            <div style="padding: 3px 0 0 3px;background: none;float: left;">
                <a href="tel:+1-612-720-8427" style="font-size: 1.2em;">612-720-8427</a>
            </div>
        </div>
 
        <div class="col-xs-3">
	        <%Select Case Session("role")%>
                <%Case "admin"%>
		            <a href="/default.asp?sign_out=y" style="color: #039;">Sign Out</a>
		            |
		            <a href="/admin/admin.asp" rel="nofollow" style="color: #039;">Admin Portal</a>
	            <%Case "coach"%>
		            <a href="/default.asp?sign_out=y" style="color: red;">Sign Out</a>
		            |
		            <a href="/cc_meet/coach/coach_home.asp" style="color: red;">Coach Home</a>
	            <%Case "meet_dir"%>
		            <a href="/default.asp?sign_out=y" style="color: #039;">Sign Out</a>
		            |
		            <a href="/cc_meet/meet_dir/meet_dir_home.asp" rel="nofollow" style="color: #039;">Meet Director Home</a>
                    |
	            <%Case "event_dir"%>
		            <a href="/default.asp?sign_out=y" style="color: #039;">Sign Out</a>
		            |
		            <a href="/events/event_dir/event_dir_home.asp" rel="nofollow" style="color: #039;">Event Director Home</a>
                    |
	            <%Case "my_hist"%>
		            <a href="/default.asp?sign_out=y" style="color: #039;">Sign Out</a>
		            |
		            <a href="/perf_center/profile.asp" rel="nofollow" style="color: #039;">My Profile</a>
	            <%Case "staff"%>
		            <a href="/default.asp?sign_out=y" style="color: #039;">Sign Out</a>
		            |
		            <a href="/staff/profile.asp" rel="nofollow" style="color: #039;">My Profile</a>
                <%Case Else%>
                    &nbsp;
	        <%End Select%>
        </div>
    </div> 
</div>
<section class="menu cid-qzXzsb3BT2" once="menu" id="menu1-q" data-rv-view="62">
    <nav class="navbar navbar-expand beta-menu navbar-dropdown align-items-center navbar-fixed-top navbar-toggleable-sm">
        <button class="navbar-toggler navbar-toggler-right" type="button" data-toggle="collapse" data-target="#navbarSupportedContent" aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
            <div class="hamburger">
                <span></span>
                <span></span>
                <span></span>
                <span></span>
            </div>
        </button>
        <div class="menu-logo">
            <div class="navbar-brand">
                <span class="navbar-logo">
                    <a href="https://mobirise.com">
                         <img src="assets/images/g-transparent2-351x345.png" alt="Gopher State Events Logo" title="Gopher State Events Logo" media-simple="true" style="height: 3.8rem;">
                    </a>
                </span>
                <span class="navbar-caption-wrap"><a class="navbar-caption text-black display-5" href="http://www.gopherstateevents.com">
                        Gopher State Events</a></span>
            </div>
        </div>
        <div class="collapse navbar-collapse" id="navbarSupportedContent">
            <ul class="navbar-nav nav-dropdown" data-app-modern-menu="true"><li class="nav-item dropdown">
                    <a class="nav-link link dropdown-toggle text-black display-4" href="https://mobirise.com" data-toggle="dropdown-submenu" aria-expanded="false"><span class="mbri-numbered-list mbr-iconfont mbr-iconfont-btn"></span>
                        Results</a><div class="dropdown-menu"><div class="dropdown"><a class="dropdown-item dropdown-toggle text-black display-4" href="https://mobirise.com" data-toggle="dropdown-submenu" aria-expanded="false">Fitness Events</a><div class="dropdown-menu dropdown-submenu"><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=5">Road Race</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=46">Nordic Ski/Snowshoe</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=3">Off-Road Bike</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=910">Multi-Sport</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=7">Mud/Obstacle/Trail</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/fitness_events/results.asp?event_type=2">Specialty</a></div></div><div class="dropdown"><a class="dropdown-item dropdown-toggle text-black display-4" href="https://mobirise.com" aria-expanded="false" data-toggle="dropdown-submenu">School Events<br></a><div class="dropdown-menu dropdown-submenu"><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/cc_rslts/cc_rslts.asp?sport=cc" aria-expanded="false">Cross-Country</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/cc_rslts/cc_rslts.asp?sport=nordic" aria-expanded="false">Nordic Ski</a></div></div></div>
                </li>
                <li class="nav-item dropdown">
                    <a class="nav-link link dropdown-toggle text-black display-4" href="https://mobirise.com" data-toggle="dropdown-submenu" aria-expanded="false"><span class="mbri-numbered-list mbr-iconfont mbr-iconfont-btn"></span>
                        
                        Resources</a><div class="dropdown-menu"><a class="dropdown-item text-black display-4" href="http://www.gseannouncer.com/" target="_blank">Announcer</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/misc/sms_kiosk.asp" target="_blank">SMS Set-Up</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/results/fitness_events/digital_results.asp" target="_blank">Results Kiosk</a><div class="dropdown"><a class="dropdown-item dropdown-toggle text-black display-4" href="https://mobirise.com" data-toggle="dropdown-submenu" aria-expanded="false">Series</a><div class="dropdown-menu dropdown-submenu"><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/series/series_info.asp?year=2017">Fitness Events</a><a class="dropdown-item text-black display-4" href="http://www.gopherstateevents.com/series/cc_nordic/series_info.asp?year=2017">School Events</a></div></div></div>
                </li><li class="nav-item dropdown open"><a class="nav-link link text-black dropdown-toggle display-4" href="https://mobirise.com" aria-expanded="true" data-toggle="dropdown-submenu"><span class="mbri-numbered-list mbr-iconfont mbr-iconfont-btn"></span>
                        
                        Performance</a><div class="dropdown-menu"><a class="text-black dropdown-item display-4" href="http://www.gopherstateevents.com/cc_meet/perf_trkr/login.asp" aria-expanded="true" target="_blank">Performance Tracker</a><a class="text-black dropdown-item display-4" href="http://www.gopherstateevents.com/misc/honor_roll.asp" aria-expanded="true" target="_blank">Honor Roll</a></div></li></ul>
            <div class="navbar-buttons mbr-section-btn"><a class="btn btn-sm btn-black-outline display-4" href="http://www.gopherstateevents.com/misc/login.asp"><span class="mbri-key mbr-iconfont mbr-iconfont-btn"></span>
                    
                    Login!
                </a></div>
        </div>
    </nav>
</section>
