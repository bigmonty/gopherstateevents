<style type="text/css">
#leftMenu .accordion-group {
    margin-bottom: 0px;
    border:0px;
    -webkit-border-radius: 0px;
    -moz-border-radius: 0px;
    border-radius: 0px;
}    

#leftMenu .accordion-heading {
    height: 34px;
    border-top: 1px solid #717171; /* inner stroke */
    border-bottom: 1px solid #5A5A5A; /* inner stroke */
    background-color: #353535; /* layer fill content */
    background-image: -moz-linear-gradient(90deg, #595b59 0%, #616161 100%); /* gradient overlay */
    background-image: -o-linear-gradient(90deg, #595b59 0%, #616161 100%); /* gradient overlay */
    background-image: -webkit-linear-gradient(90deg, #595b59 0%, #616161 100%); /* gradient overlay */
    background-image: linear-gradient(90deg, #595b59 0%, #616161 100%); /* gradient overlay */
    list-style-type:none;
}  

#leftMenu .accordion-heading  a{  
    color: #cbcbcb; /* text color */
    text-shadow: 0 1px 0 #3b3b3b; /* drop shadow */
    text-decoration:none;
    font-weight:bold;  
}

#leftMenu .accordion-heading  a:hover{  
    color:#ccc     
}

#leftMenu .accordion-heading .active {
    width: 182px;
    height: 34px;
    border: 1px solid #5b5b5b; /* inner stroke */
    background-color: #353535; /* layer fill content */
    background-image: -moz-linear-gradient(90deg, #4b4b4b 0%, #555 100%); /* gradient overlay */
    background-image: -o-linear-gradient(90deg, #4b4b4b 0%, #555 100%); /* gradient overlay */
    background-image: -webkit-linear-gradient(90deg, #4b4b4b 0%, #555 100%); /* gradient overlay */
    background-image: linear-gradient(90deg, #4b4b4b 0%, #555 100%); /* gradient overlay */
}
</style>

<div class="col-sm-2">
    <div class="accordion" id="leftMenu">
        <div class="accordion-group">
            <div class="accordion-heading">
                <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseOne">
                    <i class="icon-th"></i> Finances
                </a>
            </div>
            <div id="collapseOne" class="accordion-body collapse" style="height: 0px; ">
                <div class="accordion-inner">
                    <ul>
                        <li><a href="/admin/finances/items.asp">Items</a></li>
                        <li><a href="/admin/finances/staff_ledger.asp">Staff</a></li>
                        <li><a href="/admin/finances/events_ledger.asp">Events</a></li>
                        <li><a href="/admin/finances/income_expenses.asp">Income/Expense</a></li>
                        <li><a href="/admin/finances/summary.asp">Summary</a></li>
                        <li><a href="/admin/finances/graphs/yearly.asp">Graphs</a></li>
                        <li><a href="/admin/finances/month-by-month.asp">Month-By-Month</a></li>
                        <li><a href="/admin/debt/debt.asp" rel="nofollow">Debt</a></li>
                        <li><a href="/admin/finances/balance_sheet.asp">Balance Sheet</a></li>
                    </ul>
                </div>
            </div>
        </div>
        <div class="accordion-group">
            <div class="accordion-heading">
                <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseTwo">
                    <i class="icon-th-list"></i> Staff Manager
                </a>
            </div>
            <div id="collapseTwo" class="accordion-body collapse" style="height: 0px; ">
                <div class="accordion-inner">
                    <ul>
                        <li><a href="/admin/staff/add_staff.asp">Add Staff</a></li>
                        <li><a href="/admin/staff/staff_pref.asp">Availability</a></li>
                        <li><a href="/staff/calendar.asp">Calendar</a></a></li>
                        <li><a href="/admin/staff/contact_staff.asp">Contact</a></li>
                        <li><a href="/admin/staff/edit_staff.asp">Edit</a></li>
                        <li><a href="/admin/staff/event_assign.asp">Events</a></li>
                        <li><a href="/admin/staff/event_matrix.asp">Event Matrix</a></li>
                        <li><a href="/admin/staff/staff_login.asp">Logins</a></li>
                        <li><a href="/admin/staff/resources.asp">Resources</a></li>
                        <li><a href="/admin/staff/training_videos.asp">Training Videos</a></li>
                        <li><a href="/admin/staff/staff.asp">View</a></li>
                    </ul>                 
                </div>
                </div>
        </div>
        <div class="accordion-group">
            <div class="accordion-heading">
                <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseThree">
                    <i class="icon-list-alt"></i> Fitness Events
                </a>
            </div>
            <div id="collapseThree" class="accordion-body collapse" style="height: 0px; ">
                <div class="accordion-inner">
                    <ul>
		                <li><a href="/admin/events/create_event.asp" rel="nofollow">Create Event</a></li>
                        <li><a href="/admin/participants/dont_send_log.asp" rel="nofollow">Dont Send Log</a></li>
                        <li><a href="/admin/event_groups/event_groups.asp" rel="nofollow">Event Groups</a></li>
                        <li><a href="/admin/events/event_mgr.asp" rel="nofollow">Events</a></li>
		                <li><a href="/admin/event_dir/event_dirs.asp" rel="nofollow">Event Directors</a></li>
		                <li><a href="/admin/events_promo/email_event_promo.asp" rel="nofollow">Event Promo</a></li>
                        <li><a href="/admin/events/event_surveys.asp" rel="nofollow">Event Surveys</a></li>
                        <li><a href="/admin/events/event_timeline.asp" rel="nofollow">Event Timeline</a></li>
                        <li><a href="/admin/events/timeline_matrix.asp" rel="nofollow">Timeline Matrix</a></li>
                        <li><a href="/admin/featured_events/featured_events.asp">Featured Events</a></li>
		                <li><a href="/admin/events/group_email.asp" rel="nofollow">Group Email</a></li>
                        <li><a href="/admin/logins.asp" rel="nofollow">Logins</a></li>
		                <li><a href="/admin/participants/part_manager.asp" rel="nofollow">Participants</a></li>
                        <li><a href="/admin/media/all_pix.asp" rel="nofollow">Pictures</a></li>
		                <li><a href="/admin/prospects/prspcts_mgr.asp" rel="nofollow">Prospects</a></li>
                        <li><a href="/admin/media/all_vids.asp" rel="nofollow">Videos</a></li>
                        <li><a href="/admin/series/series_mgr.asp" rel="nofollow">Series</a></li>
                    </ul>                 
                </div>
            </div>
        </div>
        <div class="accordion-group">
            <div class="accordion-heading">
                <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseFour">
                    <i class="icon-list-alt"></i>CC/Nordic
                </a>
            </div>
            <div id="collapseFour" class="accordion-body collapse" style="height: 0px; ">
                <div class="accordion-inner">
                    <ul>
		                <li><a href="/ccmeet_admin/meets.asp" rel="nofollow">Meets</a></li>
                        <li><a href="/ccmeet_admin/create_meet.asp" rel="nofollow">Add Meet</a></li>
                        <li><a href="/ccmeet_admin/add_meet_dir.asp" rel="nofollow">Add Meet Director</a></li>
		                <li><a href="/ccmeet_admin/manage_meet/meet_dir_data.asp" rel="nofollow">Meet Dir Data</a></li>
		                <li><a href="/ccmeet_admin/manage_coach/coach_data.asp" rel="nofollow">Coach Data</a></li>
		                <li><a href="/ccmeet_admin/manage_team/team_data.asp" rel="nofollow">Team Data</a></li>
		                <li><a href="/ccmeet_admin/grp_email/email_coaches.asp" rel="nofollow">Email Coaches</a></li>
		                <li><a href="/ccmeet_admin/grp_email/email_meet_dirs.asp" rel="nofollow">Email Meet Dir</a></li>
                        <li><a href="/ccmeet_admin/series/series_mgr.asp" rel="nofollow">Series</a></li>
                        <li><a href="/ccmeet_admin/visitors.asp" rel="nofollow">Visitors</a></li>
                    </ul>                 
                </div>
            </div>
        </div>
        <div class="accordion-group">
            <div class="accordion-heading">
                <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseFive">
                    <i class="icon-cog"></i> Miscellaneous
                </a>
            </div>
            <div id="collapseFive" class="accordion-body collapse" style="height: 0px; ">
                <div class="accordion-inner">
                    <ul>
                        <li><a href="/admin/bib_inventory.asp" rel="nofollow">Bib Inventory</a></li>
                        <li><a href="/admin/condense_users/condense_users.asp" rel="nofollow">Condense Users</a></li>
		                <li><a href="/admin/contact_log.asp" rel="nofollow">Contact Log</a></li>
		                <li><a href="/admin/data_modify.asp" rel="nofollow">Data Modify</a></li>
                        <li><a href="/admin/admin_dont_send.asp" rel="nofollow">Dont Send</a></li>
                        <li><a href="/admin/followers/followers.asp" rel="nofollow">Followers</a></li>
                        <li><a href="/admin/media/media_mgr.asp" rel="nofollow">Media Manager</a></li>
                        <li><a href="/misc/rfid_instructions.doc" rel="nofollow">RFID Process</a></li>
	                    <li><a href="/admin/sponsors/sponsors.asp" rel="nofollow">Sponsors</a></li>
                        <li><a href="/trends/event_trends.asp" rel="nofollow">Trends</a></li>
                        <li><a href="/admin/vira_visitors.asp" rel="nofollow">Visitors</a></li>	
                    </ul>                 
                </div>
            </div>
        </div>
        <div class="accordion-group">
            <div class="accordion-heading">
                <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseSix">
                    <i class="icon-file"></i> Performance
                </a>
            </div>
            <div id="collapseSix" class="accordion-body collapse" style="height: 0px; ">
                <div class="accordion-inner">
                    <ul>
	                    <li><a href="/admin/my_hist/accounts.asp" rel="nofollow">My History</a></li>
	                    <li><a href="/admin/perf_trkr/accounts.asp" rel="nofollow">Perf Tracker</a></li>
                    </ul>                 
                </div>
            </div>
        </div>
    </div>
</div>