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

<div class="accordion" id="leftMenu">
    <div class="accordion-group">
        <div class="accordion-heading">
            <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseOne">
                <i class="icon-th"></i> Profile
            </a>
        </div>
        <div id="collapseOne" class="accordion-body collapse" style="height: 0px; ">
            <div class="accordion-inner">
                <ul>
                    <%If Session("role") = "coach" Then%>
                        <li><a href="/cc_meet/coach/coach_home.asp" rel="nofollow">My Profile</a></li>
                        <li>
                            <a href="#" class="dropdown-toggle" data-toggle="dropdown">Team Staff<b class="caret"></b></a>
                            <ul class="dropdown-menu multi-level">
                                <li><a href="/cc_meet/coach/staff.asp" rel="nofollow">View/Edit</a></li>
                                <li><a href="/cc_meet/coach/add_staff.asp" rel="nofollow">Add New</a></li>
                                <li><a href="/cc_meet/coach/contacts/team_contacts.asp" rel="nofollow">Team Contacts</a></li>
                            </ul>
                        </li>
                    <%ElseIf Session("role") = "team_staff" Then%>
                        <li><a href="/cc_meet/coach/staff_profile.asp" rel="nofollow">My Profile</a></li>
                    <%End If%>
                </ul>
            </div>
        </div>
    </div>
    <div class="accordion-group">
        <div class="accordion-heading">
            <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseTwo">
                <i class="icon-th-list"></i> Communications
            </a>
        </div>
        <div id="collapseTwo" class="accordion-body collapse" style="height: 0px; ">
            <div class="accordion-inner">
                <ul>
                    <li><a href="/cc_meet/coach/communications/grp_email/grp_email.asp" rel="nofollow">Send Email</a></li>
                    <li><a href="/cc_meet/coach/communications/grp_email/email_log.asp" rel="nofollow">Email Log</a></li>
                    <li><a href="/cc_meet/coach/communications/grp_email/email_validate.asp" rel="nofollow">Email Validator</a></li>
                    <li><a href="/cc_meet/coach/communications/txt_msg/txt_msg.asp" rel="nofollow">Send Text Msg</a></li>
                </ul>
            </div>
        </div>
    </div>
    <div class="accordion-group">
        <div class="accordion-heading">
            <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseThree">
                <i class="icon-list-alt"></i>Roster & Line-Ups
            </a>
        </div>
        <div id="collapseThree" class="accordion-body collapse" style="height: 0px; ">
            <div class="accordion-inner">
                <ul>
                    <li><a href="/cc_meet/coach/roster/view_roster.asp" rel="nofollow">View/Edit</a></li>
                    <li><a href="/cc_meet/coach/roster/archived_roster.asp" rel="nofollow">Archives</a></li>
                    <li><a href="/cc_meet/coach/meets/lineup_mgr.asp" rel="nofollow">Meets/Line-Ups</a></li>
                </ul>                 
            </div>
        </div>
    </div>
    <div class="accordion-group">
        <div class="accordion-heading">
            <a class="accordion-toggle" data-toggle="collapse" data-parent="#leftMenu" href="#collapseFour">
                <i class="icon-list-alt"></i>Help Videos
            </a>
        </div>
        <div id="collapseFour" class="accordion-body collapse" style="height: 0px; ">
            <div class="accordion-inner">
                <ul>
                    <li><a href="javascript:pop('https://youtu.be/Zs1cpdc_Ilo',1024,768)" rel="nofollow">Site Use</a></li>
                    <li><a href="javascript:pop('http://youtu.be/joJdj5TIFIs',1024,768)" rel="nofollow">Relay Manager</a></li>
                    <li><a href="javascript:pop('https://youtu.be/e3VRgq48k8A',1024,768)" rel="nofollow">RFID Tags</a></li>
                </ul>                 
            </div>
        </div>
    </div>
</div>


