<div class="pos-f-t">
  <div class="collapse" id="navbarToggleExternalContent">
    <div class="bg-dark p-4">
		<a href="/ccmeet_admin/manage_meet/manage_meet.asp?meet_id=<%=lThisMeet%>">Meet Info</a><br>
		<a href="/ccmeet_admin/manage_meet/manage_teams.asp?meet_id=<%=lThisMeet%>">Teams</a><br>
		<a href="javascript:pop('/ccmeet_admin/manage_meet/part_upload/batch_upload.asp?meet_id=<%=lThisMeet%>',400,700)">Upload Participants</a><br>
		<a href="/ccmeet_admin/manage_meet/team_data.asp?meet_id=<%=lThisMeet%>">Team Data</a><br>
		<a href="/ccmeet_admin/manage_meet/races.asp?meet_id=<%=lThisMeet%>">Races</a><br>
		<a href="/ccmeet_admin/manage_meet/results/results_mgr.asp?meet_id=<%=lThisMeet%>">Results</a><br>
		<a href="/ccmeet_admin/manage_meet/manage_bibs.asp?meet_id=<%=lThisMeet%>">Manage Bibs</a><br>
		<a href="/ccmeet_admin/manage_meet/lineups.asp?meet_id=<%=lThisMeet%>" onclick="openThis(this.href,1024,768);return false;">Line-Ups</a><br>
		<a href="/ccmeet_admin/manage_meet/team_labels.asp?meet_id=<%=lThisMeet%>" onclick="openThis(this.href,1024,768);return false;">Team Labels</a><br>
		<a href="/ccmeet_admin/manage_meet/meet_classes.asp?meet_id=<%=lThisMeet%>">Meet Classes</a><br>
		<a href="javascript:pop('/results/cc_rslts/mf_combined.asp?meet_id=<%=lThisMeet%>',1024,768)">Boys-Girls Combined Results</a><br>
		<a href="javascript:pop('/ccmeet_admin/manage_meet/logo.asp?meet_id=<%=lThisMeet%>',1024,768)">Logo</a><br>
		<a href="/ccmeet_admin/manage_meet/team_instr.asp?meet_id=<%=lThisMeet%>">Coach Instr</a><br>
		<a href="/ccmeet_admin/manage_meet/roster_request.asp?meet_id=<%=lThisMeet%>">Roster Reqst</a><br>
		<a href="/ccmeet_admin/manage_meet/lineup_rqst.asp?meet_id=<%=lThisMeet%>">Line-up Rqst</a><br>
		<a href="/ccmeet_admin/manage_meet/confirm_lineup.asp?meet_id=<%=lThisMeet%>">Line-up Confirm</a><br>
		<a href="/ccmeet_admin/manage_meet/cc_vids.asp?event_type=cc&amp;meet_id=<%=lThisMeet%>">Race Videos</a><br>
		<a href="/ccmeet_admin/manage_meet/cc_pix.asp?event_type=cc&amp;meet_id=<%=lThisMeet%>">Race Pix</a><br>
		<a href="/ccmeet_admin/manage_meet/results/email_results.asp?event_type=cc&amp;meet_id=<%=lThisMeet%>">Email Results</a><br>
		<a href="/ccmeet_admin/manage_meet/media_notif.asp?meet_id=<%=lThisMeet%>">Media Notif</a><br>
		<a href="/ccmeet_admin/manage_meet/email_coaches.asp?meet_id=<%=lThisMeet%>">Email Coaches</a><br>
    </div>
  </div>
  <nav class="navbar navbar-dark bg-dark">
    <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarToggleExternalContent" aria-controls="navbarToggleExternalContent" aria-expanded="false" aria-label="Toggle navigation">
      <span class="navbar-toggler-icon"></span>
    </button>
  </nav>
</div>
