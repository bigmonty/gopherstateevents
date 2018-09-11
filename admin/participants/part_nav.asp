
			<ul class="nav">
				<li class="nav-item"><a class="nav-link" href="/staff/parts_events.asp?event_id=<%=lEventID%>">Enter Parts</a></li>
				<li class="nav-item"><a class="nav-link" href="javascript:pop('batch_upload/upload_template.xls',1024,750)">Upload Template</a></li>
				<li class="nav-item"><a class="nav-link" href="javascript:pop('batch_upload/batch_upload.asp?event_id=<%=lEventID%>',1000,750)">Batch Upload</a></li>
				<li class="nav-item"><a class="nav-link" href="/admin/participants/part_sms.asp?event_id=<%=lEventID%>">Participant SMS</a></li>
                <li class="nav-item"><a class="nav-link" href="/admin/participants/part_data.asp?event_id=<%=lEventID%>">Participant List</a></li>
                <li class="nav-item"><a class="nav-link" href="/admin/participants/part_data.asp?event_id=<%=lEventID%>&amp;age_grp_update=y">Age Grps</a></li>
                <li class="nav-item"><a class="nav-link" href="/admin/participants/part_data.asp?event_id=<%=lEventID%>&amp;clean_data=y">Clean Data</a></li>
            </ul>
