				<br>
                <ul class="list-inline">
					<li class="list-group-item list-group-item-success">
                        <a href="/cc_meet/meet_dir/meets/meet_info/view_meet.asp?meet_id=<%=lThisMeet%>">Meet Home Page</a>
					</li>
					<li class="list-group-item list-group-item-success">
                        <a href="/cc_meet/meet_dir/meets/meet_info/meet_classes.asp?meet_id=<%=lThisMeet%>">Meet Classes</a>
					</li>
					<li class="list-group-item list-group-item-success">
					    <a href="/cc_meet/meet_dir/teams/team_data.asp?meet_id=<%=lThisMeet%>">Team Data</a>
					</li>
					<li class="list-group-item list-group-item-success">
                        <a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/course_maps/upload_map.asp?meet_id=<%=lThisMeet%>',500,300)">Upload Course Map</a>
					</li>
					<li class="list-group-item list-group-item-success">
					    <a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/info_sheets/upload_info.asp?meet_id=<%=lThisMeet%>',500,300)">Upload Info Sheet</a>
					</li>
                    <%If Not sMapLink = vbNullString Then%>
						<li class="list-group-item list-group-item-success">
						    <a href="javascript:pop('<%=sMapLink%>',1024,768)">MapQuest Link to Site</a>
                        </li>
					<%End If%>
					<%If Not sMeetInfoSheet = vbNullString Then%>
						<li class="list-group-item list-group-item-success">
						    <a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/info_sheets/<%=sMeetInfoSheet%>',1024,768)">Info Sheet</a>
                        </li>
					<%End If%>
					<%If Not sCourseMap = vbNullString Then%>
						<li class="list-group-item list-group-item-success">
						    <a href="javascript:pop('/cc_meet/meet_dir/meets/meet_info/course_maps/<%=sCourseMap%>',1024,768)">Course Map</a>
                        </li>
					<%End If%>
					<%If Date < CDate(dWhenShutdown) Then%>
						<li class="list-group-item list-group-item-success">
						    <a href="/cc_meet/meet_dir/teams/edit_teams.asp?meet_id=<%=lThisMeet%>">Add/Remove Teams</a>
                        </li>
					<%End If%>
				</ul>
