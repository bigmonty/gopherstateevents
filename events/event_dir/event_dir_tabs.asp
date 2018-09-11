                    <table class="table">
                        <tr>
                            <%Select Case sWhichTab%>
                                <%Case "General"%>
                                    <th class="tabs" style="background-color: #ececec;"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                                <%Case "Venue"%>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs" style="background-color: #ececec;"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                                <%Case "Preferences"%>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs" style="background-color: #ececec;"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                                <%Case "Registration"%>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs" style="background-color: #ececec;"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                                <%Case "Race Data"%>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs" style="background-color: #ececec;"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                                <%Case "Post Race"%>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs" style="background-color: #ececec;"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                                <%Case "Participants"%>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs" style="background-color: #ececec;"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                                <%Case "Documents"%>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=General">General</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Venue">Venue</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Preferences">Preferences</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Registration">Registration</a></th>
                                    <th class="tabs"><a href="race_data.asp?event_id=<%=lEventID%>&amp;which_tab=Race Data">Race Data</a></th>
                                    <th class="tabs"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Post Race">Post Race</a></th>
                                    <th class="tabs"><a href="part_data.asp?event_id=<%=lEventID%>&amp;which_tab=Participants">Participants</a></th>
                                    <th class="tabs" style="background-color: #ececec;"><a href="event_admin.asp?event_id=<%=lEventID%>&amp;which_tab=Documents">Documents</a></th>
                            <%End Select%>
                        </tr>
                    </table>
