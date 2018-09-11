            <%If UBound(Events, 2) > 1 Then%>
                <div class="row">
                    <form role="form" class="form-inline" name="which_event" method="post" action="<%=sThisPage%>?event_id=<%=lEventID%>&amp;which_tab=<%=sWhichTab%>">
                    <label for="events">Select Event:</span>
                    <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
                        <%For i = 0 to UBound(Events, 2) - 1%>
                            <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
                                <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%></option>
                            <%Else%>
                                <option value="<%=Events(0, i)%>"><%=Events(1, i)%></option>
                            <%End If%>
                        <%Next%>
                    </select>
                    <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
                    <input class="form-control" type="submit" name="get_event" id="get_event" value="Get This Event">
                    </form>
                </div>
            <%End If%>		
                    
            <div class="row">
                <div>
                    <span style="font-weight:bold;">PLEASE NOTE:</span>  We will come prepared to manage your event based on these settings.  Please have them up-to-date AT LEAST TWO WEEKS
                    PRIOR TO THE EVENT!  <span style="color: red;">You will not be able to make any changes in this data the final week before the
                    event.  If you notice any changes to be made then please contact us <a href="mailto:bob.schneider@gopherstateevents.com">
                    <span style="font-weight: bold;">via email</span></a> or by calling 
                    <span style="font-weight: bold;color: #039;">612-720-8427</span></span>
                </div>

                <ul class="nav">
                    <li class="nav-item"><a class="nav-link" href="event_settings.asp?event_id=<%=lEventID%>" onclick="openThis(this.href,1024,768);return false;">Print Settings</a></li>
                    <li class="nav-item"><a class="nav-link" href="<%=sInfoLink%>" style="font-weight: bold;" onclick="openThis(this.href,1024,768);return false;">Event Info</a></li>
                </ul>
            </div>
