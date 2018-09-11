<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim sShowAge, sShowDOB, sShowPhone, sShowCity, sShowSt, sShowEmail, sShowSize, sShowBib, sShowFbook, sShowTwitter, sShowProvider
Dim sShowCell, sShowTeams
Dim iAgeTab, iDOBTab, iPhoneTab, iCityTab, iStTab, iEmailTab, iSizeTab, iBibTab, iFbookTab, iTwitterTab, iProviderTab, iCellTab, iTeamsTab
Dim Events
Dim bInsert

If Session("role") = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventID, EventName, EventDate, Location FROM Events ORDER By EventDate DESC"
rs.Open sql, conn, 1, 2
Events = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_tabs") = "submit_tabs" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SElECT ShowAge, AgeTab, ShowDOB, DOBTab, ShowPhone, PhoneTab, ShowCity, CityTab, ShowSt, StTab, ShowEmail, EmailTab, ShowSize, SizeTab, "
    sql = sql & "ShowBib, BibTab, ShowProvider, ProviderTab, ShowCell, CellTab, ShowTeams, TeamsTab, ShowFbook, FBookTab, "
    sql = sql & "ShowTwitter, TwitterTab FROM PartEntryTabs WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    rs(0).value = Request.Form.Item("show_age")
    rs(1).Value = Request.Form.Item("age_tab")
    rs(2).value = Request.Form.Item("show_dob")
    rs(3).Value = Request.Form.Item("dob_tab")
    rs(4).value = Request.Form.Item("show_phone")
    rs(5).Value = Request.Form.Item("phone_tab")
    rs(6).value = Request.Form.Item("show_city")
    rs(7).Value = Request.Form.Item("city_tab")
    rs(8).value = Request.Form.Item("show_st")
    rs(9).Value = Request.Form.Item("st_tab")
    rs(10).value = Request.Form.Item("show_email")
    rs(11).Value = Request.Form.Item("email_tab")
    rs(12).value = Request.Form.Item("show_size")
    rs(13).Value = Request.Form.Item("size_tab")
    rs(14).value = Request.Form.Item("show_bib")
    rs(15).Value = Request.Form.Item("bib_tab")
    rs(16).value = Request.Form.Item("show_provider")
    rs(17).Value = Request.Form.Item("provider_tab")
    rs(18).value = Request.Form.Item("show_cell")
    rs(19).Value = Request.Form.Item("cell_tab")
    rs(20).value = Request.Form.Item("show_teams")
    rs(21).Value = Request.Form.Item("teams_tab")
    rs(22).value = Request.Form.Item("show_fbook")
    rs(23).Value = Request.Form.Item("fbook_tab")
    rs(24).value = Request.Form.Item("show_twitter")
    rs(25).Value = Request.Form.Item("twitter_tab")
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_event") = "submit_event" Then
	lEventID = Request.Form.Item("events")
End If

If Cstr(lEventID) = vbNullString Then lEventID = 0

If CLng(lEventID) > 0 Then
    bInsert = True
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SElECT ShowAge, AgeTab, ShowDOB, DOBTab, ShowPhone, PhoneTab, ShowCity, CityTab, ShowSt, StTab, ShowEmail, EmailTab, ShowSize, SizeTab, "
    sql = sql & "ShowBib, BibTab, ShowProvider, ProviderTab, ShowCell, CellTab, ShowTeams, TeamsTab, ShowFbook, FBookTab, "
    sql = sql & "ShowTwitter, TwitterTab FROM PartEntryTabs WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        sShowAge = rs(0).value
        iAgeTab = rs(1).Value
        sShowDOB = rs(2).value
        iDOBTab = rs(3).Value
        sShowPhone = rs(4).value
        iPhoneTab = rs(5).Value
        sShowCity = rs(6).value
        iCityTab = rs(7).Value
        sShowSt = rs(8).value
        iStTab = rs(9).Value
        sShowEmail = rs(10).value
        iEmailTab = rs(11).Value
        sShowSize = rs(12).value
        iSizeTab = rs(13).Value
        sShowBib = rs(14).value
        iBibTab = rs(15).Value
        sShowProvider = rs(16).value
        iProviderTab = rs(17).Value
        sShowCell = rs(18).value
        iCellTab = rs(19).Value
        sShowTeams = rs(20).value
        iTeamsTab = rs(21).Value
        sShowFbook = rs(22).value
        iFbookTab = rs(23).Value
        sShowTwitter = rs(24).value
        iTwitterTab = rs(25).Value

        bInsert = False
    End If
    rs.Close
    Set rs = Nothing

    If bInsert = True Then
        sql = "INSERT INTO PartEntryTabs (EventID) VALUES (" & lEventID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        sShowAge = "y"
        iAgeTab = 4
        sShowDOB = "n"
        iDOBTab = 9
        sShowPhone = "n"
        iPhoneTab = 10
        sShowCity = "y"
        iCityTab = 5
        sShowSt = "y"
        iStTab = 6
        sShowEmail = "y"
        iEmailTab = 7
        sShowSize = "n"
        iSizeTab = 11
        sShowBib = "y"
        iBibTab = 8
        sShowProvider = "n"
        iProviderTab = 14
        sShowCell = "n"
        iCellTab = 15
        sShowTeams = "n"
        iTeamsTab = 16
        sShowFbook = "n"
        iFbookTab = 12
        sShowTwitter = "n"
        iTWitterTab = 13
    End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>Enter Participants</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<%If Session("role") = "admin" Then%>
            <!--#include file = "../includes/admin_menu.asp" -->
		    <div class="col-sm-10">
			    <h3 class="h3">Participant Data</h3>
			
			    <!--#include file = "../includes/event_nav.asp" -->
			    <!--#include file = "../admin/participants/part_nav.asp" -->
        <%Else%>
		    <!--#include file = "staff_menu.asp" -->
		    <div class="col-sm-10">
        <%End If%>
			    <h3 class="h3">Enter Participants Control Panel</h3>

			    <form class="form-inline" name="which_event" method="post" action="parts_events.asp?event_id=<%=lEventID%>">
			    <label for="events">Select Event:</label>
			    <select class="form-control" name="events" id="events" onchange="this.form.get_event.click()">
				    <option value="">&nbsp;</option>
				    <%For i = 0 to UBound(Events, 2) - 1%>
					    <%If CLng(lEventID) = CLng(Events(0, i)) Then%>
						    <option value="<%=Events(0, i)%>" selected><%=Events(1, i)%> (On <%=Events(2, i)%> in <%=Events(3, i)%>)</option>
					    <%Else%>
						    <option value="<%=Events(0, i)%>"><%=Events(1, i)%> (On <%=Events(2, i)%> in <%=Events(3, i)%>)</option>
					    <%End If%>
				    <%Next%>
			    </select>
			    <input type="hidden" name="submit_event" id="submit_event" value="submit_event">
			    <input type="submit" class="form-control" name="get_event" id="get_event" value="Get This Event" style="font-size:0.8em;">
			    </form>
            </div>

            <%If Not CLng(lEventID) = 0 Then%>
                <div class="col-sm-10">
                    <h4 class="h4">Fields and Tab Order</h4>

                    <div class="col-md-5">
                        <h5 class="h5">Fixed Fields</h5>
                        <table class="table table-striped">
                            <tr>
                                <th style="text-align:left;">Field</th>
                                <th>Show</th>
                                <th>Tab Order</th>
                            </tr>
                            <tr>
                                <td style="text-align: left;">First Name</td>
                                <td>Yes</td>
                                <td>0</td>
                            </tr>
                            <tr>
                                <td style="text-align: left;">Last Name</td>
                                <td>Yes</td>
                                <td>1</td>
                            </tr>
                            <tr>
                                <td style="text-align: left;">Gender</td>
                                <td>Yes</td>
                                <td>2</td>
                            </tr>
                            <tr>
                                <td style="text-align: left;">Races</td>
                                <td>Yes</td>
                                <td>3</td>
                            </tr>
                        </table>
                    </div>
                    <div class="col-sm-5">
                        <h5 class="h5">Select Fields and Set Tab Order</h5>
                        <form class="form" name="set_tabs" method="post" action="parts_events.asp?event_id=<%=lEventID%>">
                        <table class="table table-striped">
                            <tr>
                                <th>Field</th>
                                <th>Show</th>
                                <th>Tab Order</th>
                            </tr>
                            <tr>
                                <td>Age</td>
                                <td>
                                    <select class="form-control" name="show_age" id="show_age">
                                        <%If sShowAge = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="age_tab" id="age_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iAgeTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                                <tr>
                                <td>City</td>
                                <td>
                                    <select class="form-control" name="show_city" id="show_city">
                                        <%If sShowCity = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="city_tab" id="city_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iCityTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>St</td>
                                <td>
                                    <select class="form-control" name="show_st" id="show_st">
                                        <%If sShowSt = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="st_tab" id="st_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iStTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Email</td>
                                <td>
                                    <select class="form-control" name="show_email" id="show_email">
                                        <%If sShowEmail = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="email_tab" id="email_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iEmailTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Bib</td>
                                <td>
                                    <select class="form-control" name="show_bib" id="show_bib">
                                        <%If sShowBib = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="bib_tab" id="bib_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iBibTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>DOB</td>
                                <td>
                                    <select class="form-control" name="show_dob" id="show_dob">
                                        <%If sShowDOB = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="dob_tab" id="dob_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iDOBTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Phone</td>
                                <td>
                                    <select class="form-control" name="show_phone" id="show_phone">
                                        <%If sShowPhone = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="phone_tab" id="phone_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iPhoneTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Size</td>
                                <td>
                                    <select class="form-control" name="show_size" id="show_size">
                                        <%If sShowSize = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="size_tab" id="size_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iSizeTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Facebook</td>
                                <td>
                                    <select class="form-control" name="show_fbook" id="show_fbook">
                                        <%If sShowFbook = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="fbook_tab" id="fbook_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iFbookTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Twitter</td>
                                <td>
                                    <select class="form-control" name="show_twitter" id="show_twitter">
                                        <%If sShowTwitter = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="twitter_tab" id="twitter_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iTwitterTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Cell Provider</td>
                                <td>
                                    <select class="form-control" name="show_provider" id="show_provider">
                                        <%If sShowProvider = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td >
                                    <select class="form-control" name="provider_tab" id="provider_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iProviderTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Cell Number</td>
                                <td>
                                    <select class="form-control" name="show_cell" id="show_cell">
                                        <%If sShowCell = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="cell_tab" id="cell_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iCellTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td>Teams</td>
                                <td>
                                    <select class="form-control" name="show_teams" id="show_teams">
                                        <%If sShowTeams = "n" Then%>
                                            <option value="n">No</option>
                                            <option value="y">Yes</option>
                                        <%Else%>
                                            <option value="n">No</option>
                                            <option value="y" selected>Yes</option>
                                        <%End If%>
                                    </select>
                                </td>
                                <td>
                                    <select class="form-control" name="teams_tab" id="teams_tab">
                                        <%For i = 4 To 16%>
                                            <%If CInt(iTeamsTab) = CInt(i) Then%>
                                                <option value="<%=i%>" selected><%=i%></option>
                                            <%Else%>
                                                <option value="<%=i%>"><%=i%></option>
                                            <%End If%>
                                        <%Next%>
                                    </select>
                                </td>
                            </tr>
                         </table>
                        <div class="form-group">
			                <input type="hidden" name="submit_tabs" id="submit_tabs" value="submit_tabs">
			                <input type="submit" class="form-control" name="set_tabs" id="set_tabs" value="Set Visibility and Tabs">
                        </div>
                        </form>
                    </div>
                    <div class="col-sm-2 bg-success">
                        <a href="javascript:pop('enter_parts.asp?event_id=<%=lEventID%>',900,700)">Enter Participants</a>
                    </div>
                </div>
            <%End If%>
		</div>
	</div>
</div>
<!--#include file = "../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>