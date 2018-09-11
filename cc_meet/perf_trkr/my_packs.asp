<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i, j
Dim lPacksID, lThisTeam
Dim sPackName, sSport, sGender
Dim MyPacks(), Teams(), Roster(), PackMmbrs(), Remove()
Dim dWhenCreated

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lPacksID = Request.QueryString("packs_id")
lThisTeam = Request.QueryString("this_team")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_remove") = "submit_remove" Then
    i = 0
    ReDim Remove(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PTPackMmbrsID FROM PTPackMmbrs WHERE PerfTrkrPacksID = " & lPacksID
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("remove_" & rs(0).Value) = "on" Then
            Remove(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Remove(i)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Remove) - 1
        sql = "DELETE FROM PTPackMmbrs WHERE PTPackMmbrsID = " & Remove(i)
        Set rs = conn2.Execute(sql)
        Set rs = Nothing
    Next
ElseIf Request.Form.Item("submit_roster") = "submit_roster" Then
    Dim sNewMmbrs
    Dim NewMmbrs()

	sNewMmbrs = Request.Form.Item("roster")

	ReDim NewMmbrs(0)
	
	If Not CStr(sNewMmbrs) = vbNullString Then
		j = 0
		For i = 1 To Len(sNewMmbrs)
			If Mid(sNewMmbrs, i, 1) = "," Then
				NewMmbrs(j) = Trim(CStr(NewMmbrs(j)))
				j = j + 1
				ReDim Preserve NewMmbrs(j)
			Else
				NewMmbrs(j) = NewMmbrs(j) & Mid(sNewMmbrs, i, 1)
			End If
		Next
	
		For i = 0 To UBound(NewMmbrs)
			sql = "INSERT INTO PTPackMmbrs(PerfTrkrPacksID, RosterID, WhenAdded) VALUES (" & lPacksID & ", " & NewMmbrs(i) & ", '" & Now() & "')"
			Set rs = conn2.Execute(sql)
			Set rs = Nothing
		Next
	End If
ElseIf Request.Form.Item("submit_team") = "submit_team" Then
    lThisTeam = Request.Form.Item("teams")
ElseIf Request.Form.Item("submit_pack") = "submit_pack" Then
    lPacksID = Request.Form.Item("my_packs")
End If

If CStr(lPacksID) = vbNullString Then lPacksID = 0
If CStr(lThisTeam) = vbNullString Then lThisTeam = 0

i = 0
ReDim Preserve MyPacks(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT PerfTrkrPacksID, PackName FROM PerfTrkrPacks WHERE PerfTrkrID = " & Session("perf_trkr_id")
rs.Open sql, conn2, 1, 2
Do While Not rs.EOF
    MyPacks(0, i) = rs(0).Value
    MyPacks(1, i) = Replace(rs(1).Value, "''", "'")
    i = i + 1
    ReDim Preserve MyPacks(1, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Not CLng(lPacksID) = 0 Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PackName, Sport, Gender, WhenCreated FROM PerfTrkrPacks WHERE PerfTrkrPacksID = " & lPacksID
    rs.Open sql, conn2, 1, 2
    sPackName = Replace(rs(0).Value, "''", "'")
    sSport = rs(1).Value
    sGender = rs(2).Value
    dWhenCreated = rs(3).Value
    rs.Close
    Set rs = Nothing

    i = 0
    ReDim PackMmbrs(3, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT pt.PTPackMmbrsID, pt.RosterID, r.FirstName, r.LastName, r.TeamsID FROM PTPackMmbrs pt INNER JOIN Roster r ON pt.RosterID = r.RosterID "
    sql = sql & "WHERE pt.PerfTrkrPacksID = " & lPacksID & " ORDER BY r.LastName, r.FirstName"
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        PackMmbrs(0, i) = rs(0).Value
        PackMmbrs(1, i) = rs(1).Value
        PackMmbrs(2, i) = Replace(rs(3).Value, "''", "'") & ", " & Replace(rs(2).Value, "''", "'")
        PackMmbrs(3, i) = GetTeamName(rs(4).Value)
        i = i + 1
        ReDim Preserve PackMmbrs(3, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    i = 0
    ReDim Teams(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamsID, TeamName FROM Teams WHERE Sport = '" & sSport & "' AND Gender = '" & sGender & "' ORDER BY TeamName"
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        Teams(0, i) = rs(0).Value
        Teams(1, i) = Replace(rs(1).Value, "''", "'")
        i = i + 1
        ReDim Preserve Teams(1, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If Not CLng(lThisTeam) = 0 Then
        i = 0
        ReDim Roster(1, 0)
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT RosterID, FirstName, LastName FROM Roster WHERE TeamsID = " & lThisTeam & " ORDER BY LastName, FirstName"
        rs.Open sql, conn2, 1, 2
        Do While Not rs.EOF
            If UBound(PackMmbrs, 2) = 0 Then
                Roster(0, i) = rs(0).Value
                Roster(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
                i = i + 1
                ReDim Preserve Roster(1, i)
            Else
                For j = 0 To UBound(PackMmbrs, 2) - 1
                    If CLng(rs(0).Value) = CLng(PackMmbrs(1, j)) Then
                        Exit For
                    Else
                        If j = UBound(PackMmbrs, 2) - 1 Then
                            Roster(0, i) = rs(0).Value
                            Roster(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
                            i = i + 1
                            ReDim Preserve Roster(1, i)
                        End If
                    End If
                Next
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
End If

Private Function GetTeamName(lTeamID)
    Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT TeamName FROM Teams WHERE TeamsID = " & lTeamID
    rs2.Open sql2, conn2, 1, 2
    GetTeamName = Replace(rs2(0).Value, "''", "'")
    rs2.Close
    Set rs2 = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>My GSE Performance Tracker Packs</title>
</head>

<body>
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->
    <!--#include file = "perf_trkr_nav.asp" -->

    <div class="row">
         <div class="col-sm-10">
            <h4 class="h4">My Packs</h4>

           <form class="form-inline" name="get_pack" method="post" action="my_packs.asp">
            <div class="form-group">
                <label for="pack_name">Select a Pack:</label>
                <select class="form-control" name="my_packs" id="my_packs" onchange="this.form.submit1.click();">
                    <option value="">&nbsp;</option>
                    <%For i = 0 To UBound(MyPacks, 2) - 1%>
                        <%If CLng(lPacksID) = CLng(MyPacks(0, i)) Then%>
                            <option value="<%=MyPacks(0, i)%>" selected><%=MyPacks(1, i)%></option>
                        <%Else%>
                            <option value="<%=MyPacks(0, i)%>"><%=MyPacks(1, i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
            </div>
            <div class="form-group">
                <input type="hidden" name="submit_pack" id="submit_pack" value="submit_pack">
                <input class="form-control" type="submit" name="submit1" id="submit1" value="Get Pack">
            </div>
            </form>
            <%If Not CLng(lPacksID) = 0 Then%>
                <br>
                <h5 class="h5"><%=sPackName%> <%=sSport%>&nbsp;-&nbsp;<%=sGender%> <span style="font-weight: normal;">(created <%=dWhenCreated%>)</span></h5>

                <div class="row">
                    <div class="col-sm-6">
                        <h5 class="h5 bg-warning" style="padding:2px;">Add Members</h5>
 
                       <form class="form-inline" name="get_team" method="post" action="my_packs.asp?packs_id=<%=lPacksID%>">
                        <div class="form-group">
                            <label for="teams">Team/School:</label>
                            <select class="form-control" name="teams" id="teams" onchange="this.form.submit2.click();">
                                <option value="">&nbsp;</option>
                                <%For i = 0 To UBound(Teams, 2) - 1%>
                                    <%If CLng(lThisTeam) = CLng(Teams(0, i)) Then%>
                                        <option value="<%=Teams(0, i)%>" selected><%=Teams(1, i)%></option>
                                    <%Else%>
                                        <option value="<%=Teams(0, i)%>"><%=Teams(1, i)%></option>
                                    <%End If%>
                                <%Next%>
                            </select>
                        </div>
                        <div class="form-group">
 	                        <input type="hidden" name="submit_team" id="submit_team" value="submit_team">
	                        <input class="form-control" type="submit" name="submit2" id="submit2" value="Get Team">
                        </div>
                        </form>

                        <hr>

                        <%If Not CLng(lThisTeam) = 0 Then%>
                            <form class="form" name="get_part" method="post" action="my_packs.asp?packs_id=<%=lPacksID%>&amp;this_team=<%=lThisTeam%>">
                            <div class="form-group">
                                <label for="roster">Roster:</label>
                                <select class="form-control" name="roster" id="roster" size="15" multiple>
                                    <%For i = 0 To UBound(Roster, 2) - 1%>
                                        <option value="<%=Roster(0, i)%>"><%=Roster(1, i)%></option>
                                    <%Next%>
                                </select>
                            </div>
                            <div class="form-group">
 	                            <input type="hidden" name="submit_roster" id="submit_roster" value="submit_roster">
	                            <input class="form-control" type="submit" name="submit3" id="submit3" value="Get Participant(s)">
                            </div>
                            </form>
                        <%End If%>
                    </div>
                    <div class="col-sm-6">   
                        <h5 class="h5 bg-danger" style="padding:2px;">Existing Members <span style="font-weight: normal;">(use checkbox to remove)</span></h5>

                        <form class="form" name="remove_mmbr" method="post" action="my_packs.asp?packs_id=<%=lPacksID%>&amp;this_team=<%=lThisTeam%>">
                        <ol class="list-group">
                            <%For i = 0 To UBound(PackMmbrs, 2) - 1%>
                                <li class="list-group-item">
                                    <input type="checkbox" name="remove_<%=PackMmbrs(0, i)%>" id="remove_<%=PackMmbrs(0, i)%>">
                                    <%=PackMmbrs(2, i)%>&nbsp;(<%=PackMmbrs(3, i)%>)
                                </li>
                            <%Next%>
                        </ol>
                        <div class="form-group">
 	                        <input type="hidden" name="submit_remove" id="submit_remove" value="submit_remove">
	                        <input class="form-control" type="submit" name="submit4" id="submit4" value="Remove Participant(s)">
                        </div>
                        </form>
                    </div>
                <%End If%>
            </div>
        </div>
        <div class="col-sm-2">
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
