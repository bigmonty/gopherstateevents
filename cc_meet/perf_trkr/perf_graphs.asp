<%@ Language=VBScript %>

<%
Option Explicit

Dim conn, rs, sql, conn2, rs2, sql2
Dim i
Dim lPacksID, lRosterID
Dim sPackName, sSport, sGender
Dim MyPacks(), PackMmbrs(), ViewPerf()
Dim dWhenCreated

If CStr(Session("perf_trkr_id")) = vbNullString Then Response.Redirect "/default.asp?sign_out=y"

lPacksID = Request.QueryString("packs_id")
lRosterID = Request.QueryString("roster_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

ReDim ViewPerf(0)

If Request.Form.Item("submit_select") = "submit_select" Then
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT PTPackMmbrsID FROM PTPackMmbrs WHERE PerfTrkrPacksID = " & lPacksID
    rs.Open sql, conn2, 1, 2
    Do While Not rs.EOF
        If Request.Form.Item("remove_" & rs(0).Value) = "on" Then
            ViewPerf(i) = rs(0).Value
            i = i + 1
            ReDim Preserve ViewPerf(i)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_pack") = "submit_pack" Then
    lPacksID = Request.Form.Item("my_packs")
End If

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

If CStr(lPacksID) = vbNullString Then lPacksID = 0
If lPacksID = "0" Then lPacksID = MyPacks(0, 0)

If CStr(lRosterID) = vbNullString Then lRosterID = 0

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
End If

For i = 0 To UBound(ViewPerf) - 1
Next

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
<title>GSE Performance Tracker Performance Graphs</title>
</head>

<body>
<div class="container">
    <!--#include file = "perf_trkr_hdr.asp" -->
    <!--#include file = "perf_trkr_nav.asp" -->

    <div class="row">
         <div class="col-sm-10">
           <h4 class="h4">My Performance Tracker Graphs</h4>

            <div>
                <form class="form-inline" name="get_pack" method="post" action="perf_graphs.asp">
                <label for="my_packs">Select a Pack:</label>&nbsp;
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
	            <input type="hidden" name="submit_pack" id="submit_pack" value="submit_pack">
	            <input class="form-control" type="submit" name="submit1" id="submit1" value="Get Pack">
                </form>
            </div>

            <%If Not CLng(lPacksID) = 0 Then%>
                <h5 class="h5"><%=sPackName%> <%=sSport%>&nbsp;&nbsp;<%=sGender%> <span style="font-weight: normal;">(created <%=dWhenCreated%>)</span></h5>

                <div class="row">
                    <div class="col-sm-3">
                        <h5 class="h5 bg-warning" style="padding:2px;">Pack Members</h5>

                        <form name="select_mmbr" method="post" action="perf_graphs.asp?packs_id=<%=lPacksID%>">
                        <ul class="list-group">
                            <%For i = 0 To UBOund(PackMmbrs, 2) - 1%>
                                <li class="list-group-item">
                                    <input type="checkbox" name="select_<%=PackMmbrs(0, i)%>" id="select_<%=PackMmbrs(0, i)%>">
                                    <%=PackMmbrs(2, i)%>&nbsp;(<%=PackMmbrs(3, i)%>)
                                </li>
                            <%Next%>
                        </ul>
 	                    <input type="hidden" name="submit_select" id="submit_select" value="submit_select">
	                    <input class="form-control" type="submit" name="submit2" id="submit2" value="Select Participant(s)">
                        </form>
                    </div>
	                <div class="col-sm-9">
                        <h5 class="h5 bg-danger" style="padding:2px;">Graph</h5>

                        (This utility is still being developed.  We appreciate your patience!)
                        <div class="embed-responsive embed-responsive-16by9">
                            <iframe class="embed-responsive-item" src="show_graph.asp"></iframe>
                        </div>
                    </div>
                </div>
            <%End If%>
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
