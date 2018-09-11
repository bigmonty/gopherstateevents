<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lngTtlRcds, lPartToKeep, lUserToGet, lSeriesID
Dim sWhichAlpha, sDeleteNoRace, sQuickCondense
Dim PartList, Fields(6), MatchList(), ChangeTable1(6), ChangeTable2(2), CondenseThese(), DeleteThese(), Series
Dim bMatch

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Server.ScriptTimeout = 1200

sWhichAlpha = Request.QueryString("which_alpha")
lUserToGet = Request.QueryString("user_to_get")
lSeriesID = Request.QueryString("series_id")

sDeleteNoRace = "n"
sDeleteNoRace = Request.QueryString("delete_norace")

sQuickCondense = Request.QueryString("quick_condense")
If sQuickCondense = vbNullString Then sQuickCondense = "n"

Response.Buffer = False		'Turn buffering off
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

i = 0
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT SeriesID, SeriesName, SeriesYear FROM Series ORDER BY SeriesYear DESC, SeriesName"
rs.Open sql, conn, 1, 2
Series = rs.GetRows()
rs.Close
Set rs = Nothing

Fields(0) =  "Part ID"
Fields(1) =  "First Name"
Fields(2) =  "Last Name"
Fields(3) =  "City"
Fields(4) =  "Phone"
Fields(5) =  "Email"
Fields(6) =  "DOB"

If Request.Form.Item("submit_series") = "submit_series" Then
    lSeriesID = Request.Form.Item("series")
ElseIf Request.Form.Item("submit_batch2") = "submit_batch2" Then    'if I want to use this I need to take the 2 out of submit_batch2
    'if we ever add this back in we need to add the script to delete multiples from series parts
    Call GetPartList()

    'get match list
    j = 0
    ReDim MatchList(0)
    For i = 1 To 6
        If Request.Form.Item("match_" & i) = "on" Then
            MatchList(j) = i
            j = j + 1
            ReDim Preserve MatchList(j)
        End If
    Next

    If UBound(MatchList) > 0 Then
        i = 0
        j = 1
        Do While i < UBound(PartList, 2)
            bMatch = True

            Do While j <= UBound(PartList, 2)
                For k = 0 To UBound(MatchList) - 1
                    If Len(Trim(PartList(MatchList(k), i))) <> Len(Trim(PartList(MatchList(k), j))) Then
                        bMatch = False
                        Exit For
                    ElseIf Trim(UCase(PartList(MatchList(k), i))) <> Trim(UCase(PartList(MatchList(k), j))) Then 
                        bMatch = False
                        Exit For
                    End If
                Next

                If bMatch = True Then
                    Call CondenseParts(PartList(0, j), PartList(0, i))

                    j = j + 1
                Else
                    i = i + 1
                    j = i + 1
                    Exit Do
                End If
            Loop
        Loop
    End If
ElseIf Request.Form.Item("submit_manual") = "submit_manual" Then
    Call GetPartList()

    'get items to delete
    j = 0
    ReDim DeleteThese(0)
    For i = 0 To UBound(PartList, 2)
        If Request.Form.Item("action_" & PartList(0, i)) = "delete" Then
            DeleteThese(j) = PartList(0, i)
            j = j + 1
            ReDim Preserve DeleteThese(j)
        End If
    Next 

    For i = 0 To UBound(DeleteThese) - 1
        Call DeleteParts(DeleteThese(i))
    Next

    'get participant to merge to
    lPartToKeep = 0
    For i = 0 To UBound(PartList, 2)
        If Request.Form.Item("action_" & PartList(0, i)) = "keep" Then
            lPartToKeep = PartList(0, i)
            Exit For
        End If
    Next 

    If CLng(lPartToKeep) > 0 Then
        'get participant(s) to condense
        j = 0
        ReDim CondenseThese(0)
        For i = 0 To UBound(PartList, 2)
            If Request.Form.Item("action_" & PartList(0, i)) = "condense" Then
                CondenseThese(j) = PartList(0, i)
                j = j + 1
                ReDim Preserve CondenseThese(j)
            End If
        Next 

        For i = 0 To UBound(CondenseThese) - 1
            Call CondenseParts(CondenseThese(i), lPartToKeep)

            'removed duplicates from series parts
            If CLng(lSeriesID) > 0 Then
                j = 0
                k = 0
                ReDim DeleteThese(0)
                Set rs = Server.CreateObject("ADODB.Recordset")
                sql = "SELECT SeriesPartsID FROM SeriesParts WHERE ParticipantID = " & lPartToKeep & " AND SeriesID = " & lSeriesID
                rs.Open sql, conn, 1, 2
                Do While Not rs.EOF 
                    If j > 0 Then   'don't delete the first one
                        DeleteThese(k) = rs(0).Value
                        k = k + 1
                        ReDim Preserve DeleteThese(k)
                    End If
                    j = j + 1
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing

                For k = 0 To UBound(DeleteThese) - 1
                    sql = "DELETE FROM SeriesParts WHERE SeriesPartsID = " & DeleteThese(k)
                    Set rs = conn.Execute(sql)
                    Set rs = Nothing
                Next
            End If
        Next

        sDeleteNoRace = "y"

        'delete from participant
'        sql = "DELETE FROM Participant WHERE ParticipantID = " & lPartFrom
'        Set rs = conn.Execute(sql)
'        Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_user_id") = "submit_user_id" Then
    lUserToGet = Request.Form.Item("user_to_get")
ElseIf Request.Form.Item("submit_alpha") = "submit_alpha" Then
    sWhichAlpha = Request.Form.Item("which_alpha")
End If

If CStr(lUserToGet) = vbNullString Then lUserToGet = 0
If CStr(lSeriesID) = vbNullString Then lSeriesID = 0

If sQuickCondense = "y" Then
    If Not sWhichAlpha = vbNullString Then      'prevent doing the whole thing at once...want to be able to see that it is working
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, p.City, p.Country FROM Participant p INNER JOIN SeriesParts sp "
        sql = sql & "ON p.ParticipantID = sp.ParticipantID WHERE sp.SeriesID = " & lSeriesID & " AND p.LastName Like '" & sWhichAlpha & "%'"
        sql = sql & " ORDER BY p.Gender, p.LastName, p.FirstName, p.ParticipantID DESC"
        rs.Open sql, conn, 1, 2
        PartList = rs.GetRows()     'country is just to hold a space for num races
        rs.Close
        Set rs = Nothing

        For i = 0 To UBound(PartList, 2)
            PartList(4, i) = NumRaces(PartList(0, i))
        Next

        For i = 0 To UBound(PartList, 2)
            If CInt(PartList(4, i)) = 0 Then 
                Call DeleteParts(PartList(0, i))    'delete those with no races
            End If
        Next

        Set rs= Server.CreateObject("ADODB.Recordset")
        sql = "SELECT p.ParticipantID, p.FirstName, p.LastName, p.City FROM Participant p INNER JOIN SeriesParts sp "
        sql = sql & "ON p.ParticipantID = sp.ParticipantID WHERE sp.SeriesID = " & lSeriesID & " AND p.LastName Like '" & sWhichAlpha & "%'"
        sql = sql & " ORDER BY p.Gender, p.LastName, p.FirstName, p.ParticipantID DESC"
        rs.Open sql, conn, 1, 2
        PartList = rs.GetRows()
        rs.Close
        Set rs = Nothing

        ReDim MatchList(2)
        MatchList(0) = 1
        MatchList(1) = 2
        MatchList(2) = 3

        If UBound(MatchList) > 0 Then
            i = 0
            j = 1
            Do While i <= UBound(PartList, 2) - 1
                bMatch = True

                Do While j <= UBound(PartList, 2)
                    For k = 0 To 2
                        If Len(Trim(PartList(MatchList(k), i))) <> Len(Trim(PartList(MatchList(k), j))) Then
                            bMatch = False
                            Exit For
                        ElseIf Trim(UCase(PartList(MatchList(k), i))) <> Trim(UCase(PartList(MatchList(k), j))) Then 
                            bMatch = False
                            Exit For
                        End If
                    Next

                    If bMatch = True Then
                        Call CondenseParts(PartList(0, j), PartList(0, i))

                        'delete other from series parts

                        j = j + 1
                    Else
                        i = i + 1
                        j = i + 1
                        Exit Do
                    End If
                Loop
            Loop
        End If
    End If
End If

Call GetPartList()

If sDeleteNoRace = "y" Then
    For i = 0 To UBound(PartList, 2)
        If CInt(PartList(7, i)) = 0 Then 
            Call DeleteParts(PartList(0, i))
        End If
    Next

    Call GetPartList()
End If

Set rs= Server.CreateObject("ADODB.Recordset")
sql = "SELECT ParticipantID FROM Participant"
rs.Open sql, conn, 1, 2
lngTtlRcds = rs.RecordCount
rs.Close
Set rs = Nothing

Private Sub GetPartList()
    Dim x

    If CLng(lSeriesID) = 0 And sWhichAlpha = vbNullString Then sWhichAlpha = "ab"

    Set rs= Server.CreateObject("ADODB.Recordset")
    If CLng(lUserToGet) = 0 Then
        If CLng(lSeriesID) = 0 Then
            sql = "SELECT ParticipantID, FirstName, LastName, City, Phone, Email, DOB, Country, SendInfo FROM Participant WHERE LastName Like '" 
            sql = sql & sWhichAlpha & "%' ORDER BY Gender, LastName, FirstName, ParticipantID DESC"
        Else
            sql = "SELECT sp.ParticipantID, p.FirstName, p.LastName, p.City, p.Phone, p.Email, p.DOB, p.Country, p.SendInfo FROM Participant p INNER JOIN "
            sql = sql & "SeriesParts sp ON p.ParticipantID = sp.ParticipantID WHERE p.LastName Like '" & sWhichAlpha 
            sql = sql & "%' AND sp.SeriesID = " & lSeriesID & " ORDER BY p.Gender, p.LastName, p.FirstName, p.ParticipantID DESC"
        End If
    Else
        If CLng(lSeriesID) = 0 Then
            sql = "SELECT ParticipantID, FirstName, LastName, City, Phone, Email, DOB, Country, SendInfo FROM Participant WHERE ParticipantID = " 
            sql = sql & lUserToGet & " ORDER BY Gender, LastName, FirstName, ParticipantID DESC"
        Else
            sql = "SELECT sp.ParticipantID, p.FirstName, p.LastName, p.City, p.Phone, p.Email, p.DOB, p.Country, p.SendInfo FROM Participant p INNER JOIN "
            sql = sql & " SeriesParts sp ON p.ParticipantID = sp.ParticipantID WHERE p.ParticipantID = " & lUserToGet & " AND sp.SeriesID = "
            sql = sql & lSeriesID & " ORDER BY p.Gender, p.LastName, p.FirstName, p.ParticipantID DESC"
        End If
    End If
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        PartList = rs.GetRows()
    Else
        ReDim PartList(8, 0)
    End If
    rs.Close
    Set rs = Nothing

    If UBound(PartList, 2) > 0 Then
        For x = 0 To UBound(PartList, 2)
            PartList(7, x) = NumRaces(PartList(0, x))
            PartList(8, x) = GetAge(PartList(0, x))
        Next
    End If
End Sub

Private Sub CondenseParts(lPartFrom, lPartTo)
    If Not CLng(lPartFrom) = CLng(lPartTo) Then
        Dim x

        ChangeTable1(0) = "IndResults"
        ChangeTable1(1) = "PartRace"
        ChangeTable1(2) = "PartReg"
        ChangeTable1(3) = "PreRaceRecips"
        ChangeTable1(4) = "ResultsSent"
        ChangeTable1(5) = "TeamMmbrs"
        ChangeTable1(6) = "SeriesParts"

        For x = 0 To 6
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT ParticipantID FROM " & ChangeTable1(x) & " WHERE ParticipantID = " & lPartFrom
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF 
                rs(0).Value = lPartTo
                rs.Update
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Next

        ChangeTable2(0) = "PartReminders"
        ChangeTable2(1) = "Records"
        ChangeTable2(2) = "Splits"

        For x = 0 To 2
            Set rs = Server.CreateObject("ADODB.Recordset")
            sql = "SELECT PartID FROM " & ChangeTable2(x) & " WHERE PartID = " & lPartFrom
            rs.Open sql, conn, 1, 2
            Do While Not rs.EOF 
                rs(0).Value = lPartTo
                rs.Update
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        Next
    End If
End Sub

Private Sub DeleteParts(lDeleteMe)
    Dim x

    ChangeTable1(0) = "IndResults"
    ChangeTable1(1) = "PartRace"
    ChangeTable1(2) = "PartReg"
    ChangeTable1(3) = "PreRaceRecips"
    ChangeTable1(4) = "ResultsSent"
    ChangeTable1(5) = "TeamMmbrs"
    ChangeTable1(6) = "SeriesParts"

    For x = 0 To 6
        sql = "DELETE FROM " & ChangeTable1(x) & " WHERE ParticipantID = " & lDeleteMe
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next

    ChangeTable2(0) = "PartReminders"
    ChangeTable2(1) = "Records"
    ChangeTable2(2) = "Splits"

    For x = 0 To 2
        sql = "DELETE FROM " & ChangeTable2(x) & " WHERE PartID = " & lDeleteMe
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next

    'delete from participant
    sql = "DELETE FROM Participant WHERE ParticipantID = " & lDeleteMe
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End Sub

Private Function NumRaces(lPartID)
    NumRaces = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM PartRace WHERE ParticipantID = " & lPartID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        NumRaces = CInt(NumRaces) + 1
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function

Private Function GetAge(lPartID)
    GetAge = 0

    Set rs = Server.CreateObject("ADODB.Recordset")
    If CLng(lSeriesID) = 0 Then
        sql = "SELECT Age FROM PartRace WHERE ParticipantID = " & lPartID & " ORDER BY Age DESC"
    Else
        sql = "SELECT Age FROM SeriesParts WHERE ParticipantID = " & lPartID & " AND SeriesID = " & lSeriesID
    End If
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then GetAge = rs(0).Value
    rs.Close
    Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Admin Condense Participants Utility</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">Condense Users Utility</h4>

            <div class="row">
                <div class="col-sm-6">
                    <span style="font-weight: bold;">Total Records: <%=lngTtlRcds%></span>
                </div>
                <div class="col-sm-6" style="text-align:right;">
                    <ul class="list-inline">
                        <%If CLng(lSeriesID) > 0 Then%>
                            <li><a href="condense_users.asp?quick_condense=y&amp;series_id=<%=lSeriesID%>&which_alpha=<%=sWhichAlpha%>">Quick Condense</a></li>
                        <%End If%>
                        <li><a href="condense_users.asp?which_alpha=<%=sWhichAlpha%>&amp;delete_norace=y&amp;series_id=<%=lSeriesID%>">Delete "No Race"</a></li>
                    </ul>
                </div>
            </div>

            <div class="row">
                <div class="bg-danger">
                    <form class="form-inline" name="get_series" method="post" action="condense_users.asp?which_alpha=<%=sWhichAlpha%>">
                    <label for="series">Series Filter:</label>
                    <select class="form-control"name="series" id="series" onchange="this.form.submit1x.click();">
                        <option value="">&nbsp;</option>
                        <%For i = 0 To UBound(Series, 2) - 1%>
                            <%If CLng(lSeriesID) = CLng(Series(0, i)) Then%>
                                <option value="<%=Series(0, i)%>" selected><%=Series(1, i)%> (<%=Series(2, i)%>)</option>
                            <%Else%>
                                <option value="<%=Series(0, i)%>"><%=Series(1, i)%> (<%=Series(2, i)%>)</option>
                            <%End If%>
                        <%Next%>
                    </select>
			        <input type="hidden" name="submit_series" id="submit_series" value="submit_series">
			        <input type="submit" class="form-control" name="submit1x" id="submit1x" value="Get This Series">
                    </form>
                </div>
            </div>

            <div class="row">
                <div class="col-sm-6 bg-info">
                    <form class="form-inline" name="condense" method="post" action="condense_users.asp?series_id=<%=lSeriesID%>">
                    <label for="which_alpha">Alpha Filter:</label>
                    <input type="text" class="form-control" name="which_alpha" id="which_alpha" value="<%=sWhichAlpha%>">
			        <input type="hidden" name="submit_alpha" id="submit_alpha" value="submit_alpha">
			        <input type="submit" class="form-control" name="submit1" id="submit1" value="Get These">
                    </form>
                </div>
                <div class="col-sm-6 bg-warning">
                    <form class="form-inline" name="condense" method="post" action="condense_users.asp?series_id=<%=lSeriesID%>">
                    <label for="user_to_get">User ID Filter:</label>
                    <input type="text" class="form-control" name="user_to_get" id="user_to_get" value="<%=lUserToGet%>">
			        <input type="hidden" name="submit_user_id" id="submit_user_id" value="submit_user_id">
			        <input type="submit" class="form-control" name="submit3" id="submit3" value="Get These">
                    </form>
                </div>
            </div>

            <!--  I don't think I want to allow this but not sure I want to delete it yet
            <div class="row bg-success">
                <form class="form-inline" name="condense_batch" method="post" action="condense_users.asp?which_alpha=<%=sWhichAlpha%>&amp;series_id=<%=lSeriesID%>">
                <label>Condense Batch:  Match On </label>
                <%For i = 1 To UBound(Fields)%>
                    <input type="checkbox" name="match_<%=i%>" id="match_<%=i%>" checked><%=Fields(i)%>&nbsp;&nbsp;
                <%Next%>
 			    <input type="hidden" name="submit_batch" id="submit_batch" value="submit_batch">
			    <input type="submit" class="form-control" name="submit2" id="submit2" value="Condense Batch">
                </form>
            </div>
            -->

            <div class="row">
                <div class="col-md-10">
                    <h4 class="h4">Condense Pool:</h4>
                </div>
                <div class="col-md-2">
                    <h5 class="h5">List Total:&nbsp;<%=UBound(PartList, 2)%></h5>
                </div>
            </div>

            <div class="row">
                <form class="form" name="condense_manually" method="post" action="condense_users.asp?which_alpha=<%=sWhichAlpha%>&amp;series_id=<%=lSeriesID%>">
                <table class="table table-striped table-condensed table-responsive">
                    <tr>
                        <td colspan="11">
 			                <input type="hidden" name="submit_manual" id="submit_manual" value="submit_manual">
			                <input type="submit" class="form-control" name="submit3" id="submit3" value="Condense Selected">
                        </td>
                    </tr>
                <tr>
                    <th>ID</th>
                    <th>First</th>
                    <th>Last</th>
                    <th>Age</th>
                    <th>City</th>
                    <th>Phone</th>
                    <th>Email</th>
                    <th>DOB</th>
                    <th>Races</th>
                    <th>Action</th>
                </tr>
                <%For i = 0 To UBound(PartList, 2)%>
                    <tr>
                        <td><a href="javascript:pop('user_hist.asp?part_id=<%=PartList(0, i)%>',300,700)"><%=PartList(0, i)%></a></td>
                        <td><%=PartList(1, i)%></td>
                        <td><%=PartList(2, i)%></td>
                        <td><%=PartList(8, i)%></td>
                        <td><%=PartList(3, i)%></td>
                        <td><%=PartList(4, i)%></td>
                        <td><%=PartList(5, i)%></td>
                        <td><%=PartList(6, i)%></td>
                        <td><%=PartList(7, i)%></td>
                        <td>
                            <select name="action_<%=PartList(0, i)%>" id="action_<%=PartList(0, i)%>">
                                <option value="">&nbsp;<option>
                                <option value="keep">Keep<option>
                                <option value="condense">Condense<option>
                                <option value="delete">Delete<option>
                            </select>
                        </td>
                    </tr>
                <%Next%>
            </table>
            </form>
        </div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
