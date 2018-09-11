<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim lRaceID, lEventID, lCustomFieldsID
Dim sEventName, sRaceName, sFieldName
Dim CustomFields(), Races(), RaceParts, SortArr(6)
Dim dEventDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
lRaceID = Request.QueryString("race_id")
lCustomFieldsID = Request.QueryString("custom_fields_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

i = 0
ReDim Races(1, 0)
sql = "SELECT RaceID, RaceName FROM RaceData WHERE EventID = " & lEventID
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Races(0, i) = rs(0).Value
	Races(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve Races(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("submit_delete") = "submit_delete" Then
    sql = "DELETE FROM CustomFields WHERE CustomFieldsID = " & lCustomFieldsID
    Set rs = conn.Execute(sql)
    Set rs = Nothing

    lCustomFieldsID = 0
ElseIf Request.Form.Item("submit_status") = "submit_status" Then
	Call GetRaceParts()
    Call GetCustomFields()

    For i = 0 To UBound(RaceParts, 2) - 1
        If Request.Form.Item("status_" & RaceParts(5, i)) = "n" Then
            sql = "DELETE FROM CustomFieldsParts WHERE ParticipantID = " & RaceParts(5, i) & " AND CustomFieldsID = " & lCustomFieldsID
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        Else
            If RaceParts(6, i) = "n" Then 
                sql = "INSERT INTO CustomFieldsParts (CustomFieldsID, ParticipantID) VALUES (" & lCustomFieldsID & ", " & RaceParts(5, i) & ")"
                Set rs = conn.Execute(sql)
                Set rs = Nothing
            End If
        End If
    Next
ElseIf Request.Form.Item("submit_field") = "submit_field" Then
	lCustomFieldsID = Request.Form.Item("fields")
ElseIf Request.Form.Item("submit_race") = "submit_race" Then
	lRaceID = Request.Form.Item("races")
ElseIf Request.Form.Item("submit_new_field") = "submit_new_field" Then
    sFieldName = Request.Form.Item("field_name")

    sql = "INSERT INTO CustomFields (RaceID, FieldName) VALUES (" & lRaceID & ", '" & sFieldName & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
End If

If UBound(Races, 2) = 1 Then lRaceID = Races(0, 0)
If CStr(lRaceID) = vbNullString Then lRaceID = Races(0, 0)

sql = "SELECT RaceName FROM RaceData WHERE RaceID = " & lRaceID
Set rs = conn.Execute(sql)
sRaceName = Replace(rs(0).Value, "''", "'")
Set rs = Nothing

Call GetCustomFields()

Private Sub GetCustomFields()
    i = 0
    ReDim CustomFields(1, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT CustomFieldsID, FieldName FROM CustomFields WHERE RaceID = " & lRaceID
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        CustomFields(0, i) = rs(0).Value
        CustomFields(1, i) = rs(1).Value
        i = i + 1
        ReDim Preserve CustomFields(1, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    If CStr(lCustomFieldsID) = vbNullString Then 
        If UBound(CustomFields, 2) > 0 Then
            lCustomFieldsID = CustomFields(0, 0)
        Else
            lCustomFieldsID = 0
        End If
    End If
End Sub

Call GetRaceParts()

Private Sub GetRaceParts()
    i = 0
    ReDim RaceParts(6, 0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT rc.Bib, p.FirstName, p.LastName, p.Gender, rc.Age, p.ParticipantID, p.DOB FROM Participant p INNER JOIN PartRace rc "
    sql = sql & "ON rc.ParticipantID = p.ParticipantID WHERE rc.RaceID = " & lRaceID & " ORDER BY p.LastName, p.FirstName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        RaceParts(0, i) = rs(0).Value
        RaceParts(1, i) = Replace(rs(1).Value, "''", "'")
        RaceParts(2, i) = Replace(rs(2).Value, "''", "'")
        RaceParts(3, i) = rs(3).Value
        RaceParts(4, i) = rs(4).Value
        RaceParts(5, i) = rs(5).Value
        RaceParts(6, i) = rs(6).Value
        i = i + 1
        ReDim Preserve RaceParts(6, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs=Nothing

    'replace field 6 with status
    For i = 0 To UBound(RaceParts, 2) - 1
        RaceParts(6, i) = MyStatus(RaceParts(5, i))
    Next

    'sort by status
    For i = 0 To UBound(RaceParts, 2) - 2
        For j = i + 1 To UBound(RaceParts, 2) - 1
            If RaceParts(6, i) < RaceParts(6, j) Then
                For k = 0 to 6
                    SortArr(k) = RaceParts(k, i)
                    RaceParts(k, i) = RaceParts(k, j)
                    RaceParts(k, j) = SortArr(k)
                Next
            End If
        Next
    Next
End Sub

Private Function MyStatus(lMyID)
    MyStatus = "n"

    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT ParticipantID FROM CustomFieldsParts WHERE CustomFieldsID = " & lCustomFieldsID & " AND ParticipantID = " & lMyID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then MyStatus = "y"
    rs.Close
    Set rs=Nothing
End Function
%>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title><%=sEventName%> Custom Fields Manager</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<img src="/graphics/html_header.png" alt="Gopher State Events" class="img-responsive">
<div class="container">
	<h3 class="h3">Custom Fields For <%=sRaceName%> in <%=sEventName%> (<%=dEventDate%>)</h3>
				
	<%If UBound(Races, 2) > 1 Then%>
		<form class="form-inline" name="get_races" method="post" action="custom_fields.asp?event_id=<%=lEventID%>">
		<label for="races">Select Race:</label>
		<select class="form-control" name="races" id="races" onchange="this.form.get_race.click()">
			<%For i = 0 to UBound(Races, 2) - 1%>
				<%If CLng(lRaceID) = CLng(Races(0, i)) Then%>
					<option value="<%=Races(0, i)%>" selected><%=Races(1, i)%></option>
				<%Else%>
					<option value="<%=Races(0, i)%>"><%=Races(1, i)%></option>
				<%End If%>
			<%Next%>
		</select>
		<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
		<input type="submit" class="form-control" name="get_race" id="get_race" value="Get Race Info">
		</form>
	<%End If%>
	
    <div class="bg-success">
        <h4 class="h4">Add Custom Field</h4>
        <form role="form" class="form-inline" name="create_field" method="post" action="custom_fields.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
        <label for="field_name">Field Name:</label>
        <input type="text" class="form-control" name="field_name" id="field_name">
        <input type="hidden" name="submit_new_field" id="submit_new_field" value="submit_new_field">
        <input type="submit" class="form-control" name="submit1" id="submit1" value="Create Field">
        </form>
    </div>
	
    <div>
        <h4 class="h4">Manage Custom Fields</h4>

		<form class="form-inline" name="get_cust_fields" method="post" action="custom_fields.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>">
		<label for="fields">Select Field:</label>
		<select class="form-control" name="fields" id="fields" onchange="this.form.get_field.click()">
			<%For i = 0 to UBound(CustomFields, 2) - 1%>
				<%If CLng(lCustomFieldsID) = CLng(CustomFields(0, i)) Then%>
					<option value="<%=CustomFields(0, i)%>" selected><%=CustomFields(1, i)%></option>
				<%Else%>
					<option value="<%=CustomFields(0, i)%>"><%=CustomFields(1, i)%></option>
				<%End If%>
			<%Next%>
		</select>
		<input type="hidden" name="submit_field" id="submit_field" value="submit_field">
		<input type="submit" class="form-control" name="get_field" id="get_field" value="Get Custom Field">
		</form>

        <%If CLng(lCustomFieldsID) > 0 Then%>
            <form class="form" name="delete_field" method="post" action="custom_fields.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;custom_fields_id=<%=lCustomFieldsID%>">
            <input type="hidden" name="submit_delete" id="submit_delete" value="submit_delete">
            <input type="submit" class="form-control" name="delete_field" id="delete_field" 
                value="Delete This Field (There is no undo for this action!">
            </form>

            <hr>
        <%End If%>

        <form class="form" name="edit_status" method="post" action="custom_fields.asp?event_id=<%=lEventID%>&amp;race_id=<%=lRaceID%>&amp;custom_fields_id=<%=lCustomFieldsID%>">
        <table class="table table-striped">
            <tr>
                <th>No.</th>
                <th>Bib</th>
                <th>First</th>
                <th>Last</th>
                <th>Gender</th>
                <th>Age</th>
                <th>Status</th>
            </tr>
            <%For i = 0 To UBound(RaceParts, 2) - 1%>
                <tr>
                    <td><%=i + 1%></td>
                    <td><%=RaceParts(0, i)%></td>
                    <td><%=RaceParts(1, i)%></td>
                    <td><%=RaceParts(2, i)%></td>
                    <td><%=RaceParts(3, i)%></td>
                    <td><%=RaceParts(4, i)%></td>
                    <td>
                        <select class="form-control" name="status_<%=RaceParts(5, i)%>" id="status_<%=RaceParts(5, i)%>">
                            <%If RaceParts(6, i) = "y" Then%>
                                <option value="n">No</option>
                                <option value="y" selected>Yes</option>
                            <%Else%>
                                <option value="n">No</option>
                                <option value="y">Yes</option>
                            <%End If%>
                        </select>
                    </td>
                </tr>
            <%Next%>
            <tr>
                <td colspan="6">
		            <input type="hidden" name="submit_status" id="submit_status" value="submit_status">
		            <input type="submit" class="form-control" name="get_status" id="get_status" value="Save Changes">
                </td>
            </tr>
        </table>
        </form>
    </div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>