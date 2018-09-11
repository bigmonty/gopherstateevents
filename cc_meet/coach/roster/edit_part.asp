<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql, rs2, sql2
Dim lRosterID, lCellProvidersID
Dim CellProviders
Dim i, j
Dim sFirstName, sLastName, sGender, iGrade, sEmail, sCellPhone, sArchiveThis
Dim sErrMsg
Dim iGradeYear
Dim bInsertThis, bNullGrade

If Not (Session("role") = "coach" Or Session("role") = "team_staff") Then Response.Redirect "/default.asp?sign_out=y"

lRosterID = Request.QueryString("roster_id")

'get year for roster grades
If Month(Date) <=7 Then
	iGradeYear = CInt(Right(CStr(Year(Date) - 1), 2))
Else
	iGradeYear = Right(CStr(Year(Date)), 2)	
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
												
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get cell providers
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT CellProvidersID, Provider FROM CellProviders ORDER BY Provider"
rs.Open sql, conn, 1, 2
CellProviders = rs.GetRows()
rs.Close
Set rs = Nothing

If Request.Form.Item("edit_roster") = "edit_roster" Then
    If Request.Form.Item("delete") = "y" Then
        sArchiveThis = "n"
		'check ind results to make sure this person has not been entered in results
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "SELECT RosterID FROM IndRslts WHERE RosterID = " & lRosterID
		rs.Open sql, conn, 1, 2
		If rs.RecordCount > 0 Then sArchiveThis = "y"
		rs.Close
		Set rs = Nothing

        If sArchiveThis = "y" Then
			'if they exist in ind results, just archive them
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT Archive FROM Roster WHERE RosterID = " & lRosterID
			rs.Open sql, conn, 1, 2
			rs(0).Value = "y"
			rs.Update
			rs.Close
			Set rs = Nothing
        Else
			'if they don't exist in ind results then delete them
			sql = "DELETE FROM Roster Where RosterID = " & lRosterID
			Set rs = conn.Execute(sql)
			Set rs = Nothing
        End If

        Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
		Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
    Else
	    sFirstName = Replace(Request.Form.Item("first_name"), "'", "''")
	    sLastName = Replace(Request.Form.Item("last_name"), "'", "''")
	    sGender = Request.Form.Item("gender")
        sEmail = Request.Form.Item("email")
        sCellPhone = Request.Form.Item("cell_phone")
        lCellProvidersID = Request.Form.Item("cell_provider")
        iGrade = Request.Form.Item("grade")

	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT FirstName, LastName, Gender, Email, CellPhone, CellProvidersID FROM Roster WHERE RosterID = " & lRosterID
	    rs.Open sql, conn, 1, 2
		If sFirstName = vbNullString Then
			rs(0).Value = rs(0).OriginalValue
		Else
			rs(0).Value = sFirstName
		End If
	
		If sLastName = vbNullString Then
			rs(1).Value = rs(1).OriginalValue
		Else
			rs(1).Value = sLastName
		End If
	
		rs(2).Value = sGender
		rs(3).Value = sEmail
        rs(4).Value = sCellPhone
		rs(5).Value = lCellProvidersID				
		rs.Update
	    rs.Close
	    Set rs = Nothing
				
		Call UpdateGrade(iGrade)

        Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
    End If		
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Gender, Email, CellPhone, CellProvidersID FROM Roster WHERE RosterID = " & lRosterID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    sFirstName = Replace(rs(0).Value, "''", "'")
    sLastName = Replace(rs(1).Value, "''", "'")
    sGender = rs(2).Value
    sEmail = rs(3).Value
    sCellPhone = rs(4).Value
    lCellProvidersID = rs(5).Value
    iGrade = GetGrade()
Else
	Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
End If
rs.Close
Set rs = Nothing
	
Private Function UpdateGrade(iCurrGrade)
    bInsertThis = False

	Set rs2 = Server.CreateObject("ADODB.Recordset")
	sql2 = "SELECT Grade" & iGradeYear & " FROM Grades WHERE RosterID = " & lRosterID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then 
        rs2(0).Value = iCurrGrade
	    rs2.Update
    Else
        bInsertThis = True
    End If
	rs2.Close
	Set rs2 = Nothing

    If bInsertThis = True Then
        sql2 = "INSERT INTO Grades (RosterID,  Grade" & iGradeYear & ") Values (" & lRosterID & ", " & iCurrGrade & ")"
        Set rs2 = conn.Execute(sql2)
        Set rs2 = Nothing
    End If
End Function
	
Private Function GetGrade()
    GetGrade = "0"

    bNullGrade = False
	Set rs2 = Server.CreateObject("ADODB.Recordset")
    sql2 = "SELECT Grade" & iGradeYear & " FROM Grades WHERE RosterID = " & lRosterID
	rs2.Open sql2, conn, 1, 2
	If rs2.RecordCount > 0 Then 
        GetGrade = rs2(0).Value
        bNullGrade = True
    End If
	rs2.Close
	Set rs2 = Nothing

    If bNullGrade = False Then
        sql2 = "INSERT INTO Grades (RosterID,  Grade" & iGradeYear & ") Values (" & lMyID & ", 0)"
        Set rs2 = conn.Execute(sql2)
        Set rs2 = Nothing
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE View Roster</title>
</head>
<body>
<div class="container">
	<h4 class="h4">GSE Cross-Country/Nordic Ski Edit Roster Page</h4>

    <p>
        Note:  Deleting a participant that already has results in the system from this year or a previous year will
        actually result in their account being archived to preserve past results.  They will only be deleted from the system if they have not finished
        a race timed by Gopher State Events.
    </p>

	<form role="form" class="form-horizontal" name="edit_roster" method="post" action="edit_part.asp?roster_id=<%=lRosterID%>">
 	<div class="form-group row">
		<label for="first_name" class="control-label col-sm-1">First:</label>
		<div class="col-sm-3">
            <input type="text" class="form-control" name="first_name" id="first_name" value="<%=sFirstName%>">
        </div>
		<label for="last_name" class="control-label col-sm-1">Last:</label>
		<div class="col-sm-3">
            <input type="text" class="form-control" name="last_name" id="last_name" value="<%=sLastName%>">
        </div>
		<label for="grade" class="control-label col-sm-1">Grade:</label>
		<div class="col-sm-3">
			<select class="form-control" name="grade" id="grade"> 
				<%For j = 0 to 16%>
					<%If CInt(iGrade) = CInt(j) Then%>
						<option value="<%=j%>" selected><%=j%></option>
					<%Else%>
						<option value="<%=j%>"><%=j%></option>
					<%End If%>
				<%Next%>
			</select>
        </div>
	</div>
 	<div class="form-group row">
		<label for="gender" class="control-label col-sm-1">M/F:</label>
		<div class="col-sm-3">
            <input type="text" class="form-control" name="gender" id="gender" value="<%=sGender%>">
        </div>
		<label for="email" class="control-label col-sm-1">Email:</label>
		<div class="col-sm-3">
            <input type="text" class="form-control" name="email" id="email" value="<%=sEmail%>">
        </div>
		<label for="cell_phone" class="control-label col-sm-1">Cell:</label>
		<div class="col-sm-3">
			<input type="text" class="form-control" name="cell_phone" id="cell_phone" value="<%=sCellPhone%>">
        </div>
	</div>
    <div class="form-group row">
		<label for="cell_provider" class="control-label col-sm-2">Provider:</label>
		<div class="col-sm-4">
			<select class="form-control" name="cell_provider" id="cell_provider"> 
                <option value="0">None</option>
				<%For j = 0 To UBound(CellProviders, 2)%>
                    <%If CLng(lRosterID) = CLng(CellProviders(0, j)) Then%>
						<option value="<%=CellProviders(0, j)%>" selected><%=CellProviders(1, j)%></option>
					<%Else%>
						<option value="<%=CellProviders(0, j)%>"><%=CellProviders(1, j)%></option>
					<%End If%>
                <%Next%>
			</select>
        </div>
		<label for="email" class="control-label col-sm-2">Delete?</label>
		<div class="col-sm-4">
 			<select class="form-control" name="delete" id="delete"> 
                <option value="n">No</option>
				<option value="y">Yes</option>
			</select>
        </div>
	</div>
    <div class="form-group row">
		<input type="hidden" name="edit_roster" id="edit_roster" value="edit_roster">
		<input class="form-control" type="submit" name="submit1" id="submit1" value="Save Changes">
	</div>
	</form>
</div>
<!--#include file = "../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
