<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lTeamID, lCellProvidersID
Dim sContactName, sAddress, sCity, sSt, sZip, sEmail, sCellPhone, sComments, sRole
Dim Contacts(), Teams(), CellProviders

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

lTeamID = Request.QueryString("team_id")
If CStr(lTeamID) = vbNullString Then lTeamID = 0

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

i = 0
ReDim Teams(1, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT TeamsID, TeamName, Gender FROM Teams WHERE CoachesID = " & Session("my_id")
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Teams(0, i) = rs(0).value 
	Teams(1, i) = rs(1).Value & " (" & rs(2).Value & ")"
	i = i + 1
	ReDim Preserve Teams(1, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If UBound(Teams, 2) = 1 Then lTeamID = Teams(0, 0)

If Request.Form.item("get_team") = "get_team" Then
	lTeamID = Request.Form.Item("teams")
    If CStr(lTeamID) = vbNullString Then lTeamID = 0
    'note that team 7 is a flag for displaying all teams by this coach because that number will not be used for a specific team
ElseIf Request.Form.item("submit_contact") = "submit_contact" Then
	sContactName = Replace(Request.Form.Item("contact_name"), "''", "'")
	sRole =  Request.Form.Item("role")
	sAddress =  Replace(Request.Form.Item("address"), "''", "'")
	sCity =  Replace(Request.Form.Item("city"), "''", "'")
	sSt = Request.Form.Item("state")
	sZip = Request.Form.Item("zip")
	sEmail = Request.Form.Item("email")
	sCellPhone = Request.Form.Item("cell_phone")
	sComments =  Replace(Request.Form.Item("comments"), "''", "'")
	lCellProvidersID = Request.Form.Item("cell_provider")

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "INSERT INTO TeamContacts (TeamsID, ContactName, Role, Address, City, St, Zip, CellPhone, Email, Comments, CellProvidersID) VALUES (" & lTeamID 
    sql = sql & ", '" & sContactName & "', '" & sRole & "', '" & sAddress & "', '" & sCity & "', '" & sSt & "', '" & sZip & "', '"  & sCellPhone & "', '" 
    sql = sql & sEmail & "', '" & sComments & "', " & lCellProvidersID & ")"
	rs = conn.Execute(sql)
	Set rs = Nothing
End If


ReDim Contacts(6, 0)
If Not CLng(lTeamID) = 0 Then
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    If CLng(lTeamID) = 7 Then
        sql = "SELECT tc.TeamContactsID, tc.ContactName, tc.Role, tc.CellPhone, tc.Email, tc.Comments, tc.CellProvidersID FROM TeamContacts tc "
        sql = sql & "INNER JOIN Teams t ON tc.TeamsID = t.TeamsID WHERE t.CoachesID = " & Session("my_id")
        sql = sql & "ORDER BY tc.ContactName"
    Else
        sql = "SELECT TeamContactsID, ContactName, Role, CellPhone, Email, Comments, CellProvidersID FROM TeamContacts WHERE TeamsID = " & lTeamID
        sql = sql & "ORDER BY ContactName"
    End If
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        For j = 0 To 6
            If Not rs(j).Value & "" = "" Then Contacts(j, i) = Replace(rs(j).Value, "''", "'")
        Next
        i = i + 1
        ReDim Preserve Contacts(6, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = NOthing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Team Contacts</title>

<script>
function chkFlds(){
 	if (document.add_contact.contact_name.value == '' ||
        document.add_contact.role.value == '' )
		{
  		alert('Contact Name and Role are required.');
  		return false
  		}
	else
   		return true
}
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
            <h4 class="h4">GSE Cross-Country/NordicTeam Contact Page</h4>

            <%If UBound(Teams, 2) > 1 Then%>
                <form role="form" class="form-inline" name="get_team" method="post" action="team_contacts.asp">
                <label for="teams">Select Team:</label>
                <select class="form-control" name="teams" id="teams" onchange="this.form.submit2.click();">
                    <option value="0">&nbsp;</option>
                    <%For i = 0 to UBound(Teams, 2) - 1%>
                        <%If CLng(Teams(0, i)) = CLng(lTeamID) Then%>
                            <option value="<%=Teams(0, i)%>" selected><%=Teams(1, i)%></option>
                        <%Else%>
                            <option value="<%=Teams(0, i)%>"><%=Teams(1, i)%></option>
                        <%End If%>
                    <%Next%>
                    <%If CLng(lTeamID) = 7 Then%>
                        <option value="7" selected>All My Contacts</option>
                    <%Else%>
                        <option value="7">All My Teams</option>
                    <%End If%>
                </select>
                <input type="hidden" name="get_team" id="get_team" value="get_team">
                <input type="submit" class="form-control" name="submit2" id="submit2" value="Get This Team">
                </form>
            <%End If%>		
                        
            <%If Not CLng(lTeamID) = 0 Then%>
                <%If CLng(lTeamID) = 7 Then%>
                    <p>You must select a specific team in order to add contacts.</p>
                <%Else%>
                    <hr>
                    <h4 class="h4">Add Contact (shaded fields are required):</h4>
                    <form role="form" class="form-horizontal" name="add_contact" method="post" action="team_contacts.asp?team_id=<%=lTeamID%>" onsubmit="return chkFlds()">
                    <div class="form-group row">
                        <label for="contact_name" class="control-label col-sm-2">Name:</label>
                        <div class="col-sm-4">
                            <input type="text" class="form-control" name="contact_name" id="contact_name">
                        </div>
                        <label for="role" class="control-label col-sm-2">Role:</label>
                        <div class="col-sm-4">
                            <select class="form-control" name="role" id="role">
                                <option value="">&nbsp;</option>
                                <option value="Parent">Parent</option>
                                <option value="Medical">Medical</option>
                                <option value="Other">Other</option>
                            </select>
                        </div>
                    </div>
                    <div class="form-group row">
                       <label for="email" class="control-label col-sm-2">Email:</label>
                        <div class="col-sm-4">
                            <input type="text" class="form-control" name="email" id="email">
                        </div>
                        <label for="cell_phone" class="control-label col-sm-2">Mobile:</label>
                        <div class="col-sm-4">
                            <input type="text" class="form-control" name="cell_phone" id="cell_phone">
                        </div>
                    </div>
                    <div class="form-group row">
                        <label for="cell_provider" class="control-label col-sm-2">Provider:</label>
                        <div class="col-sm-4">
                            <select class="form-control" name="cell_provider" id="cell_provider">
                                <option value="0">None</option>
                                <%For i = 0 To UBound(CellProviders, 2) - 1%>
                                    <option value="<%=CellProviders(0, i)%>"><%=CellProviders(1, i)%></option>
                                <%Next%>
                            </select>
                        </div>
                        <label for="address" class="control-label col-sm-2">Address:</label>
                        <div class="col-sm-4">
                            <input type="text" class="form-control" name="address" id="address">
                        </div>
                    </div>
                    <div class="form-group row">
                        <label for="city" class="control-label col-sm-2">City:</label>
                        <div class="col-sm-4">
                            <input type="text" class="form-control" name="city" id="city">
                        </div>
                        <label for="state" class="control-label col-sm-2">St:</label>
                        <div class="col-sm-4">
                            <input type="text" class="form-control" name="state" id="state">
                        </div>
                    </div>
                    <div class="form-group row">
                       <label for="zip" class="control-label col-sm-2">Zip:</label>
                        <div class="col-sm-4">
                            <input type="text" class="form-control" name="zip" id="zip">
                        </div>
                        <label for="comments" class="control-label col-sm-2">Comments:</label>
                        <div class="col-sm-4">
                            <textarea class="form-control" name="comments" id="comments" rows="3"></textarea>
                        </div>
                    </div>
                    <div class="form-group">
                        <input type="hidden" name="submit_contact" id="submit_contact" value="submit_contact">
                        <input type="submit" class="form-control" name="submit1" id="submit1" value="Submit Contact">
                    </div>
                    </form>
                <%End If%>

                <h4 class="h4">Team Contacts</h4>
                <table class="table table-striped">
                    <tr>
                        <th>No.</th>
                        <th>Name (click to edit)</th>
                        <th>Role</th>
                        <th>Phone</th>
                        <th>Email</th>
                        <th>Comments</th>
                    </tr>
                    <%For i = 0 To UBound(Contacts, 2) - 1%>
                        <tr>
                            <td><%=i + 1%>)</td>
                            <td><a href="javascript:pop('edit_contact.asp?which_contact=<%=Contacts(0, i)%>',1000,400)"><%=Contacts(1, i)%></a></td>
                            <td><%=Contacts(2, i)%></td>
                            <td><%=Contacts(3, i)%></td>
                            <td><a href="mailto:<%=Contacts(4, i)%>"><%=Contacts(4, i)%></a></td>
                            <td><%=Contacts(5, i)%></td>
                        </tr>
                    <%Next%>
                </table>
            <%End If%>
        </div>
    </div>
</div>
<!--#include file = "../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
