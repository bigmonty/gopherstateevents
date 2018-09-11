<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lWhichContact, lTeamID, lCellProvidersID
Dim sContactName, sAddress, sCity, sSt, sZip, sEmail, sCellPhone, sComments, sRole
Dim Teams(), Contact(10), Roles(2), CellProviders

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

lWhichContact = Request.QueryString("which_contact")

Roles(0) = "Parent"
Roles(1) = "Medical"
Roles(2) = "Other"

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

If Request.Form.item("submit_changes") = "submit_changes" Then
    If Request.Form.Item("delete_contact") = "on" Then
        sql = "DELETE FROM TeamContacts WHERE TeamContactsID = " & lWhichContact
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
        Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
    Else
	    sContactName = Replace(Request.Form.Item("contact_name"), "''", "'")
	    sRole =  Request.Form.Item("role")
	    sAddress =  Replace(Request.Form.Item("address"), "''", "'")
	    sCity =  Replace(Request.Form.Item("city"), "''", "'")
	    sSt = Request.Form.Item("state")
	    sZip = Request.Form.Item("zip")
	    sEmail = Request.Form.Item("email")
	    sCellPhone = Request.Form.Item("cell_phone")
	    sComments =  Replace(Request.Form.Item("comments"), "''", "'")
        lTeamID = Request.Form.Item("team_id")
        lCellProvidersID = Request.Form.Item("cell_provider")

        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ContactName, Role, Address, City, St, Zip, CellPhone, Email, Comments, TeamsID, CellProvidersID FROM TeamContacts "
        sql = sql & "WHERE TeamContactsID = " & lWhichContact
        rs.Open sql, conn, 1, 2
        If sContactName & "" = "" Then
            rs(0).Value = rs(0).OriginalValue
        Else
            rs(0).Value = sContactName
        End If
        rs(1).Value = sRole
        rs(2).Value = sAddress
        rs(3).Value = sCity
        rs(4).Value = sSt
        rs(5).Value = sZip
        rs(6).Value = sCellPhone
        rs(7).Value = sEmail
        rs(8).Value = sComments
        rs(9).Value = lTeamID
        rs(10).Value = lCellProvidersID
        rs.Update
        rs.Close
        Set rs = NOthing
    End If
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ContactName, Role, Address, City, St, Zip, CellPhone, Email, Comments, TeamsID, CellProvidersID FROM TeamContacts "
sql = sql & "WHERE TeamContactsID = " & lWhichContact
rs.Open sql, conn, 1, 2
For i = 0 To 10
    If Not rs(i).Value & "" = "" Then Contact(i) = Replace(rs(i).Value, "''", "'")
Next
rs.Close
Set rs = NOthing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Cross-Country/Nordic Ski Team Contacts</title>
<!--#include file = "../../../includes/js.asp" --> 
</head>
<body>
<div class="container">
    <h4 class="h4">Edit Contact:&nbsp;<%=Contact(0)%>:</h4>
	<form role="form" class="form-horizontal" name="add_contact" method="post" action="edit_contact.asp?which_contact=<%=lWhichContact%>">
	<div class="form-group">
		<label for="contact_name" class="control-label col-xs-1">Name:</label>
		<div class="col-xs-3">
            <input type="text" class="form-control" name="contact_name" id="contact_name" value="<%=Contact(0)%>">
        </div>
		<label for="role" class="control-label col-xs-1">Role:</label>
		<div class="col-xs-3">
            <select class="form-control" name="role" id="role">
                <%Select Case Contact(1)%>
                    <%Case "Parent"%>
                        <option value="Parent" selected>Parent</option>
                        <option value="Medical">Medical</option>
                        <option value="Other">Other</option>
                    <%Case "Medical"%>
                        <option value="Parent">Parent</option>
                        <option value="Medical" selected>Medical</option>
                        <option value="Other">Other</option>
                    <%Case "Other"%>
                        <option value="Parent">Parent</option>
                        <option value="Medical">Medical</option>
                        <option value="Other" selected>Other</option>
                <%End Select%>
            </select>
        </div>
		<label for="email" class="control-label col-xs-1">Email:</label>
		<div class="col-xs-3">
            <input type="text" class="form-control" name="email" id="email" value="<%=Contact(7)%>">
        </div>
	</div>
	<div class="form-group">
		<label for="cell_phone" class="control-label col-xs-1">Mobile:</label>
		<div class="col-xs-3">
            <input type="text" class="form-control" name="cell_phone" id="cell_phone" value="<%=Contact(6)%>">
        </div>
		<label for="cell_provider" class="control-label col-xs-1">Provider:</label>
		<div class="col-xs-3">
            <select class="form-control" name="cell_provider" id="cell_provider">
                <option value="0">None</option>
                <%For i = 0 To UBound(CellProviders, 2) - 1%>
                    <%If CLng(Contact(10)) = CLng(CellProviders(0, i)) Then%>
                        <option value="<%=CellProviders(0, i)%>" selected><%=CellProviders(1, i)%></option>
                    <%Else%>
                        <option value="<%=CellProviders(0, i)%>"><%=CellProviders(1, i)%></option>
                    <%End If%>
                <%Next%>
            </select>
        </div>
		<label for="address" class="control-label col-xs-1">Address:</label>
		<div class="col-xs-3">
            <input type="text" class="form-control" name="address" id="address" value="<%=Contact(2)%>">
        </div>
	</div>
	<div class="form-group">
		<label for="city" class="control-label col-xs-1">City:</label>
		<div class="col-xs-3">
            <input type="text" class="form-control" name="city" id="city" value="<%=Contact(3)%>">
        </div>
		<label for="state" class="control-label col-xs-1">St:</label>
		<div class="col-xs-3">
            <input type="text" class="form-control" name="state" id="state" value="<%=Contact(4)%>">
        </div>
		<label for="zip" class="control-label col-xs-1">Zip:</label>
		<div class="col-xs-3">
            <input type="text" class="form-control" name="zip" id="zip" value="<%=Contact(5)%>">
        </div>
	</div>
	<div class="form-group">
		<label for="team_id" class="control-label col-xs-1">Team:</label>
		<div class="col-xs-3">
            <select class="form-control" name="team_id" id="team_id">
                <%For i = 0 To UBound(Teams, 2) - 1%>
                    <%If CLng(Contact(9)) = CLng(Teams(0, i)) Then%>
                        <option value="<%=Teams(0, i)%>" selected><%=Teams(1, i)%></option>
                    <%Else%>
                        <option value="<%=Teams(0, i)%>"><%=Teams(1, i)%></option>
                    <%End If%>
                <%Next%>
            </select>
        </div>
		<label for="comments" class="control-label col-xs-1">Comments:</label>
		<div class="col-xs-7">
            <textarea class="form-control" name="comments" id="comments" rows="3"><%=Contact(8)%></textarea>
        </div>
	</div>
	<div class="form-group">
		<input type="checkbox" name="delete_contact" id="delete_contact">&nbsp;Delete Contact (No Undo!)
	</div>
	<div class="form-group">
		<input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
		<input type="submit" class="form-control" name="submit1" id="submit1" value="Save Changes">
	</div>
	</form>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
