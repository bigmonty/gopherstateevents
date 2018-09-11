<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim lProvider, lTeamStaffID
Dim i, j
Dim StaffRoles(2), CellProviders, StaffData(7)

If Not Session("role") = "coach" Then Response.Redirect "/default.asp?sign_out=y"
lTeamStaffID = Request.QueryString("staff_id")

StaffRoles(0) = "Asst Coach"
StaffRoles(1) = "Manager"
StaffRoles(2) = "Other"

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

If Request.Form.Item("submit_this") = "submit_this" Then
    If Request.Form.Item("delete") = "y" Then
        sql = "DELETE FROM TeamStaff WHERE TeamStaffID = " & lTeamStaffID
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
		Response.Write("<script type='text/javascript'>{window.close() ;}</script>")
    Else
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT FirstName, LastName, Email, MobilePhone, Provider, SendTo, AllowAccess, Role FROM TeamStaff WHERE TeamStaffID = "& lTeamStaffID
        rs.Open sql, conn, 1, 2
	    rs(0).Value = Replace(Request.Form.Item("first_name"), "''", "'")
	    rs(1).Value = Replace(Request.Form.Item("last_name"), "''", "'")
        rs(2).Value =  Request.Form.Item("email")
	    rs(3).Value =  Request.Form.Item("mobile_phone")
        rs(4).Value =  Request.Form.Item("provider")
	    rs(5).Value =  Request.Form.Item("send_to")
	    rs(6).Value =  Request.Form.Item("allow_access")
	    rs(7).Value = Request.Form.Item("role")
        rs.Update
        rs.Close
        Set rs = Nothing

        Response.Write("<script type='text/javascript'>{window.opener.location.reload();}</script>")
    End If
End If

Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT FirstName, LastName, Email, MobilePhone, Provider, SendTo, AllowAccess, Role FROM TeamStaff WHERE TeamStaffID = " & lTeamStaffID
rs.Open sql, conn, 1, 2
For i = 0 To 7
    If Not rs(i).Value & "" = "" Then StaffData(i) = Replace(rs(i).Value, "''", "'")
Next
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE View Staff</title>
<!--#include file = "../../includes/js.asp" --> 

</head>

<body>
<div class="container">
	<h4 class="h4">Edit Team Staff</h4>

    <form role="form" class="form-horizontal" name="edit_staff" method="post" action="edit_staff.asp?staff_id=<%=lTeamStaffID%>">
	<div class="form-group">
		<label for="first_name" class="control-label col-xs-2">First Name:</label>
		<div class="col-xs-4">
            <input type="text" class="form-control" name="first_name" id="first_name" value="<%=StaffData(0)%>">
        </div>
		<label for="last_name" class="control-label col-xs-2">Last Name:</label>
		<div class="col-xs-4">
            <input type="text" class="form-control" name="last_name" id="last_name" value="<%=StaffData(1)%>">
        </div>
	</div>
	<div class="form-group">
		<label for="email" class="control-label col-xs-2">Email:</label>
		<div class="col-xs-4">
            <input type="text" class="form-control" name="email" id="email" value="<%=StaffData(2)%>">
        </div>
		<label for="send_to" class="control-label col-xs-2">Send To:</label>
		<div class="col-xs-4">
            <select class="form-control" name="send_to" id="send_to">
                <%If StaffData(5) = "n" Then%>
                    <option value="n" selected>No</option>
                    <option value="y">Yes</option>
                <%Else%>
                    <option value="n">No</option>
                    <option value="y"selected>Yes</option>
                <%End If%>
            </select>
        </div>
	</div>
	<div class="form-group">
		<label for="mobile_phone" class="control-label col-xs-2">Mobile #:</label>
		<div class="col-xs-4">
            <input type="text" class="form-control" name="mobile_phone" id="mobile_phone" value="<%=StaffData(3)%>">
        </div>
		<label for="provider" class="control-label col-xs-2">Provider:</label>
		<div class="col-xs-4">
			<select class="form-control" name="provider" id="provider"> 
                <option value="0">None</option>
				<%For j = 0 To UBound(CellProviders, 2)%>
                    <%If CLng(StaffData(4)) = CLng(CellProviders(0, j)) Then%>
						<option value="<%=CellProviders(0, j)%>" selected><%=CellProviders(1, j)%></option>
					<%Else%>
						<option value="<%=CellProviders(0, j)%>"><%=CellProviders(1, j)%></option>
					<%End If%>
                <%Next%>
			</select>
        </div>
	</div>
	<div class="form-group">
		<label for="allow_access" class="control-label col-xs-2">Allow Access:</label>
		<div class="col-xs-4">
            <select class="form-control" name="allow_access" id="allow_access">
                <%If StaffData(6) = "n" Then%>
                    <option value="n" selected>No</option>
                    <option value="y">Yes</option>
                <%Else%>
                    <option value="n">No</option>
                    <option value="y"selected>Yes</option>
                <%End If%>
            </select>
        </div>
		<label for="role" class="control-label col-xs-2">Role:</label>
		<div class="col-xs-4">
            <select class="form-control" name="role" id="role">
                <%For j = 0 To UBound(StaffRoles)%>
                    <%IF StaffData(7) = StaffRoles(j) Then%>
                        <option value="<%=StaffRoles(j)%>" selected><%=StaffRoles(j)%></option>
                    <%Else%>
                        <option value="<%=StaffRoles(j)%>"><%=StaffRoles(j)%></option>
                    <%End If%>
                <%Next%>
            </select>
        </div>
	</div>
	<div class="form-group">
		<label for="delete" class="control-label col-xs-2">Delete:</label>
		<div class="col-xs-10">
            <select class="form-control" name="delete" id="delete">
                <option value="n" selected>No</option>
                <option value="y">Yes</option>
            </select>
        </div>
	</div>
	<div class="form-group">
		<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
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
