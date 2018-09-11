<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lEventDir
Dim EventDirArr(12), EventDir(), Events()
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sPhone, sEmail, sUserID, sPassword, sComments, sActive, sMobile

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventDir = Request.QueryString("event_dir")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_event_dir") = "submit_event_dir" Then
    If Request.Form.Item("delete") = "on" Then
        sql = "DELETE FROM EventDir WHERE EventDirID = " & lEventDir
        Set rs = conn.Execute(sql)
        Set rs = Nothing

        lEventDir = 0
    Else
	    If Not Request.Form.Item("first_name") & "" = "" Then sFirstName = Replace(Request.Form.Item("first_name"), "''", "'")
	    If Not Request.Form.Item("last_name") & "" = "" Then sLastName = Replace(Request.Form.Item("last_name"), "''", "'")
	    If Not Request.Form.Item("address") & "" = "" Then sAddress =  Replace(Request.Form.Item("address"), "''", "'")
	    If Not Request.Form.Item("city") & "" = "" Then sCity =  Replace(Request.Form.Item("city"), "''", "'")
	    If Not Request.Form.Item("state") & "" = "" Then sState =  Replace(Request.Form.Item("state"), "''", "'")
	    If Not Request.Form.Item("zip") & "" = "" Then sZip =  Replace(Request.Form.Item("zip"), "''", "'")
	    If Not Request.Form.Item("phone") & "" = "" Then sPhone =  Replace(Request.Form.Item("phone"), "''", "'")
	    If Not Request.Form.Item("email") & "" = "" Then sEmail =  Replace(Request.Form.Item("email"), "''", "'")
	    If Not Request.Form.Item("user_id") & "" = "" Then sUserID =  Replace(Request.Form.Item("user_id"), "''", "'")
	    If Not Request.Form.Item("password") & "" = "" Then sPassword =  Replace(Request.Form.Item("password"), "''", "'")
	    If Not Request.Form.Item("comments") & "" = "" Then sComments =  Replace(Request.Form.Item("comments"), "''", "'")
	    sActive =  Request.Form.Item("active")
        sMobile =  Request.Form.Item("mobile")

	    Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, UserID, Password, Comments, Active, Mobile FROM EventDir "
        sql = sql & "WHERE EventDirID = " & lEventDir
	    rs.Open sql, conn, 1, 2
	
	    If sFirstName & "" = "" Then
		    rs(0).Value = rs(0).OriginalValue
	    Else
		    rs(0).Value = sFirstName
	    End if
	
	    If sLastName & "" = "" Then
		    rs(1).Value = rs(1).OriginalValue
	    Else
		    rs(1).Value = sLastName
	    End if

	    rs(2).Value = sAddress
	    rs(3).Value = sCity
	    rs(4).Value = sState
	    rs(5).Value = sZip
	
	    If sEmail & "" = "" Then
		    rs(6).Value = rs(6).OriginalValue
	    Else
		    rs(6).Value = sEmail
	    End if
	
	    rs(7).Value = sPhone
	
	    If sUserID & "" = "" Then
		    rs(8).Value = rs(8).OriginalValue
	    Else
		    rs(8).Value = sUserID
	    End if
	
	    If sPassword & "" = "" Then
		    rs(9).Value = rs(9).OriginalValue
	    Else
		    rs(9).Value = sPassword
	    End if

	    rs(10).Value = sComments
        rs(11).Value = sActive
        rs(12).Value = sMobile
	    rs.Update
	    rs.Close
	    Set rs = Nothing
    End If
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	lEventDir = Request.Form.Item("event_dir")
End If

i = 0
ReDim EventDir(1, 0)
sql = "SELECT EventDirID, FirstName, LastName FROM EventDir ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventDir(0, i) = rs(0).Value
	EventDir(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve EventDir(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If CStr(lEventDir) = vbNullString Then lEventDir = 0

If Not CLng(lEventDir) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Phone, Address, City, State, Zip, Email, UserID, Password, Comments, Active, Mobile FROM EventDir WHERE EventDirID = " 
    sql = sql & lEventDir
	rs.Open sql, conn, 1, 2
	For i = 0 to 12
		If not rs(i).Value & "" = "" Then EventDirArr(i) =  Replace(rs(i).Value, "''", "'")
	Next
	rs.Close
	Set rs = Nothing
End If

ReDim Events(4, 0)
Private Sub MyEvents(lEventDirID)
    Dim x

    x = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EventID, EventName, EventDate, EventGrp, Edition FROM Events WHERE EventDirID = " & lEventDirID & " ORDER BY EventDate DESC, Edition"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Events(0, x) = rs(0).Value
        Events(1, x) = Replace(rs(1).Value, "''", "'")
        Events(2, x) = rs(2).Value
        Events(3, x) = rs(3).Value
        Events(4, x) = rs(4).Value

        x = x + 1
        ReDim Preserve Events(4, x)

        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Edit Event Director</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <!--#include file = "../../includes/event_dir_nav.asp" -->

			<h4 class="h4">Edit Event Director</h4>
		
			<form role="form"class="form-inline" name="edit_event_dir" method="Post" action="edit_event_dir.asp">
			<select class="form-control" name="event_dir" id="event_dir" onchange="this.form.submit2.click();">
				<option value="">&nbsp;</option>
				<%For i = 0 To UBound(EventDir, 2) - 1%>
					<%If CLng(EventDir(0, i)) = CLng(lEventDir) Then%>
						<option value="<%=EventDir(0, i)%>" selected><%=EventDir(1, i)%></option>
					<%Else%>
						<option value="<%=EventDir(0, i)%>"><%=EventDir(1, i)%></option>
					<%End If%>
				<%Next%>
			</select>
			<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
			<input class="form-control" type="submit" name="submit2" id="submit2" value="Get Event Director">
			</form>
			<br>
			<%If Not CLng(lEventDir) = 0 Then%>
				<h4 class="h4">My Data</h4>

				<form name="edit_event_dir" method="Post" action="edit_event_dir.asp?event_dir=<%=lEventDir%>">
				<table class="table">
					<tr>
						<th>First Name:</th>
						<td><input type="text" name="first_name" id="first_name" value="<%=EventDirArr(0)%>"></td>
						<th>Last Name:</th>
						<td><input type="text" name="last_name" id="last_name" value="<%=EventDirArr(1)%>"></td>
					</tr>
					<tr>
						<th>Address:</th>
						<td><input type="text" name="address" id="address" value="<%=EventDirArr(3)%>"></td>
						<th>City:</th>
						<td><input type="text" name="city" id="city" value="<%=EventDirArr(4)%>"></td>
					</tr>
					<tr>
						<th>State:</th>
						<td><input type="text" name="state" id="state" size="2" value="<%=EventDirArr(5)%>"></td>
						<th>Zip:</th>
						<td><input type="text" name="zip" id="zip" size="7" value="<%=EventDirArr(6)%>"></td>
					</tr>
					<tr>
						<th>Phone:</th>
						<td><input type="text" name="phone" id="phone" value="<%=EventDirArr(2)%>"></td>
						<th>Email:</th>
						<td><input type="text" name="email" id="email" value="<%=EventDirArr(7)%>"></td>
					</tr>
					<tr>
						<th style="white-space: nowrap;">User Name:</th>
						<td><input type="text" name="user_id" id="user_id" value="<%=EventDirArr(8)%>" maxlength="12"></td>
						<th>Password:</th>
						<td><input type="text" name="password" id="password" value="<%=EventDirArr(9)%>" maxlength="12"></td>
					</tr>
					<tr>
						<th>Active:</th>
						<td>
							<select name="active" id="active">
								<%If EventDirArr(11) = "y" Then%>
									<option value="y" selected>Yes</option>
									<option value="n">No</option>
								<%Else%>
									<option value="y">Yes</option>
									<option value="n" selected>No</option>
								<%End If%>
							</select>
						</td>
						<th>Mobile:</th>
						<td><input type="text" name="mobile" id="mobile" value="<%=EventDirArr(12)%>" maxlength="12"></td>
					</tr>
					<tr>
						<th valign="top">Comments:</th>
						<td colspan="3"><textarea name="comments" id="comments" cols="60" rows="3"><%=EventDirArr(10)%></textarea></td>
					</tr>
					<tr>
						<td style="border:1px solid #ececd8;text-align:center;color: red;" colspan="4">
							<input type="checkbox" name="delete" id="delete">&nbsp;Delete This (THERE IS NO UNDO!)
						</td>
					</tr>
					<tr>
						<td style="background-color:#ececd8;text-align:center;" colspan="4">
							<input type="hidden" name="submit_event_dir" id="submit_event_dir" value="submit_event_dir">
							<input type="submit" name="submit1" id="submit1" value="Submit Changes">
						</td>
					</tr>
				</table>
				</form>

				<%Call MyEvents(lEventDir)%>
				<h4 class="h4">My Races</h4>

				<table class="table table-striped">
					<tr>
						<th>ID</th>
						<th>Event</th>
						<th>Date</th>
						<th>Grp</th>
						<th>Edition</th>
					</tr>
					<%For i = 0 To UBound(Events, 2) - 1%>
						<tr>
							<%For j = 0 To 4%>
								<td><%=Events(j, i)%></td>
							<%Next%>
						</tr>
					<%Next%>
				</table>
			<%End If%>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>