<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lStaffID
Dim StaffArr(13), Staff()
Dim sFirstName, sLastName, sAddress, sCity, sState, sZip, sPhone, sEmail, sUserID, sPassword, sComments, sTech, sSupport, sActive

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lStaffID = Request.QueryString("staff_id")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Staff(1, 0)
sql = "SELECT StaffID, FirstName, LastName FROM Staff ORDER BY LastName, FirstName"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Staff(0, i) = rs(0).Value
	Staff(1, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
	i = i + 1
	ReDim Preserve Staff(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("submit_event_dir") = "submit_event_dir" Then
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
	sTech = Request.Form.Item("tech")
    sSupport = Request.Form.Item("support")
    sActive = Request.Form.Item("active")

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Address, City, State, Zip, Email, Phone, UserID, Password, Comments, Tech, Support, Active "
    sql = sql & "FROM Staff WHERE StaffID = " & lStaffID
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
	rs(11).Value = sTech
	rs(12).Value = sSupport
    rs(13).Value = sActive

	rs.Update
	rs.Close
	Set rs = Nothing
ElseIf Request.Form.Item("submit_this") = "submit_this" Then
	lStaffID = Request.Form.Item("staff")
End If

If CStr(lStaffID) = vbNullString Then lStaffID = 0

If Not CLng(lStaffID) = 0 Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT FirstName, LastName, Phone, Address, City, State, Zip, Email, UserID, Password, Comments, Tech, Support, Active "
    sql = sql & "FROM Staff WHERE StaffID = " & lStaffID
	rs.Open sql, conn, 1, 2
	For i = 0 to 13
		If not rs(i).Value & "" = "" Then StaffArr(i) =  Replace(rs(i).Value, "''", "'")
	Next
	rs.Close
	Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE Edit Staff</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
			<h4 class="h4">Edit Staff</h4>			
		    <div class="row">
			    <form role="form" class="form-inline" name="edit_staff" method="Post" action="edit_staff.asp">
			    <select class="form-control" name="staff" id="staff" onchange="this.form.submit2.click();" style="font-size: 0.85em;">
				    <option value="">&nbsp;</option>
				    <%For i = 0 To UBound(Staff, 2) - 1%>
					    <%If CLng(Staff(0, i)) = CLng(lStaffID) Then%>
						    <option value="<%=Staff(0, i)%>" selected><%=Staff(1, i)%></option>
					    <%Else%>
						    <option value="<%=Staff(0, i)%>"><%=Staff(1, i)%></option>
					    <%End If%>
				    <%Next%>
			    </select>
			    <input type="hidden" name="submit_this" id="submit_this" value="submit_this">
			    <input class="form-control" type="submit" name="submit2" id="submit2" value="Get Staff Member" style="font-size: 0.8em;">
			    </form>
		    </div>
			
		    <%If Not CLng(lStaffID) = 0 Then%>
			    <form name="edit_staff" method="Post" action="edit_staff.asp?staff_id=<%=lStaffID%>">
			    <table style="margin:10px;font-size: 0.85em;">
				    <tr>
					    <th>First Name:</th>
					    <td><input type="text" name="first_name" id="first_name" value="<%=StaffArr(0)%>"></td>
					    <th>Last Name:</th>
					    <td><input type="text" name="last_name" id="last_name" value="<%=StaffArr(1)%>"></td>
				    </tr>
				    <tr>
					    <th>Address:</th>
					    <td><input type="text" name="address" id="address" value="<%=StaffArr(3)%>" size="30"></td>
					    <th>City:</th>
					    <td><input type="text" name="city" id="city" value="<%=StaffArr(4)%>"></td>
				    </tr>
				    <tr>
					    <th>State:</th>
					    <td><input type="text" name="state" id="state" size="2" value="<%=StaffArr(5)%>"></td>
					    <th>Zip:</th>
					    <td><input type="text" name="zip" id="zip" size="7" value="<%=StaffArr(6)%>"></td>
				    </tr>
				    <tr>
					    <th>Phone:</th>
					    <td><input type="text" name="phone" id="phone" value="<%=StaffArr(2)%>"></td>
					    <th>Email:</th>
					    <td><input type="text" name="email" id="email" value="<%=StaffArr(7)%>" size="30"></td>
				    </tr>
				    <tr>
					    <th>User Name:</th>
					    <td><input type="text" name="user_id" id="user_id" value="<%=StaffArr(8)%>" maxlength="12"></td>
					    <th>Password:</th>
					    <td><input type="text" name="password" id="password" value="<%=StaffArr(9)%>" maxlength="12"></td>
				    </tr>
				    <tr>
					    <th>Tech:</th>
					    <td>                   
                            <select name="tech" id="tech">
                                <%If StaffArr(11) = "y" Then%>
                                    <option value="y" selected>Yes</option>
                                    <option value="n">No</option>
                                <%Else%>
                                    <option value="y">Yes</option>
                                    <option value="n" selected>No</option>
                                <%End If%>
                            </select>
                        </td>
					    <th>Support:</th>
					    <td>                   
                            <select name="support" id="support">
                                <%If StaffArr(12) = "y" Then%>
                                    <option value="y" selected>Yes</option>
                                    <option value="n">No</option>
                                <%Else%>
                                    <option value="y">Yes</option>
                                    <option value="n" selected>No</option>
                                <%End If%>
                            </select>
                        </td>
				    </tr>
				    <tr>
					    <th>Active:</th>
					    <td colspan="3">                   
                            <select name="active" id="active">
                                <%If StaffArr(13) = "y" Then%>
                                    <option value="y" selected>Yes</option>
                                    <option value="n">No</option>
                                <%Else%>
                                    <option value="y">Yes</option>
                                    <option value="n" selected>No</option>
                                <%End If%>
                            </select>
                        </td>
				    </tr>
				    <tr>
					    <th valign="top">Comments:</th>
					    <td colspan="3"><textarea name="comments" id="comments" cols="60" rows="3"><%=StaffArr(10)%></textarea></td>
				    </tr>
				    <tr>
					    <td style="background-color:#ececd8;text-align:center;" colspan="4">
						    <input type="hidden" name="submit_event_dir" id="submit_event_dir" value="submit_event_dir">
						    <input type="submit" name="submit1" id="submit1" value="Submit Changes">
					    </td>
				    </tr>
			    </table>
			    </form>
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