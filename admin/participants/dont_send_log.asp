<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim DontSend()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim DontSend(2, 0)
Set rs = SErver.CreateObject("ADODB.Recordset")
sql = "SELECT DontSendID, Email, WhenEntered FROM DontSend ORDER BY WhenEntered DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    DontSend(0, i) = rs(0).Value
    DontSend(1, i) = rs(1).Value
    DontSend(2, i) = rs(2).Value
    i = i + 1
    ReDim Preserve DontSend(2, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_remove") = "submit_remove" Then
    For i = 0 To UBound(DontSend, 2) - 1
        If Request.Form.Item("remove_" & DontSend(0, i)) = "y" Then
            sql = "DELETE FROM DontSend WHERE DontSendID = " & DontSend(0, i)
            Set rs = conn.Execute(sql)
            Set rs = Nothing
        End If
    Next
ElseIf Request.Form.Item("submit_email") = "submit_email" Then
	sEmail = Request.Form.Item("email")

    If ValidEmail = True Then
	    sql = "INSERT INTO DontSend(Email, WhenEntered) VALUES ('" & sEmail & "', '" & Now() & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If

    i = 0
    ReDim DontSend(2, 0)
    Set rs = SErver.CreateObject("ADODB.Recordset")
    sql = "SELECT DontSendID, Email, WhenEntered FROM DontSend ORDER BY WhenEntered DESC"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        DontSend(0, i) = rs(0).Value
        DontSend(1, i) = rs(1).Value
        DontSend(2, i) = rs(2).Value
        i = i + 1
        ReDim Preserve DontSend(2, i)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If

%>
<!--#include file = "../../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>GSE&copy; Dont Send Log</title>
<!--#include file = "../../includes/js.asp" -->
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">Dont Send Log</h4>

            <div style="margin: 10px;">
                <h4 class="h4">Insert New</h4>

                <br>

	            <form name="login" method="post" action="dont_send_log.asp">
	            <span style="font-weight:bold;font-size:0.9em;">Email address:</span>
	            <input type="text" name="email" id="email" size="30">
	            <input type="hidden" name="submit_email" id="submit_email" value="submit_email">
	            <input type="submit" name="submit1" id="submit1" value="Submit Email">
	            </form>    
            </div>

            <hr>

            <div style="margin: 10px;">
                <h4 class="h4">Existing</h4>
 
                <br>

               <form name="remove_from_list" method="post" action="dont_send_log.asp">
                <table>
                    <tr>
                        <th>No.</th>
                        <th>Email</th>
                        <th>When Sent</th>
                        <th>Remove</th>
                    </tr>
                    <%For i = 0 To UBound(DontSend, 2) - 1%>
                        <%If i mod 2 = 0 Then%>
                            <tr>
                                <td class="alt"><%=i + 1%></td>
                                <td class="alt"><a href="mailto:<%=DontSend(1, i)%>"><%=DontSend(1, i)%></a></td>
                                <td class="alt"><%=DontSend(2, i)%></td>
                                <td class="alt">
                                    <select name="remove_<%=DontSend(0, i)%>" id="remove_<%=DontSend(0, i)%>">
                                        <option value="n">No</option>
                                        <option value="y">Yes</option>
                                    </select>
                                </td>
                            </tr>
                        <%Else%>
                            <tr>
                                <td><%=i + 1%></td>
                                <td><a href="mailto:<%=DontSend(1, i)%>"><%=DontSend(1, i)%></a></td>
                                <td><%=DontSend(2, i)%></td>
                                <td>
                                    <select name="remove_<%=DontSend(0, i)%>" id="remove_<%=DontSend(0, i)%>">
                                        <option value="n">No</option>
                                        <option value="y">Yes</option>
                                    </select>
                                </td>
                            </tr>
                        <%End If%>
                    <%Next%>
				    <tr>
					    <td style="text-align:center;" colspan="3">
						    <input type="hidden" name="submit_remove" id="submit_remove" value="submit_remove">
						    <input type="submit" name="submit1" id="submit1" value="Remove Selected">
					    </td>
				    </tr>
                </table>
                </form>
            </div>
		</div>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>
