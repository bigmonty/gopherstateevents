<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisFamily
Dim sBibType, sBibFamily
Dim BibFamilies(), BibRange(), BibType(3)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

ReDim BibRange(1, 0)
ReDim BibFamilies(1, 0)

BibType(0) = "Tyvek Running"
BibType(1) = "Bike Plates"
BibType(2) = "Nordic"
BibType(3) = "Tri Tags"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_family") = "submit_family" Then
    sBibFamily = Request.Form.Item("bib_family")
    sBibType = Request.Form.Item("bib_type")

    sql = "INSERT INTO BibFamilies (BibFamily, BibType) VALUES ('" & sBibFamily & "', '" & sBibType & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
    i = 0
    ReDim Delete(0)
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT TeamBibsID, FirstBib, LastBib FROM TeamBibs WHERE TeamsID = " & lThisTeam
    rs.Open sql, conn, 1,  2
    Do While Not rs.EOF
        If Request.Form.Item("delete_" & rs(0).Value) = "on" Then
            Delete(i) = rs(0).Value
            i = i + 1
            ReDim Preserve Delete(i)
        Else
            rs(1).Value = Request.Form.Item("first_bib_" & rs(0).Value)
            rs(2).Value = Request.Form.Item("last_bib_" & rs(0).Value)
            rs.Update
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    For i = 0 To UBound(Delete) - 1
        sql = "DELETE FROM TeamBibs WHERE TeamBibsID = " & Delete(i)
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    Next
    Set rs = Nothing
ElseIf Request.Form.Item("submit_range") = "submit_range" Then
	iLastBib = Request.Form.Item("last_bib")
	iFirstBib = Request.Form.Item("first_bib")

	sql = "INSERT INTO TeamBibs (TeamsID, FirstBib, LastBib) VALUES (" & lThisTeam & ", " & iFirstBib & ", " & iLastBib & ")"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->
<title>GSE Bib Inventory</title>
</head>

<body onload="javascript:set_bib_range.first_bib.focus()">
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../includes/admin_menu.asp" -->

		<div class="col-sm-10">
            <h4 class="h4">GSE Bib Inventory</h4>

            <br>
            
            <h5 class="h5">Add Bib Family</h5>
            <form class="form-inline" role="form" name="add_family" method="Post" action="bib_inventory.asp">
            <label for="bib_family">Bib Family</label>
            <input class="form-control" type="text" name="bib_family" id="bib_family">
            <label for="bib_type">Bib Type</label>
            <select class="form-control" name="bib_type" id="bib_type">
                <option value="">&nbsp;</option>
                <%For i = 0 To UBound(BibType)%>
                    <option value="<%=BibType(i)%>"><%=BibType(i)%></option>
                <%Next%>
            </select>
            <input type="hidden" name="submit_family" id="submit_family" value="submit_family">
            <input class="form-control" type="submit" name="submit_1" id="submit_1" value="Create New Family">
            </form>

            <hr>

            <h5 class="h5">Manage Inventory</h5>

            <%If lThisFamily > 0 Then%>
                <form name="edit_bib_range" method="Post" action="bib_inventory.asp?this_family=<%=lThisFamily%>">
                <table>
                    <tr><th colspan="4">Existing Bib Ranges:</th></tr>
                    <%For i = 0 To UBound(BibRange, 2) - 1%>
                        <tr>
                            <th>From</th>
                            <td><input type="text" name="first_bib_<%=BibRange(0, i)%>" id="first_bib_<%=BibRange(0, i)%>" size="4" value="<%=BibRange(1, i)%>"></td>
                            <th>To</th>
                            <td><input type="text" name="last_bib_<%=BibRange(0, i)%>" id="last_bib_<%=BibRange(0, i)%>" size="4" value="<%=BibRange(2, i)%>"></td>
                            <td style="color: red;"><input type="checkbox" name="delete_<%=BibRange(0, i)%>" id="delete_<%=BibRange(0, i)%>">Delete</td>
                        </tr>
                    <%Next%>
                    <tr>
                        <td  style="text-align:center;" colspan="4">
                            <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
                            <input type="submit" name="submit2" id="submit2" value="Submit Changes">
                        </td>
                    </tr>
                </table>
                </form>
            <%End If%>
        </div>
    </div>
</div>

<!--#include file = "../includes/footer.asp" -->
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
