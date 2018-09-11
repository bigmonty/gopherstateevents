<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim OrderChips(), OrderBibs(), OrderSpacers(), OrderPins()

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_spacers") = "submit_spacers" Then
    i = 0
    ReDim Delete(0)

    For i = 0 To UBound(Delete) - 1
    Next
ElseIf Request.Form.Item("submit_new_spacers") = "submit_new_spacers" Then
ElseIf Request.Form.Item("submit_pins") = "submit_pins" Then
    i = 0
    ReDim Delete(0)

    For i = 0 To UBound(Delete) - 1
    Next
ElseIf Request.Form.Item("submit_new_pins") = "submit_new_pins" Then
    sql = "INSERT INTO InventoryPins(NumPins, Location) VALUES (" & Request.Form.Item("num_pins") & ", '" & Request.Form.Item("pins_location") 
    sql = sql & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_prepped") = "submit_prepped" Then
    i = 0
    ReDim Delete(0)

    For i = 0 To UBound(Delete) - 1
    Next
ElseIf Request.Form.Item("submit_new_prepped") = "submit_new_prepped" Then
ElseIf Request.Form.Item("submit_bib") = "submit_bib" Then
    i = 0
    ReDim Delete(0)

    For i = 0 To UBound(Delete) - 1
    Next
ElseIf Request.Form.Item("submit_new_bibs") = "submit_new_bibs" Then
ElseIf Request.Form.Item("submit_chip") = "submit_chip" Then
    i = 0
    ReDim Delete(0)

    For i = 0 To UBound(Delete) - 1
    Next
ElseIf Request.Form.Item("submit_new_chips") = "submit_new_chips" Then
End If

i = 0
ReDim OrderChips(7, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT OrderChipsID, StartChip, EndChip, Amount, WhenOrdered, Vendor, OrderedBy, PmtProcessed FROM OrderChips ORDER BY WhenOrdered DESC"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    For j = 0 To 7
        If Not rs(j).Value & "" = "" Then OrderChips(j, i) = rs(j).Value
    Next
    i = i + 1
    ReDim Preserve OrderChips(7, i)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>GSE&copy; Orders</title>
<
<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td,th{
        padding-right: 5px;
    }
</style>
</head>

<<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE Media Orders</h4>
         </div>
	</div>
</div>	
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
