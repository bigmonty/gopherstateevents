<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim iYear
Dim sItemName, sItemType, sActive, sComments
Dim sngUnitCost
Dim Items(), ItemTypes(3)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

ItemTypes(0) = "Participant Costs"
ItemTypes(1) = "Labor Costs"
ItemTypes(2) = "Event Costs"
ItemTypes(3) = "Miscellaneous"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

ReDim Items(5, 0)

If Request.Form.Item("submit_item") = "submit_item" Then
    sItemName = Replace(Request.Form.Item("item_name"), "'", "''")
    sItemType = Request.Form.Item("item_type")
    sngUnitCost = Request.Form.Item("unit_cost")

    sComments = Request.Form.Item("comments")
    If Not sComments = vbNullString Then sComments = Replace(sComments, "'", "''")

    sql = "INSERT INTO FinanceItems(ItemName, ItemType, UnitCost, Comments) VALUES ('" & sItemName & "', '" & sItemType & "', " & sngUnitCost
    sql = sql & ", '" & sComments & "')"
    Set rs = conn.Execute(sql)
    Set rs = Nothing
ElseIf Request.Form.Item("submit_changes") = "submit_changes" Then
    Call GetItems()

    For i = 0 To UBound(Items, 2) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT ItemName, ItemType, UnitCost, Active, Comments FROM FinanceItems WHERE FinanceItemsID = " & Items(0, i)
        rs.Open sql, conn, 1, 2
        If Request.Form.Item("item_name_" & Items(0, i)) & "" = "" Then
            rs(0).Value = rs(0).OriginalValue
        Else
            rs(0).Value = Replace(Request.Form.Item("item_name_" & Items(0, i)), "'", "''")
        End If

        rs(1).Value = Request.Form.Item("item_type_" & Items(0, i))

        If Request.Form.Item("unit_cost_" & Items(0, i)) & "" = "" Or Request.Form.Item("unit_cost_" & Items(0, i)) = "0" Then
            rs(2).Value = rs(0).OriginalValue
        Else
            rs(2).Value = Request.Form.Item("unit_cost_" & Items(0, i))
        End If

        rs(3).Value = Request.Form.Item("active_" & Items(0, i))

        If Request.Form.Item("comments_" & Items(0, i)) & "" = "" Then
            rs(4).Value = Null
        Else
            rs(4).Value = Replace(Request.Form.Item("comments_" & Items(0, i)), "'", "''")
        End If

        rs.Update
        rs.Close
        Set rs = Nothing
    Next
End If

Call GetItems()

Private Sub GetItems()
    i = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT FinanceItemsID, ItemName, ItemType, UnitCost, Comments, Active FROM FinanceItems ORDER BY Active DESC, ItemType, ItemName"
    rs.Open sql, conn, 1, 2
    Do While Not rs.EOF
        Items(0, i) = rs(0).Value
        Items(1, i) = Replace(rs(1).Value, "''", "'")
        Items(2, i) = rs(2).Value
        Items(3, i) = rs(3).Value
        Items(4, i) = rs(4).Value
        Items(5, i) = rs(5).Value
        i = i + 1 
        ReDim Preserve Items(5, i)
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
<title>GSE Finances: Finance Items</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
            <h3 class="h3">GSE Finances: Finance Items</h3>            

            <div class="row">
		        <h4 class="h4">Add Finance Item</h4>
			
		        <form class="form-inline" name="add_item" method="Post" action="items.asp">
			    <label for="item_name">Name:</label>
			    <input type="text" class="form-control input-sm" name="item_name" id="item_name">
			    <label for="item_type">Type:</label>
                <select class="form-control input-sm" name="item_type" id="item_type">
                    <option value=""></option>
                    <%For i = 0 To UBound(ItemTypes)%>
                        <option value="<%=ItemTypes(i)%>"><%=ItemTypes(i)%></option>
                    <%Next%>
                </select>
			    <label for="unit_cost">Unit Cost:</label>
			    <input type="text" class="form-control input-sm" name="unit_cost" id="unit_cost">
			    <label for="comments">Cmnts:</label>
			    <input type="text" class="form-control input-sm" name="comments" id="comments">
			    <input type="hidden" name="submit_item" id="submit_item" value="submit_item">
			    <input type="submit" class="form-control input-sm" name="submit1" id="submit1" value="Submit">
		        </form>

                <hr>
            </div>

		    <h4 class="h4">Edit Finance Items</h4>
			
		    <form class="form" name="add_staff" method="Post" action="items.asp">
		    <table class="table">
			    <tr>
				    <th>Item Name</th>
				    <th>Item Type</th>
				    <th>Unit Cost</th>
                    <th>Comments</th>
                    <th>Active</th>
                </tr>
                <%For i = 0 To UBound(Items, 2) - 1%>
                    <%If Items(5, i) = "n" Then%>
			            <tr>
				            <td><%=Items(1, i)%></td>
				            <td><%=Items(2, i)%></td>
				            <td><%=Items(3, i)%></td>
                            <td><%=Items(4, i)%></td>
                            <td>
                                <select class="form-control input-sm" name="active_<%=Items(0, i)%>" id="active_<%=Items(0, i)%>">
                                    <%If Items(5, i) = "y" Then%>
                                        <option value="y" selected>y</option>
                                        <option value="n">n</option>
                                    <%Else%>
                                        <option value="y">y</option>
                                        <option value="n" selected>n</option>
                                    <%End If%>
                                </select>
                            </td>
                        </tr>
                    <%Else%>
			            <tr>
				            <td>
                                <input type="text" class="form-control input-sm" name="item_name_<%=Items(0, i)%>" id="item_name_<%=Items(0, i)%>" value="<%=Items(1, i)%>">
                            </td>
				            <td>
                                <select class="form-control input-sm" name="item_type_<%=Items(0, i)%>" id="item_type_<%=Items(0, i)%>">
                                    <%For j = 0 To UBound(ItemTypes)%>
                                        <%If Items(2, i) = ItemTypes(j) Then%>
                                            <option value="<%=ItemTypes(j)%>" selected><%=ItemTypes(j)%></option>
                                        <%Else%>
                                            <option value="<%=ItemTypes(j)%>"><%=ItemTypes(j)%></option>
                                        <%End If%>
                                    <%Next%>
                                </select>
                            </td>
				            <td>
                                <input type="text" class="form-control input-sm" name="unit_cost_<%=Items(0, i)%>" id="unit_cost_<%=Items(0, i)%>" value="<%=Items(3, i)%>">
                            </td>
                            <td>
                                <input type="text" class="form-control input-sm" name="comments_<%=Items(0, i)%>" id="comments_<%=Items(0, i)%>" value="<%=Items(4, i)%>">
                            </td>
                            <td>
                                <select class="form-control input-sm" name="active_<%=Items(0, i)%>" id="active_<%=Items(0, i)%>">
                                    <%If Items(5, i) = "y" Then%>
                                        <option value="y" selected>y</option>
                                        <option value="n">n</option>
                                    <%Else%>
                                        <option value="y">y</option>
                                        <option value="n" selected>n</option>
                                    <%End If%>
                                </select>
                            </td>
                        </tr>
                    <%End If%>
                <%Next%>
                <tr>
				    <td colspan="5">
					    <input type="hidden" name="submit_changes" id="submit_changes" value="submit_changes">
					    <input type="submit" class="form-control input-sm" name="submit2" id="submit2" value="Submit Changes">
				    </td>
			    </tr>
		    </table>
		    </form>
        </div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
<%	
conn.Close
Set conn = Nothing
%>
</body>
</html>
