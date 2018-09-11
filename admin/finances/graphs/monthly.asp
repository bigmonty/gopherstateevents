<%@ Language=VBScript %>
<%
Option Explicit

Dim i
Dim sWhichGraph
Dim Graphs(4)

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Graphs(0) = "Income"
Graphs(1) = "Expenses"
Graphs(2) = "Profit"
Graphs(3) = "Events"
Graphs(4) = "Margin"

If Request.Form.Item("submit_graph") = "submit_graph" Then
    sWhichGraph = Request.Form.Item("which_graph")
    If Request.Form.Item("show_avg") = "on" Then sShowAvg = "n"
End If

If sWhichGraph = vbNullString Then sWhichGraph = "Events"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Finance Graphs: Monthly</title>
</head>

<body>
<div class="container">
  	<!--#include file = "../../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
            <!--#include file = "graphs_menu.asp" -->

            <h3 class="h3">Finance Graphs: Monthly Graphs</h3>

            <form class="form-inline" name="get_monthly" method="post" action="monthly.asp">
            <label for="staff">Select Graph:</label>&nbsp;
            <select class="form-control" name="which_graph" id="which_graph" onchange="this.form.submit1.click();">
                <option value=""></option>
                <%For i = 0 To UBound(Graphs)%>
                    <%If CStr(Graphs(i)) = CStr(sWhichGraph) Then%>
                        <option value="<%=Graphs(i)%>" selected><%=Graphs(i)%></option>
                    <%Else%>
                        <option value="<%=Graphs(i)%>"><%=Graphs(i)%></option>
                    <%End If%>
                <%Next%>
            </select>&nbsp;&nbsp;
            <input type="hidden" name="submit_graph" id="submit_graph" value="submit_graph">
            <input type="submit" class="form-control" name="submit1" id="submit1" value="View This Graph">
            </form>
            <div class="embed-responsive embed-responsive-16by9">
				<iframe name="yearly_graph" id="yearly_graph" frameborder="0" 
					src="monthly_graph.asp?which_graph=<%=sWhichGraph%>" style="width:800px;height:400px;"></iframe>
            </div>
        </div>
	</div>
</div>
<!--#include file = "../../../includes/footer.asp" -->
</body>
</html>
