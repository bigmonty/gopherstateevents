<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim Accounts(), SortArr(10)
Dim i, j, k

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

Dim conn2	
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.ConnectionTimeout = 30
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=etraxc;Uid=broad_user;Pwd=Zeroto@123;"

i = 0
ReDim Accounts(10, 0)
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT MyHistID, UserName, Password, WhenCreated FROM MyHist"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	Accounts(0, i) = rs(0).Value
    Accounts(8, i) = rs(1).Value
    Accounts(9, i) = rs(2).Value
    Accounts(10, i) = rs(3).Value
	i = i + 1
	ReDim Preserve Accounts(10, i)
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

For i = 0 To UBound(Accounts, 2) - 1
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT pd.PartID, pd.FirstName, pd.LastName, pd.Gender, pd.ScreenName, pd.Email, pp.City, pp.St FROM PartData pd INNER JOIN PartProfile pp "
    sql = sql & "ON pd.PartID = pp.PartID WHERE MyHistID = " & Accounts(0, i)
	rs.Open sql, conn2, 1, 2
    If rs.RecordCount > 0 Then
        Accounts(1, i) = rs(0).Value
        Accounts(2, i) = Replace(rs(2).Value, "''", "'") & ", " & Replace(rs(1).Value, "''", "'")
        Accounts(3, i) = rs(3).Value
        Accounts(4, i) = rs(4).Value
        Accounts(5, i) = rs(5).Value
        Accounts(6, i) = rs(6).Value
        Accounts(7, i) = rs(7).Value
    End If
	rs.Close
	Set rs = Nothing
Next

For i = 0 To UBound(Accounts, 2) - 2
    For j = i + 1 To UBound(Accounts, 2) - 1
        If CStr(UCase(Accounts(2, i))) > CStr(UCAse(Accounts(2, j))) Then
            For k = 0 To 10
                SortArr(k) = Accounts(k, i)
                Accounts(k, i) = Accounts(k, j)
                Accounts(k, j) = SortArr(k)
            Next
        End If
    Next
Next
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE My History Accounts</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
    td,th{padding-right: 5px;}
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <!--#include file = "my_hist_nav.asp" -->

		    <h4 class="h4">My History Accounts</h4>
		
		    <table style="font-size:0.85em;margin:10px 0 0 0;">
			    <tr>
				    <th style="border-bottom:1px solid silver;text-align:right;border-bottom:1px solid silver;">No.</th>
				    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">Name</th>
                    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">M/F</th>
				    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">Screen Name</th>
                    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">Email</th>
				    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">City</th>
				    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">St</th>
                    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">User Name</th>
                    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">Password</th>
				    <th style="border-bottom:1px solid silver;border-bottom:1px solid silver;">When Reg</th>
			    </tr>
			    <%For i = 0 to UBound(Accounts, 2) - 1%>
				    <tr>
					    <%If i mod 2 = 0 Then%>
						    <td class="alt" style="text-align:right">
							    <%=i + 1%>)
						    </td>
						    <%For j = 2 to 10%>
							    <td class="alt">
								    <%If j = 5 Then%>
									    <a href="mailto:<%=Accounts(5, i)%>"><%=Accounts(5, i)%></a>
								    <%Else%>
									    <%=Accounts(j, i)%>
								    <%End If%>
							    </td>
						    <%Next%>
					    <%Else%>
						    <td style="text-align:right;">
							    <%=i + 1%>)
						    </td>
						    <%For j = 2 to 10%>
							    <td>
								    <%If j = 5 Then%>
									    <a href="mailto:<%=Accounts(5, i)%>"><%=Accounts(5, i)%></a>
								    <%Else%>
									    <%=Accounts(j, i)%>
								    <%End If%>
							    </td>
						    <%Next%>
					    <%End If%>
				    </tr>
			    <%Next%>
		    </table>
        </div>
	</div>
</div>
<%
conn2.Close
Set conn2 = Nothing

conn.Close
Set conn = Nothing
%></body>
</html>
