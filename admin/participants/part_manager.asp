<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim sSortBy, sAlpha
Dim PartArray, Alpha(25)

sSortBy = Request.QueryString("sort_by")
If sSortBy = vbNullString Then sSortBy = "LastName, FirstName, Email"

sAlpha = Request.QueryString("alpha")
If sAlpha = vbNullString Then sAlpha = "A"

Alpha(0) = "A"
Alpha(1) = "B"
Alpha(2) = "C"
Alpha(3) = "D"
Alpha(4) = "E"
Alpha(5) = "F"
Alpha(6) = "G"
Alpha(7) = "H"
Alpha(8) = "I"
Alpha(9) = "J"
Alpha(10) = "K"
Alpha(11) = "L"
Alpha(12) = "M"
Alpha(13) = "N"
Alpha(14) = "O"
Alpha(15) = "P"
Alpha(16) = "Q"
Alpha(17) = "R"
Alpha(18) = "S"
Alpha(19) = "T"
Alpha(20) = "U"
Alpha(21) = "V"
Alpha(22) = "W"
Alpha(23) = "X"
Alpha(24) = "Y"
Alpha(25) = "Z"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_filter") = "submit_filter" Then
    sSortBy = Request.Form.Item("sort_by")
    sAlpha = Request.Form.Item("alpha")
End If

i = 0
ReDim PartArray(10, 0)			
Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT ParticipantID, FirstName, LastName, Gender, City, St, Phone, DOB, Email FROM Participant WHERE LastName LIKE '" & sAlpha & "%' ORDER BY " & sSortBy
rs.Open sql, conn, 1, 2
PartArray = rs.GetRows()
rs.Close
Set rs=Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->

<title>GSE Participant Manager</title>

<!--#include file = "../../includes/js.asp" -->

<style type="text/css">
	th{
		text-align:left;
		white-space:nowrap;
		padding:0 5px 0 0;
	}
	
	td{
		padding:0 5px 0 0;
	}
</style>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE Participant Data Manager</h4>

            <div style="margin: 10px 0 10px 0;font-size: 0.9em;background-color: #ececd8;">
                <form name="filter_data" method="post" action="part_manager.asp?alpha=<%=sAlpha%>">
                Sort By:
                <select name="sort_by" id="sort_by">
                    <%If sSortBy = "LastName, FirstName, Email" Then%>
                        <option value="LastName, FirstName, Email" selected>LastName, FirstName, Email</option>
                        <option value="Email, LastName, FirstName">Email, LastName, FirstName</option>
                    <%Else%>
                        <option value="LastName, FirstName, Email">LastName, FirstName, Email</option>
                        <option value="Email, LastName, FirstName" selected>Email, LastName, FirstName</option>
                    <%End If%>
                </select>
                &nbsp;&nbsp;
                Alpha:
                <select name="alpha" id="alpha">
                    <%For i = 0 To 25%>
                        <%If sAlpha = Alpha(i) Then%>
                            <option value="<%=Alpha(i)%>" selected><%=Alpha(i)%></option>
                        <%Else%>
                            <option value="<%=Alpha(i)%>"><%=Alpha(i)%></option>
                        <%End If%>
                    <%Next%>
                </select>
                 &nbsp;&nbsp;
                 <input type="hidden" name="submit_filter" id="submit_filter" value="submit_filter">
                <input type="submit" name="submit1" id="submit1" value="Filer Data">
               </form>
            </div>

			<table style="border-collapse:collapse;width:800px;font-size:0.85em;">
				<tr>
					<th style="text-align:center;width:10px">No</th>
					<th>Name</th>
					<th>M/F</th>
					<th>City</th>
					<th>St</th>
					<th>Phone</th>
					<th>DOB</th>
					<th>Email</th>
				</tr>
				<%For j = 0 to UBound(PartArray, 2)%>
					<%If j mod 2 = 0 Then%>
						<tr>
							<td class="alt"><%=j+1%>)</td>
                            <td class="alt" style="white-space:nowrap">
                                <a href="javascript:pop('part_details.asp?part_id=<%=PartArray(0, j)%>',800,600)"><%=PartArray(2, j)%>, <%=PartArray(1, j)%></a>
                            </td>
                            <td class="alt" style="white-space:nowrap"><%=PartArray(3, j)%></td>
                            <td class="alt" style="white-space:nowrap"><%=PartArray(4, j)%></td>
                            <td class="alt" style="white-space:nowrap"><%=PartArray(5, j)%></td>
                            <td class="alt" style="white-space:nowrap"><%=PartArray(6, j)%></td>
                            <td class="alt" style="white-space:nowrap"><%=PartArray(7, j)%></td>
                            <td class="alt" style="white-space:nowrap"><%=PartArray(8, j)%></td>						
                        </tr>
					<%Else%>
						<tr>
							<td><%=j+1%>)</td>
                            <td style="white-space:nowrap">
                                <a href="javascript:pop('part_details.asp?part_id=<%=PartArray(0, j)%>',800,600)"><%=PartArray(2, j)%>, <%=PartArray(1, j)%></a>
                            </td>
                            <td style="white-space:nowrap"><%=PartArray(3, j)%></td>
                            <td style="white-space:nowrap"><%=PartArray(4, j)%></td>
                            <td style="white-space:nowrap"><%=PartArray(5, j)%></td>
                            <td style="white-space:nowrap"><%=PartArray(6, j)%></td>
                            <td style="white-space:nowrap"><%=PartArray(7, j)%></td>
                            <td style="white-space:nowrap"><%=PartArray(8, j)%></td>						
						</tr>
					<%End If%>
				<%Next%>
			</table>
		</div>
	</div>
	<!--#include file = "../../includes/footer.asp" -->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>