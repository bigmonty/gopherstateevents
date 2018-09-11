<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j, k
Dim PartList

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'If Request.Form.Item("submit_this") = "submit_this" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT p.FirstName, p.LastName, p.City, p.St, p.Email, pr.RaceID FROM Participant p INNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID "
    sql = sql & "WHERE p.Email IS NOT NULL ORDER BY p.Email"
    rs.Open sql, conn, 1, 2
    PartList = rs.GetRows()
    rs.Close
    Set rs = Nothing
'End If

'trim entries
For i = 0 To UBound(PartList, 2)
    For j = 0 To 5
        If PartList(j, i) & "" = "" Then 
            PartList(j, i) = "empty"
        Else
            Trim(PartList(j, i))
        End If
    Next
Next

'remove duplicates
k = 0
Dim sOldEmail
Dim BestList(5, 20000)

For i = 0 To UBound(PartList, 2)
    If i = 0 Then 
        sOldEmail = PartList(4, i)
        
        For j = 0 To 5
            BestList(j, k) = PartList(j, i)
        Next
            
        k = k + 1
    Else
        If Not CStr(PartList(4, i)) = CStr(sOldEmail) Then 
            sOldEmail = PartList(4, i)
        
            For j = 0 To 5
                BestList(j, k) = PartList(j, i)
            Next
                
            k = k + 1
        End If
    End If
Next

Private Function EventDate(lRaceID)
    EventDate = "1/1/1900"

    If Not CStr(lRaceID) = vbNullString Then
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT e.EventDate FROM Events e INNER JOIN RaceData rd ON e.EventID = rd.EventID WHERE rd.RaceID = " & lRaceID
        rs.Open sql, conn, 1, 2
        EventDate = rs(0).Value
        rs.Close
        Set rs = Nothing
    End If
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../includes/meta2.asp" -->

<title>GSE&copy; Admin Data Modify</title>

<!--#include file = "../includes/js.asp" -->

</head>

<body>
<div class="container">
  	<!--#include file = "../includes/header.asp" -->

	<div id="row">
		<!--#include file = "../includes/admin_menu.asp" -->
		<div class="col-md-10">
		    <h4 class="h4">GSE&copy; Email Extract</h4>

			<form name="data_mod" method="post" action="data_modify.asp" style="font-size:0.85em;">
				<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
				<input type="submit" name="submit" id="submit" value="Submit This">
			</form>

            <%For i = 0 To UBound(BestList, 2)%>
                <%If BestList(0, i) = vbNullString Then%>
                    <%Exit For%>
                <%Else%>
                    <%For j = 0 To 4%>
                        <%=BestList(j, i) & vbTab%>
                    <%Next%>
                    <%=EventDate(BestList(5, i))%>
                    <br>
                <%End If%>
            <%Next%>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
