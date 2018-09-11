<%@ Language=VBScript%>

<%
Option Explicit

Dim sql, rs, conn
Dim i, j, k
Dim sTeamIDs
Dim AllEmails(), GoodEmails(), BadEmails()

If Not Session("role") = "coach" Then  
    If Not Session("role") = "team_staff" Then Response.Redirect "/default.asp?sign_out=y"
End If

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionTimeout = 30
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get team staff
Set rs = Server.CreateObject("ADODB.Recordset")
If Session("role") = "coach" Then
    sql = "SELECT TeamsID FROM Teams WHERE CoachesID = " & Session("my_id")
Else
    sql = "SELECT TeamsID FROM Teams WHERE CoachesID = " & Session("team_coach_id")
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
    sTeamIDs = sTeamIDs & rs(0).Value & ", "
	rs.MoveNext
Loop
rs.Close
Set rs=Nothing

sTeamIDs = Left(sTeamIDs, Len(sTeamIDs) - 2)

'get roster names for the email
i = 0	
ReDim AllEmails(1, 0)
Set rs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT Email, FirstName, LastName FROM Roster WHERE TeamsID IN (" & sTeamIDs  & ") AND Archive = 'n' ORDER BY LastName, FirstName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If Not rs(0).Value & "" = "" Then
		AllEmails(0, i) = rs(0).Value
		AllEmails(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
		i = i + 1
		ReDim Preserve AllEmails(1, i)
	End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get staff names for the email
Set rs=Server.CreateObject("ADODB.Recordset")
If Session("role") = "coach" Then
    sql = "SELECT Email, FirstName, LastName FROM TeamStaff WHERE CoachesID = " & Session("my_id")  & " ORDER BY LastName, FirstName"
Else
    sql = "SELECT Email, FirstName, LastName FROM TeamStaff WHERE CoachesID = " & Session("team_coach_id")  & " ORDER BY LastName, FirstName"
End If
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If Not rs(0).Value & "" = "" Then
		AllEmails(0, i) = rs(0).Value
		AllEmails(1, i) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'")
		i = i + 1
		ReDim Preserve AllEmails(1, i)
	End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

'get team contact names for the email
Set rs=Server.CreateObject("ADODB.Recordset")
sql = "SELECT Email, ContactName FROM TeamContacts WHERE TeamsID IN (" & sTeamIDs  & ") ORDER BY ContactName"
rs.Open sql, conn, 1, 2
Do While Not rs.EOF
	If Not rs(0).Value & "" = "" Then
		AllEmails(0, i) = rs(0).Value
		AllEmails(1, i) = Replace(rs(1).Value, "''", "'")
		i = i + 1
		ReDim Preserve AllEmails(1, i)
	End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing

j = 0
k = 0
ReDim BadEmails(1, 0)
ReDim GoodEmails(1, 0)
For i = 0 to UBound(AllEmails, 2)  - 1
	AllEmails(0, i) = Trim(AllEmails(0, i))
	
	If ValidEmail(AllEmails(0, i)) = True Then
		GoodEmails(0, k) = AllEmails(0, i)
		GoodEmails(1, k) = AllEmails(1, i)
		k = k + 1
		ReDim Preserve GoodEmails(1, k)
	Else
		BadEmails(0, j) = AllEmails(0, i)
		BadEmails(1, j) = AllEmails(1, i)
		j = j + 1
		ReDim Preserve BadEmails(1, j)
	End If
Next

%>
<!--#include file = "../../../../includes/valid_email.asp" -->

<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../../includes/meta2.asp" -->
<title>Gopher State Events&reg; Email Validation</title>
</head>

<body>
<div class="container">
	<!--#include file = "../../../../includes/header.asp" -->
    
	<div class="row">
		<div class="col-sm-2">
			<!--#include file = "../../../../includes/coach_menu.asp" -->
		</div>
		<div class="col-sm-10">
			<!--#include file = "communications_nav.asp" -->
			<h4 class="h4">Gopher State Events<sup>&reg;</sup> Email Validation Report</h4>
			<br>
			<div class="row">
				<div class="col-sm-6">
					<h5 class="h5">Valid email addresses:</h5>
					<ul class="list-group">
						<%For i = 0 to UBound(GoodEmails, 2) - 1%>
							<li class="list-group-item"><%=GoodEmails(1, i)%>&nbsp;(<%=GoodEmails(0, i)%>)</li>
						<%Next%>
					</ul>
				</div>
				<div class="col-sm-6">
					<h5 class="h5">Invalid email addresses:</h5>
					<%If UBound(BadEmails, 2) > 0 Then%>
						<ul class="list-group">
							<%For i = 0 to UBound(BadEmails, 2) - 1%>
							<li class="list-group-item"><%=BadEmails(1, i)%>&nbsp;(<%=BadEmails(0, i)%>)</li>
							<%Next%>
						</ul>
					<%Else%>
						<p>All of your email addresses are valid.</p>
					<%End If%>
				</div>
			</div>
		</div>
	</div>
 </div>
 <!--#include file = "../../../../includes/footer.asp" --> 
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
