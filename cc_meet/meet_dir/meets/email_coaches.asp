<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i, j
Dim lThisMeet
Dim Coaches(), Meets(), MeetTeams(), EmailArr(), Recips()
Dim cdoMessage, cdoConfig
Dim sMsg, sSubject

If Not Session("role") = "admin" Then  
    If Not Session("role") = "meet_dir" Then Response.Redirect "/default.asp?sign_out=y"
End If

lThisMeet = Request.QueryString("this_meet")

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
									
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

ReDim MeetTeams(0)
ReDim Coaches(3, 0)

If Request.Form.Item("submit_this") = "submit_this" Then
	sSubject = Request.Form.Item("subject")
	sMsg = Request.Form.Item("msg") & vbCrLf & vbCrLf

    Call GetMeetTeams()

	i = 0
	ReDim EmailArr(1, 0)
    ReDim Recips(1, 0)

	If Request.Form.Item("all") = "on" Then
		For j = 0 To UBound(Coaches, 2) - 1
            EmailArr(0, i) = Coaches(1, j)
            EmailArr(1, i) = Coaches(3, j)
            i = i + 1
            ReDim Preserve EmailArr(1, i)
        Next
	Else	
		For j = 0 To UBound(Coaches, 2) - 1
            If Request.Form.Item("coach_" & Coaches(0, j)) = "on" Then
                EmailArr(0, i) = Coaches(1, j)
                EmailArr(1, i) = Coaches(3, j)
                i = i + 1
                ReDim Preserve EmailArr(1, i)
            End If
        Next
	End If

    'manage duplicates
    j = 0
    If Request.Form.Item("no_duplicates") = "on" Then
        For i = 0 To UBound(EmailArr, 2) - 1
            If IsDuplicate(EmailArr(1, i)) = "n" Then
                Recips(0, j) = EmailArr(0, i)
                Recips(1, j) = EmailArr(1, i)
                j = j + 1
                ReDim Preserve Recips(1, j)
            End If
        Next
    Else
        For i = 0 To UBound(EmailArr, 2) - 1
            Recips(0, j) = EmailArr(0, i)
            Recips(1, j) = EmailArr(1, i)
            j = j + 1
            ReDim Preserve Recips(1, j)
        Next
    End If

%>
<!--#include file = "../../../includes/cdo_connect.asp" -->
<%
	
	For i = 0 to UBound(Recips, 2) - 1
		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = Recips(1, i)
'			.To = "bob.schneider@gopherstateevents.com"
			.From = Session("my_email")
            If i = 0 Then
                .BCC = "bob.schneider@gopherstateevents.com;" & Session("my_email")
            End If
			.Subject = sSubject
			.TextBody = sMsg
			.Send
		End With
		Set cdoMessage = Nothing
	Next
	
	Set cdoConfig = Nothing
ElseIf Request.Form.Item("submit_meet") = "submit_meet" Then
    lThisMeet = Request.Form.Item("meets")
End If

If lThisMeet & "" = "" Then lThisMeet = 0

i = 0
ReDim Meets(1, 0)
sql = "SELECT MeetsID,  MeetName, MeetDate FROM Meets WHERE MeetDirID = " & Session("my_id") & " ORDER BY MeetDate DESC"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	Meets(0, i) = rs(0).Value
	Meets(1, i) = Replace(rs(1).Value, "''", "'") & " (" & Year(rs(2).Value) & ")"
	i = i + 1
	ReDim Preserve Meets(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If UBound(Meets, 2) = 1 Then lThisMeet = Meets(0, 0)

If CLng(lThisMeet) > 0 Then
    Call GetMeetTeams()
End If

Private Sub GetMeetTeams()
    Dim x, y

	x = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT t.TeamsID FROM Teams t INNER JOIN MeetTeams mt On mt.TeamsID = t.TeamsID "
	sql = sql & "WHERE mt.MeetsID = " & lThisMeet & " ORDER BY t.TeamName"
	rs.Open sql, conn, 1, 2
	Do While Not rs.EOF
		MeetTeams(x) = rs(0).Value
		x = x + 1
		ReDim Preserve MeetTeams(x)
		rs.MoveNext
	Loop
    rs.Close
	Set rs = Nothing

    y = 0
    For x = 0 To UBound(MeetTeams) - 1
        Set rs = Server.CreateObject("ADODB.Recordset")
        sql = "SELECT c.CoachesID, c.FirstName, c.LastName, t.TeamName, t.Gender, c.Email FROM Coaches c INNER JOIN Teams t on c.CoachesID = t.CoachesID "
        sql = sql & "WHERE t.TeamsID = " & MeetTeams(x)
        rs.Open sql, conn, 1, 2
        Coaches(0, y) = rs(0).Value
        Coaches(1, y) = Replace(rs(1).Value, "''", "'") & " " & Replace(rs(2).Value, "''", "'") 
        Coaches(2, y) = "-" & Replace(rs(3).Value, "''", "'") & " (" & rs(4).Value & ")"
        Coaches(3, y) = rs(5).Value
        y = y + 1
        ReDim Preserve Coaches(3, y)
        rs.Close
        Set rs = Nothing
    Next
End Sub

Private Function IsDuplicate(sThisEmail)
    Dim x

    IsDuplicate = "n"
    For x = 0 To UBound(Recips, 2) - 1
        If CStr(Recips(1, x)) = CStr(sThisEmail) Then
            IsDuplicate = "y"
            Exit For
        End If
    Next
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Email Cross-Country/Nordic Coaches</title>
<!--#include file = "../../../includes/js.asp" -->

</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../../includes/meet_dir_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">Email Cross-Country/Nordic Coaches</h4>

            <%If UBound(Meets, 2) > 1 Then%>
                <form class="form-inline bg-success" name="my_meets" method="post" action="email_coaches.asp">
                <label for="meets">Select Meet</label>
                <select class="form-control" name="meets" id="meets" onchange="this.form.submit1.click();">
                    <option value=""></option>
					<%For i = 0 to UBound(Meets, 2) - 1%>
                        <%If CLng(lThisMeet) = CLng(Meets(0, i)) Then%>
							<option value="<%=Meets(0, i)%>" selected><%=Meets(1, i)%></option>
                        <%Else%>
							<option value="<%=Meets(0, i)%>"><%=Meets(1, i)%></option>
                        <%End If%>						
                    <%Next%>
                </select>
				<input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
				<input type="submit" class="form-control" name="submit1" id="submit1" value="Send Email">
                </form>
            <%End If%>

            <%If CLng(lThisMeet) > 0 Then%>
			    <form class="form-horizontal" name="send_email" method="post" action="email_coaches.asp?this_meet=<%=lThisMeet%>">
				<div class="col-sm-6">
					<label>Select Recipients:</label><br>
                    <ol class="form-group">
                        <li class="form-group-item"><input type="checkbox" name="all" id="all">&nbsp;<span style="font-weight: bold;">All</span></li>
						<%For i = 0 to UBound(Coaches, 2) - 1%>
							<li class="form-group-item"><input type="checkbox" name="coach_<%=Coaches(0, i)%>" id="coach_<%=Coaches(0, i)%>">&nbsp;<%=Coaches(1, i)%><%=Coaches(2, i)%></li>
						<%Next%>
                    </ol>
				</div>
				<div class="col-sm-6 bg-warning">
                    <br>
					<div class="form-group">
						<label for="subject" class="control-label col-xs-4">Subject:</label>
						<div class="col-xs-8">
                            <input type ="text" class="form-control" name="subject" id="subject">
                        </div>
					</div>
					<div class="form-group" style="text-align: center;">
                        <input type="checkbox" name="no_duplicates" id="no_duplicates">&nbsp;No Duplicates
					</div>
					<div class="form-group">
						<label for="msg" class="control-label col-xs-4">Message:</label>
						<div class="col-xs-8">
                            <textarea class="form-control" name="msg" id="msg" rows="10"></textarea>
                        </div>
					</div>
					<div class="form-group">
						<input type="hidden" name="submit_this" id="submit_this" value="submit_this">
						<input type="submit" class="form-control" name="submit2" id="submit2" value="Send Email">
					</div>
                </div>
			    </form>
            <%End If%>
		</div>
	</div>
</div>
<%
conn.Close
Set conn = Nothing
%>
</body>
</html>
