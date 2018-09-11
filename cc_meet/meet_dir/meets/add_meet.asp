<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lMeetID
Dim sMeetName, dMeetDate, sMeetSite, sMeetHost, sWebSite, sComments, sEntryFee, sSport, sMeetDirName, sRslts
Dim cdoMessage, cdoConfig
Dim sMsg

If Not Session("role") = "meet_dir" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
											
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

If Request.Form.Item("submit_meet") = "submit_meet" Then
	sMeetName = Request.Form.Item("meet_name")
	dMeetDate = Request.Form.Item("meet_month") & "/" & Request.Form.Item("meet_day") & "/" & Request.Form.Item("meet_year") 
	sMeetSite = Request.Form.Item("meet_site")
	sSport = Request.Form.Item("meet_type")
	sMeetHost = Request.Form.Item("meet_host")
	sWebSite = Request.Form.Item("web_site")
	sComments = Request.Form.Item("comments")
	sEntryFee = Request.Form.Item("entry_fee")
	
	If sWebSite = vbNullString Then sWebSite = "www.gopherstateevents.com"
	
	If IsDate(dMeetDate) = True Then
		sql = "INSERT INTO Meets (MeetDirID, MeetName, MeetDate, MeetSite, MeetHost, WebSite, Comments, DateReg, "
		sql = sql & "EntryFee, WhenShutdown, Sport) VALUES (" & Session("my_id") & ", '" & Replace(sMeetName, "'", "''") & "', '" 
		sql = sql & dMeetDate & "', '" & Replace(sMeetSite, "'", "''") & "', '" & Replace(sMeetHost, "'", "''") & "', '" & sWebSite & "', '"
		sql = sql & Replace(sComments, "'", "''") &"', '" & Now() & "', '" & sEntryFee & "', '" & CDate(dMeetDate) - 1 
		sql = sql & " 4:00:00 PM', '" & sSport & "')"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
        
        'get meet id
        sql = "SELECT MeetsID FROM Meets WHERE MeetDirID = " & Session("my_id") & " AND MeetName = '" 
        sql = sql & Replace(sMeetName, "'", "''") & "' AND MeetDate = '" & dMeetDate & "'"
        Set rs = conn.Execute(sql)
        lMeetID = rs(0).Value
        Set rs = Nothing
        
        'insert bib range value
        sql = "INSERT INTO BibRange (MeetsID) VALUES (" & lMeetID & ")"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
		
		sql = "SELECT FirstName, LastName FROM MeetDir WHERE MeetDirID = " & Session("my_id")
		Set rs = conn.Execute(sql)
		sMeetDirName = rs(0).Value & " " & rs(1).Value
		Set rs = Nothing
		
		sMsg = vbCrLf
		sMsg = sMsg & "A new cross-country meet has just been added by " & sMeetDirName & vbCrLf & vbCrLf
		
		sMsg = sMsg & "DETAILS:" & vbCrLf
		sMsg = sMsg & "Meet Name: " & sMeetName & vbCrLf
		sMsg = sMsg & "Meet Date: " & dMeetDate & vbCrLf
		sMsg = sMsg & "Meet Site: " & sMeetSite & vbCrLf
		sMsg = sMsg & "Meet Host: " & sMeetHost & vbCrLf
		sMsg = sMsg & "Entry Fee: " & sEntryFee & vbCrLf
		sMsg = sMsg & "Web Site: " & sWebSite & vbCrLf
		sMsg = sMsg & "Comments: " & sComments & vbCrLf
		
%>
<!--#include file = "../../../includes/cdo_connect.asp" -->
<%

		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = "bob.schneider@gopherstateevents.com"
			.From = "bob.schneider@gopherstateevents.com"
			.Subject = "GSE New CCMeet"
			.TextBody = sMsg
			.Send
		End With
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing
		
		sRslts = "This meet was successfully entered."
		
		sMeetName = vbNullString
		sMeetSite = vbNullString
		sMeetHost = vbNullString
		sEntryFee = vbNullString
		sWebSite = vbNullString
		sComments = vbNullString
	Else
		sRslts = "The date you supplied is not a valid date.  Please re-enter a date and re-submit the data."
	End If
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../../includes/meta2.asp" -->
<title>GSE Upload Meet Info</title>
<!--#include file = "../../../includes/js.asp" -->

<script>
function chkFields(){
	if (document.add_meet.meet_name.value==''){
		alert('You must supply a meet name!');
		return false;
	}
	else
		if (document.add_meet.meet_day.value==''){
			alert('You must supply a day for this meet!');
			return false;
		}
	else
		if (document.add_meet.meet_month.value==''){
			alert('You must supply a month for this meet!');
			return false;
		}
	else
		if (document.add_meet.meet_year.value==''){
			alert('You must supply a year for this meet!');
			return false;
		}
	else
		if (document.add_meet.sport.value==''){
			alert('You must supply a sport for this meet!');
			return false;
		}
	else
		return true;
}
</script>
</head>
<body>
<div class="container">
	<!--#include file = "../../../includes/header.asp" -->
	
	<div id="row">
		<!--#include file = "../../../includes/meet_dir_menu.asp" -->
		<div class="col-md-10">
			<h4 class="h4">CCMeet Director: Add Meet</h4>

			<form class="form-horizontal" name="add_meet" method="post" action="add_meet.asp" onsubmit="return chkFields()">
				<p><%=sRslts%></p>
				<div class="form-group">
					<label for="meet_name" class="control-label col-xs-4">Meet Name:</label>
					<td>
                        <input type="text" name="meet_name" id="meet_name" maxlength="25" value="<%=sMeetName%>">
                    </td>
					<th>Meet Date:</th>
					<td style="width:25%;white-space:nowrap;">
						<select name="meet_month" id="meet_month">
							<option value="">&nbsp;</option>
							<%For i = 1 To 12%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
						/
						<select name="meet_day" id="meet_day">
							<option value="">&nbsp;</option>
							<%For i = 1 To 31%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
						/
						<select name="meet_year" id="meet_year">
							<%For i = 2010 To Year(Date) + 1%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
					</td>
				</tr>
				<tr>
					<th>Meet Host:</th>
					<td><input type="text" name="meet_host"id="meet_host" maxlength="50" value="<%=sMeetHost%>"></td>
					<th>Web Site:</th>
					<td><input type="text" name="web_site" id="web_site" maxlength="50" value="<%=sWebSite%>" onkeyup="chkStr(this)"></td>
				</tr>
				<tr>
					<th>Entry Fee:</th>
					<td>$<input type="text" name="entry_fee" id="entry_fee" maxlength="15" size="5" value="<%=sEntryFee%>"></td>
					<th>Sport:</th>
					<td>
						<select name="sport" id="sport">
							<option value="">&nbsp;</option>
							<%If sSport="Cross-Country" Then%>
								<option value="Cross-Country" selected>Cross-Country</option>
								<option value="Nordic Ski">Nordic Ski</option>
							<%ElseIf sSport = "Nordic Ski" Then%>
								<option value="Cross-Country">Cross-Country</option>
								<option value="Nordic Ski" selected>Nordic Ski</option>
							<%Else%>
								<option value="Cross-Country">Cross-Country</option>
								<option value="Nordic Ski">Nordic Ski</option>
							<%End If%>
						</select>
					</td>
				</tr>
				<tr>
					<th valign="top">Meet Site:</th>
					<td><textarea name="meet_site" id="meet_site" rows="2" cols="35"><%=sMeetSite%></textarea></td>
					<th valign="top">Comments:</th>
					<td><textarea name="comments" id="comments" rows="2" cols="35"><%=sComments%></textarea></td>
				</tr>
				<tr>
					<td style="text-align:center;" colspan="4">
						<input type="hidden" name="submit_meet" id="submit_meet" value="submit_meet">
						<input type="submit" name="submit" id="submit" tabindex="9" value="Submit This Meet">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
<%
conn.close
Set conn = Nothing
%>
</body>
</html>
