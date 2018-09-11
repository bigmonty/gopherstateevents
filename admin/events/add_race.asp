<%@ Language=VBScript %>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lThisType, lRaceType, lRaceID, lEventID
Dim EventTypes()
Dim sDist, sEntryFeePre, sEntryFee, sStartTime, sCertified, sRaceName, sRaceDist, sEventName, sStartType
Dim sSmall, sMedium, sLarge, sXLarge, sXXLarge, sShort, sLong, sChooseNone, sMsg
Dim iMAwds, iFAwds, iNumAwds, iEndAge
Dim dEventDate
Dim cdoMessage, cdoConfig

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")


Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
	
'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
rs.Close
Set rs = Nothing

'get event types
i = 0
ReDim EventTypes(1, 0)
sql = "SELECT EvntRaceTypesID, EvntRaceType FROM EvntRaceTypes ORDER BY EvntRaceType"
Set rs = conn.Execute(sql)
Do While Not rs.EOF
	EventTypes(0, i) = rs(0).Value
	EventTypes(1, i) = rs(1).Value
	i = i + 1
	ReDim Preserve EventTypes(1, i)
	rs.MoveNext
Loop
Set rs = Nothing

If Request.Form.Item("submit_race") = "submit_race" Then
	'get variables from add_race
	sRaceName = Replace(Request.Form.Item("race_name"), "'", "''")
	sDist = Request.Form.Item("race_dist")
	lRaceType = Request.Form.Item("race_type")
	sEntryFeePre = Request.Form.Item("pre_fee")
	sEntryFee = Request.Form.Item("fee")
	sStartTime = Request.Form.Item("hours") & ":" & Request.Form.Item("minutes") & Request.Form.Item("am_pm")
	sCertified = Request.Form.Item("certified")
	iMAwds = Request.Form.Item("m_awds")
	iFAwds = Request.Form.Item("f_awds")
	sStartType = Request.Form.Item("start_type")

	'insert into racedata
	sql = "INSERT INTO RaceData (EventID, Dist, Type, EntryFeePre, EntryFee, StartTime, Certified, RaceName, "
	sql = sql & "StartType, MAwds, FAwds) VALUES (" & lEventID & ", '" & sDist & "', '" & lRaceType & "', " & sEntryFeePre 
	sql = sql & ", " & sEntryFee & ", '" & sStartTime & "', '" & sCertified & "', '" & sRaceName & "', '" & sStartType & "', " 
	sql = sql & iMAwds & ", " & iFAwds & ")"
	Set rs = conn.Execute(sql)
	Set rs = Nothing
	
	'get race id
	sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID & " AND Dist = '" & sDist & "' AND Type = '"
	sql = sql & lRaceType & "'"
	Set rs = conn.Execute(sql)
	lRaceID = rs(0).Value
	Set rs = Nothing
	
	'write to tshirt tables if necessary
	If Request.Form.Item("has_shrts") = "y" Then
		If Request.Form.Item("small") ="on" Then
			sSmall = "y"
		Else
			sSmall = "n"
		End If
		
		If Request.Form.Item("medium") ="on" Then
			sMedium = "y"
		Else
			sMedium = "n"
		End If
		
		If Request.Form.Item("large") ="on" Then
			sLarge = "y"
		Else
			sLarge = "n"
		End If
		
		If Request.Form.Item("xlarge") ="on" Then
			sXLarge = "y"
		Else
			sXLarge = "n"
		End If
		
		If Request.Form.Item("xxlarge") ="on" Then
			sXXLarge = "y"
		Else
			sXXLarge = "n"
		End If
		
		If Request.Form.Item("short_sleeve") ="on" Then
			sShort = "y"
		Else
			sShort = "n"
		End If
		
		If Request.Form.Item("long_sleeve") ="on" Then
			sLong = "y"
		Else
			sLong = "n"
		End If
		
		If Request.Form.Item("choose_none") ="on" Then
			sChooseNone = "y"
		Else
			sChooseNone = "n"
		End If
		
		sql = "INSERT INTO TShirtData (RaceID, IsOption, Small, Medium, Large, XLarge, XXLarge, Short, Long, ChooseNone) "
		sql = sql & "VALUES (" & lRaceID & ", 'y', '" & sSmall & "', '" & sMedium & "', '" & sLarge & "', '" & sXLarge & "', '"
		sql = sql & sXXLarge & "', '" & sShort & "', '" & sLong & "', '" & sChooseNone & "')"
	Else
		sql = "INSERT INTO TShirtData (RaceID, IsOption) VALUES (" & lRaceID & ", 'n')"
	End If
	Set rs = conn.Execute(sql)
	Set rs =Nothing

	For i = 0 to 14
		'insert mens age groups/awards into agegroup table
		If Not Request.Form.Item("magegrp" & i) & "" = "" Then
            If Not Request.Form.Item("magegrp" & i) = "0" Then
			    If Request.Form.Item("mawds" & i) = vbNullString Then
				    iNumAwds = 0
			    Else
				    iNumAwds = Request.Form.Item("mawds" & i)
			    End If
			
			    sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'm', "
			    sql = sql & Request.Form.Item("magegrp" & i) & ", " & iNumAwds & ")"
			    Set rs = conn.Execute(sql)
			    Set rs = Nothing
            End If
		End if
		
		'insert womens age groups/awards into agegroup table
		If Not Request.Form.Item("fagegrp" & i) & "" = "" Then
            If Not Request.Form.Item("fagegrp" & i) = "0" Then
			    If Request.Form.Item("fawds" & i) = vbNullString Then
				    iNumAwds = 0
			    Else
				    iNumAwds = Request.Form.Item("fawds" & i)
			    End If
			
			    sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'f', "
			    sql = sql & Request.Form.Item("fagegrp" & i) & ", " & iNumAwds & ")"
			    Set rs = conn.Execute(sql)
			    Set rs = Nothing
            End If
		End If
	Next

    'check for last end age = 110
    iEndAge = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EndAge FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'm' ORDER BY EndAge DESC"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        iEndAge = rs(0).Value
    End If
    rs.Close
    Set rs = Nothing

    If CInt(iEndAge) < 110 Then
   		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'm', 110, 0)"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    End If

    iEndAge = 0
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT EndAge FROM AgeGroups WHERE RaceID = " & lRaceID & " AND Gender = 'f' ORDER BY EndAge DESC"
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        iEndAge = rs(0).Value
    End If
    rs.Close
    Set rs = Nothing

    If CInt(iEndAge) < 110 Then
   		sql = "INSERT INTO AgeGroups (RaceID, Gender, EndAge, NumAwds) VALUES (" & lRaceID & ", 'f', 110, 0)"
		Set rs = conn.Execute(sql)
		Set rs = Nothing
    End If
	
	sMsg = vbCrLf & "This is notification that a new race has been added to the following event:" & vbCrLf & vbCrLf
	sMsg = sMsg & "Event Name: " & sEventName & vbCrLf & vbCrLf
	sMsg = sMsg & "Event Date: " & dEventDate & vbCrLf & vbCrLf
	sMsg = sMsg & "Race Distance: " & sDist & vbCrLf & vbCrLf
	sMsg = sMsg & "Race Type: " &GetThisType(lRaceType) & vbCrLf & vbCrLf

%>
<!--#include file = "../../includes/cdo_connect.asp" -->
<%
	
	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
		Set .Configuration = cdoConfig
		.To = "bob.schneider@gopherstateevents.comstateevents.com"
		.From = "bob.schneider@gopherstateevents.com"
	    .Subject = "GSE New Race Registration"
		.TextBody = sMsg
		.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
	
	Response.Redirect "race_info.asp?event_id=" & lEventID
End If

Private Function GetThisType(lEventType)
	sql = "SELECT EvntRaceType FROM EvntRaceTypes WHERE EvntRaceTypesID = " & lEventType
	Set rs = conn.Execute(sql)
	GetThisType = rs(0).Value
	Set rs = Nothing
End Function
%>
<!DOCTYPE html>
<html lang="en">
<head>
<!--#include file = "../../includes/meta2.asp" -->
<title>Add a Race to <%=sEventName%></title>
<script>
<!--
function chkFlds(){
 	if (document.race_info.race_dist.value == '' || 
 	    document.race_info.race_name.value == '' ||
 	    document.race_info.race_type.value == '' ||
	 	document.race_info.hours.value == '' || 
	 	document.race_info.minutes.value == '' || 
	 	document.race_info.am_pm.value == '' ||
	 	document.race_info.fee.value == '' ||
	 	document.race_info.pre_fee.value == '' ||
	 	document.race_info.certified.value == '')
		{
  		alert('Please fill in all required fields');
  		return false
  		}
 	else
		if (isNaN(document.race_info.race_dist.value) ||
		   isNaN(document.race_info.fee.value) ||
		   isNaN(document.race_info.pre_fee.value))
    		{
			alert('The race distance, entry fee and pre-reg entry fee must be numeric values!');
			return false
			} 	
  	else
  		if (document.race_info.choose_none.checked != true &&
	 	   document.race_info.short_sleeve.checked != true &&
	 	   document.race_info.long_sleeve.checked != true)
	 	   {
	 	   alert('Please select a shirt style for this race!');
	 	   return false
	 	   }
  	else
  		if (document.race_info.choose_none.checked != true &&
	 	   (document.race_info.small.checked != true &&
			document.race_info.medium.checked != true &&	 	   
			document.race_info.large.checked !=  true && 	   
			document.race_info.xlarge.checked !=  true &&	 	   
			document.race_info.xxlarge.checked !=  true))
	 	   {
	 	   alert('Please select shirt sizes for this race!');
	 	   return false
	 	   }
	else
   		return true
}

    var OK = true;
    for(var i=0; i<=14; i++){
        if (isNaN(document.race_info['magegrp'+i].value) ||
            isNaN(document.race_info['mawds'+i].value)   ||
            isNaN(document.race_info['fagegrp'+i].value) ||
            isNaN(document.race_info['fawds'+i].value)) {
            OK = false;
            break; 
        }
    }
    if (!OK) alert('All data entered on this page must be numeric!');
    return OK;
}
-->
</script>
</head>

<body>
<div class="container">
  	<!--#include file = "../../includes/header.asp" -->

	<div class="row">
		<!--#include file = "../../includes/admin_menu.asp" -->
		<div class="col-sm-10">
		    <h4 class="h4"><%=sEventName%>: Add Race (<span style="font-weight:normal">* = required fields</span></h4>
			
			<!--#include file = "../../includes/event_nav.asp" -->

			<div>
				<a href="/admin/race_info.asp?event_id=<%=lEventID%>" rel="nofollow">Race Info</a>
            </div>

			<a href="clone_race.asp?event_id=<%=lEventID%>" style="margin-left:10px;font-size:0.85em;">Clone Existing Race</a>
			<form name="race_info" method="post" action="add_race.asp?event_id=<%=lEventID%>" onsubmit="return chkFlds()">
			<table style="margin-left:10px;">
				<tr>
					<th><span style="color:#d62002">*</span>Race Name:</th>
					<td><input name="race_name" id="race_name" maxlength="35"></td>
					<th><span style="color:#d62002">*</span>Distance:</th>
					<td><input name="race_dist" id="race_dist" size="2" maxLength="8"></td>
				</tr>
				<tr>
					<th><span style="color:#d62002">*</span>Type:</th>
					<td>
						<select name="race_type" id="race_type">
							<option value="">&nbsp;</option>
							<%For i = 0 to UBound(EventTypes, 2) - 1%>
								<option value="<%=EventTypes(0, i)%>"><%=EventTypes(1, i)%></option>
							<%Next%>
						</select>
					</td>
					<th><span style="color:#d62002">*</span>Certified?</th>
					<td>
						<select name="certified" id="certified"> 
							<option value="">&nbsp;</option>
							<option value="y">Yes</option>
							<option value="n">No</option>
						</select>
					</td>
				</tr>
				<tr>
					<th><span style="color:#d62002">*</span>Pre-Reg Fee: $</th>
					<td><input name="pre_fee" id="pre_fee" maxLength="5" size="3"></td>
					<th><span style="color:#d62002">*</span>Race Day Fee: $</th>
					<td><input name="fee" id="fee" maxLength="5" size="3"></td>
				</tr>
				<tr>
					<th><span style="color:#d62002">*</span>Start Time:</th>
					<td>
						<select name="hours" id="hours">
							<option value="">&nbsp;</option>
							<script>
							<!--
							for (i = 1; i <= 12; i++)
							{
							document.write("<option value='"+i+"'>"+i+"</option>")
							document.write("<br>")
							}
							-->
							</script>
						</select> : 
						<select name="minutes" id="minutes">
							<option value="">&nbsp;</option>
							<option value="00">00</option>
							<script>
							<!--
							for (i = 5; i <=55; i = i+5)
							{
								if(i==5){
								document.write("<option value='0"+i+"'>0"+i+"</option>")
								document.write("<br>")
								}
								else{
								document.write("<option value='"+i+"'>"+i+"</option>")
								document.write("<br>")
								}
							}
							-->
							</script>
						</select>
						<select name="am_pm" id="am_pm">
							<option value="">&nbsp;</option>
							<option value="am">am</option>
							<option value="pm">pm</option>
						</select>
					</td>
					<th>Start Type:</th>
					<td>
						<select name="start_type" id="start_type">
							<option value="mass">Mass</option>
							<option value="wave">Wave</option>
							<option value="interval">Interval</option>
						</select>
					</td>
				</tr>
				<tr>
					<th><span style="color:#d62002">*</span>Open Awards:</th>
					<td>
						Male:&nbsp; 
						<select name="m_awds" id="m_awds"> 
							<%For i = 0 To 100%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
						Female:&nbsp; 
						<select name="f_awds" id="f_awds"> 
							<%For i = 0 To 100%>
								<option value="<%=i%>"><%=i%></option>
							<%Next%>
						</select>
					</td>
					<th>T-Shirts?</th>
					<td>
						<select name="has_shrts" id="has_shrts"> 
							<option value="y">Yes</option>
							<option value="n">No</option>
						</select>
					</td>
				</tr>
				<tr>
					<th>Styles:</th>
					<td>
						<input type="checkbox" name="choose_none" id="choose_none">Choose None&nbsp;&nbsp;
						<input type="checkbox" name="short_sleeve" id="short_sleeve">Short&nbsp;&nbsp;
						<input type="checkbox" name="long_sleeve" id="long_sleeve">Long&nbsp;&nbsp;
					</td>
					<th>Sizes:</th>
					<td>
						<input type="checkbox" name="small" id="small">S&nbsp;&nbsp;
						<input type="checkbox" name="medium" id="medium">M&nbsp;&nbsp;
						<input type="checkbox" name="large" id="large">L&nbsp;&nbsp;
						<input type="checkbox" name="xlarge" id="xlarge">XL&nbsp;&nbsp;
						<input type="checkbox" name="xxlarge" id="xxlarge">XXL&nbsp;&nbsp;
					</td>
				</tr>
				<tr>
					<td style="padding-left:100px;width:50%;" colspan="2">
						<table style="font-size:1.0em;">
							<tr><th style="text-align:center;" colspan="2">Men's Age Groups/Awards</th></tr>
							<tr>
								<th>End Age</th>
								<th style="text-align:center">Awds</th>
							</tr>
							<script>
							<!--
								for (i = 0; i <=14; i++)
								{
								document.write("<tr>")
								document.write("<th>")
								document.write("Age Grp "+i+":&nbsp;")
								document.write("<input name='magegrp"+i+"' id='magegrp"+i+"' style='text-align:center' size='1' maxlength='2'>")
								document.write("</th>")
								document.write("<td style='text-align:center'>")
								document.write("<input name='mawds"+i+"' id='mawds"+i+"' style='text-align:center' size='1' maxlength='2'>")
								document.write("</td>")
								document.write("</tr>")
								}
							-->
							</script>
						</table>
					</td>
					<td style="text-align:center;width:50%;" colspan="2">
						<table style="font-size:1.0em;">
							<tr><th style="text-align:center;" colspan="2">Women's Age Groups/Awards</th></tr>
							<tr>
								<th>End Age</th>
								<th style="text-align:center">Awds</th>
							</tr>
							<script>
								for (i = 0; i <=14; i++)
								{
								document.write("<tr>")
								document.write("<td style='text-align:right'>")
								document.write("Age Grp "+i+":&nbsp;")
								document.write("<input name='fagegrp"+i+"' id='fagegrp"+i+"' style='text-align:center' size='1' maxlength='2'>")
								document.write("</td>")
								document.write("<td style='text-align:center'>")
								document.write("<input name='fawds"+i+"' id='fawds"+i+"' style='text-align:center' size='1' maxlength='2'>")
								document.write("</td>")
								document.write("</tr>")
								}
							</script>
						</table>
					</td>
				</tr>
				<tr> 
					<td style="text-align:center;" colspan="4">
						<input type="hidden" name="submit_race" id="submit_race" value="submit_race">
						<input type="submit" name="submit1" id="submit1" value="Add Race">
					</td>
				</tr>
			</table>
			</form>
		</div>
	</div>
</div>
<!--#include file = "../../includes/footer.asp" -->
</body>
<%
conn.Close
Set conn = Nothing
%>
</html>
