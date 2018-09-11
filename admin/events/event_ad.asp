<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim i
Dim lEventID
Dim sEventName, sAdExists, sLogo
Dim lngNumViews
Dim dEventDate

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

lEventID = Request.QueryString("event_id")
Session("event_id") = lEventID

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
							
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"

'get event information
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT EventName, EventDate, Logo FROM Events WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
sEventName = rs(0).Value
dEventDate = rs(1).Value
sLogo = rs(2).Value
rs.Close
Set rs = Nothing

If Request.Form.Item("submit_delete") = "submit_delete" Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Active FROM EventAds WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    rs(0).Value = "n"
    rs.Update
    rs.Close
    Set rs = Nothing
ElseIf Request.Form.Item("submit_ad") = "submit_ad" Then
    Dim bCreateAd

    bCreateAd = False
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT Active FROM EventAds WHERE EventID = " & lEventID
    rs.Open sql, conn, 1, 2
    If rs.RecordCount > 0 Then
        rs(0).Value="y"
        rs.Update
    Else
        bCreateAd = True
    End If
    rs.Close
    Set rs = Nothing

    If bCreateAd = True Then
        sql = "INSERT INTO EventAds (EventID, Image) VALUES (" & lEventID & ", '" & sLogo & "')"
        Set rs = conn.Execute(sql)
        Set rs = Nothing
    End If
End If

sAdExists = "n"
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT AdViews, Active FROM EventAds WHERE EventID = " & lEventID
rs.Open sql, conn, 1, 2
If rs.RecordCount > 0 Then
    If rs(1).Value = "y" Then
        sAdExists = "y"
        lngNumViews = rs(0).Value
    End If
End If
rs.Close
Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title><%=sEventName%> Event Ad Admin</title>
<!--#include file = "../../includes/meta2.asp" -->

</head>

<body>
<div style="padding: 10px;margin: 10px;background-color: #fff;">
	<h3 class="h3"><%=sEventName%> Event Ad Management</h3>
		
	<div style="margin-top:10px;">
        <%If Not sLogo & "" = "" Then%>
            <img src="/events/logos/<%=sLogo%>" style="float: right;width: 150px;">
        <%End If%>

        <%If sAdExists = "y" Then%>
            <p>This event has an ad.  It has had <span style="font-weight:bold;"><%=lngNumViews%></span> views.  If you would like to deactivate this 
                ad "Yes" below and submit the form.</p>

            <form name="delete_ad" method="post" action="event_ad.asp?event_id=<%=lEventID%>">
            <select name="delete" id="delete">
                <option value="n">No</option>
                <option value="y">Yes</option>
            </select>
            <input type="hidden" name="submit_delete" id="submit_delete" value="submit_delete">
            <input type="submit" name="submit1" id="submit1" value="Deactivate Ad">
            </form>
        <%Else%>
            <p>This event either does not have an ad or it is inactive.  Would you like to create one or activate it?</p>

            <form name="insert_ad" method="post" action="event_ad.asp?event_id=<%=lEventID%>">
            <input type="checkbox" name="make_ad" id="make_ad">
            <input type="hidden" name="submit_ad" id="submit_ad" value="submit_ad">
            <input type="submit" name="submit2" id="submit2" value="Submit Ad">
            </form>
        <%End If%>
	</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>