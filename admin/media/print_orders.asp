<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, conn2, rs, sql
Dim lEventID, lMeetID
Dim i
Dim MediaOrders

If Not Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

Response.Buffer = True		'Turn buffering on
Response.Expires = -1		'Page expires immediately
		
lEventID = Request.QueryString("event_id")
lMeetID = Request.QueryString("meet_id")
				
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=VIRA;Uid=broad_user;Pwd=Zeroto@123;"
							
Set conn2 = Server.CreateObject("ADODB.Connection")
conn2.Open "Provider=SQLNCLI11;Server=72.52.136.29;Database=CCMeet;Uid=broad_user;Pwd=Zeroto@123;"

'get order history
If Not (CLng(lEventID) = 0 AND CLng(lMeetID) = 0) Then
    Set rs = Server.CreateObject("ADODB.Recordset")
    If Not CLng(lEventID) = 0 Then
        sql = "SELECT MediaOrderID, WhichVid, BibNum, Email, AmtPd, MediaType, ClipStart, IPAddress, PmtLink, DatePaid FROM MediaOrder "
        sql = sql & "WHERE EventID = " & lEventID
        rs.Open sql, conn, 1, 2
    Else
        sql = "SELECT MediaOrderID, WhichVid, BibNum, Email, AmtPd, MediaType, ClipStart, IPAddress, PmtLink, DatePaid FROM MediaOrder "
        sql = sql & "WHERE MeetID = " & lMeetID
        rs.Open sql, conn2, 1, 2
    End If
    If rs.RecordCount > 0 Then 
        MediaOrders = rs.GetRows()
    Else
        ReDim MediaOrders(9, 0)
    End If
    rs.Close
    Set rs = Nothing
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title>GSE Print Media Manager</title>
<!--#include file = "../../includes/meta2.asp" -->

<script type="text/javascript" src="../../misc/scripts.js"></script>
<link rel="stylesheet" type="text/css" href="../../misc/styles.css">

<style type="text/css">
    th, td{
        padding-left: 10px;
        text-align: left;
    }
</style>
</head>

<body>
<div style="margin: 10px;padding: 5px;background-color: #fff;">
    <div style="margin: 10px 0 0 0;padding: 0;font-size: 0.85em;text-align: right;">
        <a href="javascript:print();">Print</a>
    </div>
    <h4 style="background-color: #ececec;text-align: left;">Existing Orders:</h4>
    <h5 style="text-align: left;">Num Orders:&nbsp;<%=UBound(MediaOrders, 2)%></h5>
    <table  style="font-size: 0.8em;">
        <tr>
            <th>Video</th><th>Bib</th><th>Email</th><th>AmtPd</th><th>MediaType</th>
            <th>ClipStart</th><th>IPAddres</th><th>PmtLink</th><th>DatePaid</th>
        </tr>
        <%For i = 0 To UBound(MediaOrders, 2)%>
            <tr>
                <%If i mod 2 = 0 Then%>
                    <td class="alt"><%=MediaOrders(1, i)%></td>
                    <td class="alt"><%=MediaOrders(2, i)%></td>
                    <td class="alt"><%=MediaOrders(3, i)%></td>
                    <td class="alt">$<%=MediaOrders(4, i)%></td>
                    <td class="alt"><%=MediaOrders(5, i)%></td>
                    <td class="alt"><%=MediaOrders(6, i)%></td>
                    <td class="alt"><%=MediaOrders(7, i)%></td>
                    <td class="alt"><%=MediaOrders(8, i)%></td>
                    <td class="alt"><%=MediaOrders(9, i)%></td>
                <%Else%>
                    <td><%=MediaOrders(1, i)%></td>
                    <td><%=MediaOrders(2, i)%></td>
                    <td><%=MediaOrders(3, i)%></td>
                    <td>$<%=MediaOrders(4, i)%></td>
                    <td><%=MediaOrders(5, i)%></td>
                    <td><%=MediaOrders(6, i)%></td>
                    <td><%=MediaOrders(7, i)%></td>
                    <td><%=MediaOrders(8, i)%></td>
                    <td><%=MediaOrders(9, i)%></td>
                <%End If%>
            </tr>
        <%Next%>
    </table>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing

conn2.Close
Set conn2 = Nothing
%>