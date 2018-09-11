<%@ Language=VBScript%>
<%
Option Explicit

Dim conn, rs, sql
Dim lEventID
Dim sEventName
Dim iThisBib
Dim dEventDate
Dim i, j, k, m
Dim RaceBibs(), AssgndBibs(), AvailBibs(), RaceArray(), RaceParts()

'If Session("role") = "admin" Then Response.Redirect "/default.asp?sign_out=y"

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
    
If Request.Form.Item("submit_bibs") = "submit_bibs" Then
	ReDim RaceArray(0)
	i = 0
	
	'get the race ids
	sql = "SELECT RaceID FROM RaceData WHERE EventID = " & lEventID
	Set rs = conn.Execute(sql)
	Do While Not rs.EOF
	    RaceArray(i) = rs(0).Value
	    i = i + 1
	    ReDim Preserve RaceArray(i)
	    rs.MoveNext
	Loop
	Set rs = Nothing
	
	For i = 0 To UBound(RaceArray) - 1
	    'get race bibs to an array
	    k = 0
	    ReDim RaceBibs(0)
	    sql = "SELECT BibsFrom, BibsTo FROM RaceData WHERE RaceID = " & RaceArray(i)
	    Set rs = conn.Execute(sql)
	    For j = rs(0).Value To rs(1).Value
	        RaceBibs(k) = j
	        k = k + 1
	        ReDim Preserve RaceBibs(k)
	    Next
	    Set rs = Nothing
	
	    'get all the bibs already assigned to this race
	    k = 0
	    ReDim AssgndBibs(0)
		Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT pr.Bib FROM PartRace pr iNNER JOIN Participant p ON pr.ParticipantID = p.ParticipantID Where pr.RaceID = " & RaceArray(i) 
        sql = sql & " ORDER BY pr.Bib"
	    rs.Open sql, conn, 1, 2
	    Do While Not rs.EOF
	        If Not rs(0).Value & "" = "" Then
	            AssgndBibs(k) = rs(0).Value
	            k = k + 1
	            ReDim Preserve AssgndBibs(k)
	        End If
	        rs.MoveNext
	    Loop
	    rs.Close
	    Set rs = Nothing
	    
	    'now write the available bibs to an array
	    k = 0
	    ReDim AvailBibs(1, 0)
	    For m = 0 To UBound(RaceBibs) - 1
	        If UBound(AssgndBibs) = 0 Then      'if there are no bibs assigned yet then just move all race bibs to avail bibs
	            AvailBibs(0, m) = RaceBibs(m)   'this is the bib (AvailBibs(1, m) will be for flaggin when assigned)
	            ReDim Preserve AvailBibs(1, m + 1)
	        Else
	            For j = 0 To UBound(AssgndBibs) - 1
	                If AssgndBibs(j) = RaceBibs(m) Then     'if a race bib has been assigned then move on
	                    Exit For
	                Else                                     'if no match is found make this bib available
	                    If j = UBound(AssgndBibs) - 1 Then
	                        AvailBibs(0, k) = RaceBibs(m)
	                        k = k + 1
	                        ReDim Preserve AvailBibs(1, k)
	                    End If
	                End If
	            Next
	        End If
	    Next
	
        'sort bibs numerically
        For k = 0 To UBound(AvailBibs, 2) - 2
            For j = k + 1 To UBound(AvailBibs, 2) - 1
                If CInt(AvailBibs(0, k)) > CInt(AvailBibs(0, j)) Then
                    iThisBib = AvailBibs(0, k)
                    AvailBibs(0, k) = AvailBibs(0, j)
                    AvailBibs(0, j) = iThisBib
                End If 
            Next
        Next

	    'get all participants in this race with no bibl, sorted alpha
	    k = 0
        ReDim RaceParts(0)
		Set rs = Server.CreateObject("ADODB.Recordset")
	    sql = "SELECT p.ParticipantID, pr.Bib FROM Participant p iNNER JOIN PartRace pr ON p.ParticipantID = pr.ParticipantID Where pr.RaceID = " 
        sql = sql & RaceArray(i) & " ORDER BY p.LastName, p.FirstName"
	    rs.Open sql, conn, 1, 2
	    Do While Not rs.EOF
	        If rs(1).Value & "" = "" Then
	            RaceParts(k) = rs(0).Value
	            k = k + 1
                ReDim Preserve RaceParts(k)
	        End If
	        rs.MoveNext
	    Loop
	    rs.Close
	    Set rs = Nothing

        'now assign bibs
        For k = 0 To UBound(RaceParts) - 1
            For j = 0 To UBound(AvailBibs, 2) - 1
                If AvailBibs(1, j) & "" = "" Then
		            Set rs = Server.CreateObject("ADODB.Recordset")
	                sql = "SELECT Bib FROM PartRace WHERE RaceID = " & RaceArray(i) & " AND ParticipantID = " & RaceParts(k)
	                rs.Open sql, conn, 1, 2
	                rs(0).Value = AvailBibs(0, j)
                    rs.Update
	                rs.Close
	                Set rs = Nothing

                    AvailBibs(1, j) = "x"   'flag that this one has been used

                    Exit For
                End If
            Next
        Next
	Next

	Response.Write("<script type='text/javascript'>{window.close();}</script>")
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>

<title><%=sEventName%> Batch Bib Assignment Utility</title>
<!--#include file = "../../includes/meta2.asp" -->



</head>

<body>
<div style="margin:10px;background-color:#fff;">
	<h3><%=sEventName%> Batch Bib Assignment Utilty</h3>
	
	<form name="assign_bibs" method="Post" action="bib_assign.asp?event_id=<%=lEventID%>">
	<p>
		This will assign bibs to all participants in all races in this event who do not already have one assigned.  Remember to refresh the data on the
		main page in order to see the results of this action.
		<br>
		<input type="hidden" name="submit_bibs" id="submit_bibs" value="submit_bibs">
		<input type="submit" name="submit1" id="submit1" value="Assign Bibs" style="color:#d62002">
	</p>
	</form>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>