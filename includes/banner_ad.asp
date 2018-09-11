<%
    Dim lBannerAdsIDl, lEventFamilyID, lThisEventID, lBannerAdsID
    Dim sAdURL, sAdImage, sAdAlt

    lEventFamilyID = 0

'    Set rs = Server.CreateObject("ADODB.Recordset")
'    sql = "SELECT EventFamilyID FROM Events WHERE EventID = " & lThisEventID
'    rs.Open sql, conn, 1, 2
'    If Not rs(0).Value & "" = "" Then lEventFamilyID = rs(0).Value
'    rs.Close
'    Set rs = Nothing

	Set rs = Server.CreateObject("ADODB.Recordset")
	If CLng(lEventFamilyID) = 0 Then
        sql = "SELECT BannerAdsID, AdURL, AdImage, AdAlt, AdViews FROM BannerAds ORDER BY NewID()"
    Else
        sql = "SELECT BannerAdsID, AdURL, AdImage, AdAlt, AdViews FROM BannerAds WHERE EventFamilyID = " & lEventFamilyID
    End If
	rs.Open sql, conn, 1, 2
	lBannerAdsID = rs(0).Value
    sAdURL = rs(1).Value
    sAdImage = rs(2).Value
    sAdAlt = rs(3).Value
    rs(4).Value = CSng(rs(4).Value) + 1
    rs.Update
	rs.Close
	Set rs = Nothing

	sql = "INSERT INTO AdViews (BannerAdsID, WhenViewed, IPAddress) VALUES (" & lBannerAdsID & ", '" & Now()  & "', '" 
    sql = sql & Request.ServerVariables("REMOTE_ADDR") & "')"
    Set rs = conn.Execute(sql)
	Set rs = Nothing
%>

        <div style="margin:5px 0 0 0;padding:0;text-align:center;">
			<a href="<%=sAdURL%>" onclick="openThis(this.href,1024,760);return false;"><img src="/graphics/banner_ads/<%=sAdImage%>" alt="<%=sAdAlt%>"
                 style="width: 610px;"></a>
		</div>
