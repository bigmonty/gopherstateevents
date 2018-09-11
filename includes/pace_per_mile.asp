<%
Private Function PacePerMile(sTime, sThisDist)
    Dim sglDist

    If UCase(Right(sThisDist, 2)) = "MI" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3))
    ElseIf UCase(Right(sThisDist, 4)) = "MILE" Then
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 5))
    ElseIf UCase(sThisDist) = "MARATHON" Then
        sglDist = 26.2
    ElseIf UCase(sThisDist) = "H. MAR" Then
        sglDist = 13.1
    Else
        sglDist = CSng(Left(sThisDist, Len(sThisDist) - 3)) * 0.6213712
    End If

    'calculate the pace
    PacePerMile = ConvertToMinutes(Round(CSng(ConvertToSeconds(Round(sTime, 2))) / Round(sglDist, 2), 3))
End Function
%>