<%
Function ValidEmail(sThisEmail) 
    Dim atCnt
    Dim x

	ValidEmail = True

	If Len(sThisEmail) <= 5 Then 
		ValidEmail = False
	ElseIf InStr(1, sThisEmail, "@", 1) < 2 Then
		ValidEmail = False
	ElseIf InStr(1, sThisEmail, " ", 1) > 0 Then
		ValidEmail = False
    ElseIf InStr(sThisEmail,"_") <> 0 and InStrRev(sThisEmail,"_") > InStrRev(sThisEmail,"@")  Then  '  has no "_" after the "@"
        ValidEmail = False
    Else
        atCnt = 0                             ' has only one "@"
        For x = 1 to Len(sThisEmail)
            If  Mid(sThisEmail,x,1) = "@" Then
                atCnt = atCnt + 1
            End If
        Next
        If atCnt <> 1 Then ValidEmail = False

      ' chk each char for validity
        For x = 1 to Len(sThisEmail)
            If  Not IsNumeric(Mid(sThisEmail,x,1)) And (LCase(Mid(sThisEmail,x,1)) < "a" Or LCase(Mid(sThisEmail,x,1)) > "z") And Mid(sThisEmail,x,1) <> "_" And Mid(sThisEmail,x,1) <> "." And Mid(sThisEmail,x,1) <> "@" And Mid(sThisEmail,x,1) <> "+" And Mid(sThisEmail,x,1) <> "-" Then
                ValidEmail = False
            End If
        Next
	End If
End Function

%>

