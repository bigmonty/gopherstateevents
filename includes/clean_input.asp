<%
Function CleanInput(sInput)
	Dim char, x, bHackFound, sOrigInput
	
	sOrigInput = sInput
	bHackFound = False

	If InStr(UCase(sInput), "DECLARE") > 0 Then bHackFound = True

	If bHackFound = False Then
		If InStr(UCase(sInput), "IFRAME") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "STYLE=") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "HEIGHT=") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(sInput, ":8080") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "/TITLE") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), ".RU") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "WEBSERVICE") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "SCRIPT SRC=") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "INSERT INTO") > 0 Then bHackFound = True
	End If

	If bHackFound = False Then
		If InStr(UCase(sInput), "DELETE FROM") > 0 Then bHackFound = True
	End If
 
	If bHackFound = False Then
		If InStr(UCase(sInput), "NULL") > 0 Then bHackFound = True
	End If
 
	If bHackFound = False Then
		If InStr(UCase(sInput), ".EXE") > 0 Then bHackFound = True
	End If
	
	If bHackFound = True Then
		Dim cdoMessage, cdoConfig
		Dim sMsg
		
	    sMsg = "A possible attempt to hack h51software.net has occurred.  It has been blocked!" & VbCrLf & VbCrLf
	    sMsg = sMsg & "Time of Submission: " & Now() & VbCrLf & VbCrLf
	    sMsg = sMsg & "Text Submitted: " & sInput & " char " & char &  VbCrLf & VbCrLf
	    sMsg = sMsg & "Source IP " & Request.ServerVariables("REMOTE_ADDR") &  VbCrLf & VbCrLf
	    sMsg = sMsg & "Web Page: " & Request.ServerVariables("URL") &  VbCrLf & VbCrLf
	    sMsg = sMsg & "Host: " & Request.ServerVariables("HOST") &  VbCrLf & VbCrLf
	    sMsg = sMsg & "Length of String: " & len(sInput) & vbcrlf & vbcrlf
		
		Set cdoMessage = CreateObject("CDO.Message")
		With cdoMessage
			Set .Configuration = cdoConfig
			.To = "bob.schneider@gopherstateevents.com"
			.From = "bob.schneider@gopherstateevents.com"
			.Subject = "H51Software.net Hack Attempt"
			.TextBody = sMsg
			.Send
		End With
		Set cdoMessage = Nothing
		
		sHackMsg = "Some of your input appears to be an attempt to compromise the security of this website.  It has been sent to our security department for "
		sHackMsg = sHackMsg & "review.  If this was a genuine attempt to communicate please contact bob.schneider@gopherstateevents.com"
	Else
	    sInput = Replace(lcase(sInput), "http://", "")
	    sInput = Replace(lcase(sInput), "drop", "drp")
	    sInput = Replace(lcase(sInput), "xp_", "")
	    sInput = Replace(lcase(sInput), "CRLF", "")
	    sInput = Replace(lcase(sInput), "%3A", "")';
	    sInput = Replace(lcase(sInput), "%3B", "")':
	    sInput = Replace(lcase(sInput), "%3D", "equals")
	    sInput = Replace(lcase(sInput), "%3E", "grtr than")
	    sInput = Replace(lcase(sInput), "%3F", "")'?
	    sInput = Replace(lcase(sInput), "&quot;", "")
	    sInput = replace(lcase(sInput), "&amp;", "and")
	    sInput = replace(lcase(sInput), "&lt;", "lss than")
	    sInput = replace(lcase(sInput), "&gt;", "grtr than")
	    sInput = replace(lcase(sInput), " exec ", "")
	    sInput = replace(lcase(sInput), "onvarchar", "")
	    sInput = replace(lcase(sInput), "set", "")
	    sInput = replace(lcase(sInput), " cast ", "")
	    sInput = replace(lcase(sInput), "00100111", "")
	    sInput = replace(lcase(sInput), "00100010", "")
	    sInput = replace(lcase(sInput), "00111100", "")
	    sInput = replace(lcase(sInput), "select", "selct")
	    sInput = replace(lcase(sInput), "0x", "")
	    sInput = replace(lcase(sInput), "delete", "delet")
	    sInput = replace(lcase(sInput), "go ", "")
	    sInput = replace(lcase(sInput), "create", "creat")
	    sInput = replace(lcase(sInput), "convert", "cnvrt")
	    sInput = replace(lcase(sInput), "=", "equals")
	    sInput = replace(lcase(sInput), "/", "")
	    sInput = replace(lcase(sInput), "\", "")
	    sInput = replace(lcase(sInput), "?", "")
	    sInput = replace(lcase(sInput), "# ", " ")
	    sInput = replace(lcase(sInput), ";", "")
	    sInput = replace(lcase(sInput), ":", "")
	    sInput = replace(lcase(sInput), "$", "")
	    sInput = replace(lcase(sInput), "<", "lss than")
	    sInput = replace(lcase(sInput), ">", "grtr than")
	    sInput = replace(lcase(sInput), "(", "-")
	    sInput = replace(lcase(sInput), ")", "-")
	    sInput = replace(lcase(sInput), "+ ", "plus")
	    sInput = replace(lcase(sInput), "~", "")
	    sInput = replace(lcase(sInput), "|", "")
	    sInput = replace(lcase(sInput), "$", "")

		If LCase(sOrigInput) = sInput Then
	    	CleanInput = sOrigInput
		Else
			CleanInput = sInput
		End If
	End If
End Function
%>
