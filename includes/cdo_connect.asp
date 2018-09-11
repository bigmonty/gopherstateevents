<%
Set cdoConfig = CreateObject("CDO.Configuration")
With cdoConfig.Fields
    .Item(cdoSendUsingMethod) = cdoSendUsingPort
    .Item(cdoSMTPServer) = "smtp.mandrillapp.com"
    .Item(cdoSMTPAuthenticate) = 1
    .Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
    .Item(cdoSendUsername) = "bob.schneider@gopherstateevents.com"
    .Item(cdoSendPassword) = "H49iry1SZKdY7PQ5afpfyg"
    .Update
End With
%>
 