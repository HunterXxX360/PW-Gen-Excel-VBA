Sub EMail(pw As String, hash As String, Optional AttPath As String)
Dim OutApp As Object
Dim objOMail As Object

Set OutApp = CreateObject("Outlook.Application")
Set objOMail = OutApp.CreateItem(0)

With objOMail
    .To = GetEntscheiderMail(EntscheiderNum)
    .Subject = "For Review"
    ' please add your e-mail for test@test.com
    .htmlBody = "<p>" & _
        "<a href=""mailto:test@test.com?body=" & pw & hash & "&amp;subject=Signature"">Sign</a>" & _
        "</p>" & _
        "<p>" & _
        "<a href=""mailto:test@test.com?body=delete%20key" & hash & "&amp;subject=Rejection"">Reject</a>" & _
        "</p>"
    If IsMissing(AttPath) = False Then
        .Attachments.Add AttPath
    End If
    .Display
End With

End Sub
