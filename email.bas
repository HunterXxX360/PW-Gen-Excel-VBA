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
        "<a href=""mailto:test@test.com?body=0D%0A%23%23Mark%23%23%0D%0A" & pw & hash & "%0D%0A%23%23Mark%20End%23%23&amp;subject=Signature"">Sign</a>" & _
        "</p>" & _
        "<p>" & _
        "<a href=""mailto:test@test.com?body=0D%0A%23%23Mark%23%23%0D%0Adelete%20key" & hash & "%0D%0A%23%23Mark%20End%23%23&amp;subject=Rejection"">Reject</a>" & _
        "</p>"
    If IsMissing(AttPath) = False Then
        .Attachments.Add AttPath
    End If
    .Display
End With

End Sub
