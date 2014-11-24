Option Explicit

Sub Locker()
Dim pw As String
Dim AP As String
Dim pwS As String

pw = Hash.GetString(128, 128)

AP = ThisWorkbook.Path & "\" & Workbooks(1).Name 'attachement path

pwS = SaltAndPepper(pw)

Worksheets(1).Protect pw
Workbooks(1).Close True, ThisWorkbook.Path & "\" & Hash.Hash8(pwS) & ".xlsx"

SplitKeys pwS, AP

End Sub



Sub SplitKeys(EKey As String, Optional EAttP As String)
Dim i As Integer
Dim PartS As String
Dim LnSS As Long
Dim ToS As Integer

ToS = 2 'number of recipients

    For i = 1 To ToS
        LnSS = Round(Len(EKey) / ToS, 0)
        
        If i <= 1 Then
            PartS = Mid(EKey, 1, LnSS)
        ElseIf i > 1 And i <> ToS Then
            PartS = Mid(EKey, 1 + (i - 1) * LnSS, LnSS)
        ElseIf i = ToS Then
            PartS = Mid(EKey, 1 + (i - 1) * LnSS)
        End If
        
        If IsMissing(EAttP) = False Then
            EMail.EMail PartS, EAttP
        Else
            EMail.EMail PartS
        End If
    Next i
End Sub



Function ReadKey(Path) As String
Dim OutApp As Object
Dim objOMail As Object
Dim hBody As String
Dim StrStart As Double
Dim StrEnd As Double
Dim StrLn As Double
Dim MrkStart As String
Dim MrkStop As String
Dim pw As String
Dim Sign As String

Set OutApp = CreateObject("Outlook.Application")

Set objOMail = OutApp.CreateItemFromTemplate(Path)

hBody = objOMail.Body
Sign = objOMail.SenderName

MrkStart = "##Mark##"
MrkStop = "##Mark End##"

StrStart = InStr(1, hBody, MrkStart) + Len(MrkStart) + 2    '+2 because of 2 <br>s
StrEnd = InStr(1, hBody, MrkStop)
StrLn = StrEnd - StrStart - 2   '-2 because of 2 <br>s

pw = Mid(hBody, StrStart, StrLn)

ReadKey = pw

End Function



Sub Unlocker(pw As String, hashStr As String)
Dim WBName As String

WBName = ThisWorkbook.Path & "\" & hashStr & ".xlsx"

Workbooks.Open WBName

If pw <> "revoke key" Then
    
    If Hash.Hash8(pw) = hashStr Then
      Workbooks(hashStr & ".xlsx").Worksheets(1).Unprotect pw
    End If
    
ElseIf pw = "revoke key" Then
    MsgBox "key was revoked, signature denied!", vbOKOnly + vbExclamation
End If

End Sub
