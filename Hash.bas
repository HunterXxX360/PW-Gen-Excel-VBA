Option Explicit

Function GetString(Obergrenze As Integer, Untergrenze As Integer) As String
Dim Rand As String
Dim Nx As String
Dim getLen As Integer
Dim i As Integer

Application.Volatile
getLen = Int((Obergrenze + 1 - Untergrenze) * Rnd + Untergrenze)

Do
    i = i + 1
    Randomize
    Nx = Chr(38)
    'following symbols are forbidden, due to incompatibilities with html-formatting:
    ' &
    ' '
    ' \
    Do Until Nx <> Chr(38) And Nx <> Chr(39) And Nx <> Chr(92)
        Nx = Chr(Int((85) * Rnd + 38))
    Loop
    Rand = Rand & Nx
Loop Until i = getLen
 
GetString = Rand

End Function



Function SaltAndPepper(strapStr As String) As String
Dim SaPStr As String
Dim InfStr As String
Dim InfStrHex As String
Dim R As Integer
Dim i As Integer
Dim y As Integer
Dim yE As Integer
Dim LnSaP
Dim d As Long

SaPStr = GetString(512, 512) 'choose length of string here

y = 1
LnSaP = Len(SaPStr)

For i = 0 To 508 Step 4
    R = rndInt(4, 1)
    SaPStr = Left(SaPStr, i + R - 1) & Mid(strapStr, y, 1) & Right(SaPStr, LnSaP - (i + R))
    InfStr = InfStr & R
    y = y + 1
Next i

SaltAndPepper = SaPStr

'##################################################################################
'###                security warning -> saved inside workbook                   ###
'##################################################################################
With ThisWorkbook.Sheets(1)
    d = Dwn(ThisWorkbook.Worksheets(1).Range("A1")) + 1
    .Cells(d, 2).Value = "'" & InfStr
    .Cells(d, 1).Value = "'" & Hash8(SaPStr)
End With

SaPStr = ""

End Function



Function StrapString(SaPStr As String, InfStr As String) As String
Dim strapStr As String
Dim i As Integer
Dim uI As Integer
Dim y As Integer

y = 1

For i = 0 To 508 Step 4
    uI = CInt(Mid(InfStr, y, 1))
    strapStr = strapStr & Mid(SaPStr, i + uI, 1)
    y = y + 1
Next i

StrapString = strapStr

End Function



Function DeHashStr(SaPHStr As String) As String
Dim l As Long

l = Len(SaPHStr) - 8

DeHashStr = Left(SaPHStr, l)

End Function



Function FHashStr(SaPHStr As String) As String

FHashStr = Right(SaPHStr, 8)

End Function



Function rndInt(Obergrenze As Integer, Optional Untergrenze As Integer = 0)
Dim iZ As Integer
Randomize Timer
    iZ = Int((Obergrenze - Untergrenze + 1) * Rnd + Untergrenze)
rndInt = iZ
End Function



Function Hash4(TString As String) As String
Dim X As Long
Dim mask, i, j, nC, crc As Integer
Dim c As String

crc = &HFFFF

For nC = 1 To Len(TString)
    j = Asc(Mid(TString, nC))
    crc = crc Xor j
    For j = 1 To 8
        mask = 0
        If crc / 2 <> Int(crc / 2) Then mask = &HA001
        crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
    Next j
Next nC

c = Hex$(crc)

While Len(c) < 4
    c = "0" & c
Wend

Hash4 = c
    
End Function



Function Hash12(TString As String) As String
Dim l As Integer, l3 As Integer
Dim s1 As String, s2 As String, s3 As String
Dim c As Boolean

l = Len(TString)
l3 = Int(l / 3)
s1 = Mid(TString, 1, l3)
s2 = Mid(TString, l3 + 1, l3)
s3 = Mid(TString, 2 * l3 + 1)

Hash12 = Hash4(s1) & Hash4(s2) & Hash4(s3)

End Function



Function Hash8(TString As String) As String
Dim l As Integer, l2 As Integer
Dim s1 As String, s2 As String
Dim c As Boolean

l = Len(TString)
l2 = Int(l / 2)
s1 = Mid(TString, 1, l2)
s2 = Mid(TString, l2 + 1)

Hash8 = Hash4(s1) & Hash4(s2)

End Function



Function Dwn(Trgt As Range) As Long
Dim X As Long
Dim y As Long
Dim WS As String

X = Trgt.Column
y = Trgt.Row
WS = Trgt.Worksheet.Name

If Sheets(WS).Cells(y + 1, X).Value <> "" Then
    Dwn = Sheets(WS).Cells(y, X).End(xlDown).Row
Else
    Dwn = y
End If

End Function
