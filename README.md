<div align="center">

## BinaryCrypt


</div>

### Description

This application reduces ASCII character codes to binary and then shifts the bits to the left by whatever the length of the string is.
 
### More Info
 
You need to input a string to encrypt.

It returns your encrypted string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Trent Gardner](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/trent-gardner.md)
**Level**          |Intermediate
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 5\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/trent-gardner-binarycrypt__1-32945/archive/master.zip)





### Source Code

```
'************************************************
'*******  BinaryCrypt was written by  *******
'*******      Trent Gardner     *******
'************************************************
Public BinaryStrings As New Collection
Public strEncrypted As String
Public strDecrypted As String
Public Function BinaryCrypt(strBinary As String, BinaryShift As Integer)
  On Error Resume Next
  Dim BinaryPositions As New Collection
  Dim BinaryChange As New Collection
  '  128   64   32   16   8    4    2    1
  '  [ ]   [ ]   [ ]   [ ]  [ ]   [ ]   [ ]   [ ]
  For intCounter = 0 To 7
    BinaryPositions.Add (Mid(strBinary, Len(strBinary) - intCounter, 1))
  Next intCounter
  For Positions = 1 To BinaryShift
  strFinished = vbNullString
  For intCounter = 1 To 8
    ' Rotating to the left
    If intCounter = 1 Then
      EighthPosition = BinaryPositions.Item(1)
    Else
      BinaryChange.Add (BinaryPositions.Item(intCounter))
    If intCounter = 8 Then
      BinaryChange.Add (EighthPosition)
    End If
    End If
  Next intCounter
  For i = 1 To 4
    For intCounter = 1 To 4
      'BinaryChange.Remove (intCounter)
      BinaryPositions.Remove (intCounter)
    Next intCounter
  Next i
  For i = 1 To 8
    BinaryPositions.Add (BinaryChange(i))
  Next i
  For intCounter = 1 To BinaryChange.Count
    strFinished = strFinished & BinaryPositions.Item(intCounter)
  Next intCounter
  For i = 1 To 4
    For intCounter = 1 To 4
      BinaryChange.Remove (intCounter)
      'BinaryPositions.Remove (intCounter)
    Next intCounter
  Next i
  Next Positions
BinaryCrypt = strFinished
End Function
Public Function BinaryToAsc(strBinary As String)
  Dim BinaryPositions As New Collection
  Dim AscFigures As New Collection
  '  128   64   32   16   8    4    2    1
  '  [ ]   [ ]   [ ]   [ ]  [ ]   [ ]   [ ]   [ ]
  For intCounter = 0 To 7
    BinaryPositions.Add (Mid(strBinary, Len(strBinary) - intCounter, 1))
  Next intCounter
  AscFigures.Add (BinaryPositions.Item(1))
  AscFigures.Add (BinaryPositions.Item(2) * 2)
  AscFigures.Add (BinaryPositions.Item(3) * 4)
  AscFigures.Add (BinaryPositions.Item(4) * 8)
  AscFigures.Add (BinaryPositions.Item(5) * 16)
  AscFigures.Add (BinaryPositions.Item(6) * 32)
  AscFigures.Add (BinaryPositions.Item(7) * 64)
  AscFigures.Add (BinaryPositions.Item(8) * 128)
  For intCounter = 1 To AscFigures.Count
    intAsc = intAsc + CInt(AscFigures.Item(intCounter))
  Next intCounter
  BinaryToAsc = intAsc
End Function
Public Function AscToBinary(strText As String)
  Dim AscCollection As New Collection
  Dim TempChr As Integer
  '  128   64   32   16   8    4    2    1
  '  [ ]   [ ]   [ ]   [ ]  [ ]   [ ]   [ ]   [ ]
  For intCounter = 1 To Len(strText)
    strTemp = Asc(Mid(strText, intCounter, 1))
    AscCollection.Add (strTemp)
  Next intCounter
  For intCounter = 1 To AscCollection.Count
    TempChr = AscCollection.Item(intCounter)
    If (TempChr Mod 128) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (TempChr Mod 128)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    If (TempChr Mod 64) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (AscCollection.Item(intCounter) Mod 64)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    If (TempChr Mod 32) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (AscCollection.Item(intCounter) Mod 32)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    If (TempChr Mod 16) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (AscCollection.Item(intCounter) Mod 16)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    If (TempChr Mod 8) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (AscCollection.Item(intCounter) Mod 8)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    If (TempChr Mod 4) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (AscCollection.Item(intCounter) Mod 4)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    If (TempChr Mod 2) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (AscCollection.Item(intCounter) Mod 2)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    If (TempChr Mod 1) = TempChr Then
      strBinaryTemp = strBinaryTemp & "0"
    Else
      TempChr = (AscCollection.Item(intCounter) Mod 1)
      strBinaryTemp = strBinaryTemp & "1"
    End If
    BinaryStrings.Add (strBinaryTemp)
  Next intCounter
End Function
Public Function BinaryEncrypt(strText As String)
  On Error Resume Next
  strEncrypted = vbNullString
  For intCounter = 1 To Len(strText)
    strTemp = Mid(strText, intCounter, 1)
    AscToBinary (strTemp)
  Next intCounter
  For intCounter = 1 To BinaryStrings.Count
    strTemp = Chr(BinaryToAsc(BinaryCrypt(BinaryStrings.Item(intCounter), Len(strText) + 1)))
    strEncrypted = strEncrypted & strTemp
  Next intCounter
  For i = 1 To CInt((BinaryStrings.Count / 2) + 1)
    For intCounter = 1 To BinaryStrings.Count
      BinaryStrings.Remove (intCounter)
    Next intCounter
  Next i
  BinaryEncrypt = strEncrypted
End Function
Public Function BinaryDecrypt(strText As String)
  On Error Resume Next
  strDecrypted = vbNullString
  For intCounter = 1 To Len(strText)
    strTemp = Mid(strText, intCounter, 1)
    AscToBinary (strTemp)
  Next intCounter
  For intCounter = 1 To BinaryStrings.Count
    strTemp = Chr(BinaryToAsc(BinaryCrypt(BinaryStrings.Item(intCounter), Len(strText) + 1)))
    strDecrypted = strDecrypted & strTemp
  Next intCounter
  For i = 1 To CInt((BinaryStrings.Count / 2) + 1)
    For intCounter = 1 To BinaryStrings.Count
      BinaryStrings.Remove (intCounter)
    Next intCounter
  Next i
  BinaryDecrypt = strDecrypted
End Function
' You add it to your application as follows:
Private Sub cmdDecrypt_Click()
  MsgBox BinaryDecrypt(txtEncrypted.Text)
End Sub
Private Sub cmdEncrypt_Click()
  txtEncrypted.Text = BinaryEncrypt(txtPlain.Text)
End Sub
```

