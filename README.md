<div align="center">

## HTTPSafeString


</div>

### Description

Makes a string http querystring friendly by replacing all non-alpha and non-numeric characters with the appropriate hex code. Helpful when using the wininet API.

example: "(Find This)" becomes "%28Find%20This%29"
 
### More Info
 
Text as String

String


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Greg Tyndall](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/greg-tyndall.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/greg-tyndall-httpsafestring__1-6069/archive/master.zip)





### Source Code

```
Public Function HTTPSafeString(Text As String) As String
  Dim lCounter As Long
  Dim sBuffer As String
  Dim sReturn As String
  sReturn = Text
  For lCounter = 1 To Len(Text)
    sBuffer = Mid(Text, lCounter, 1)
    If Not sBuffer Like "[a-z,A-Z,0-9]" Then
      sReturn = Replace(sReturn, sBuffer, "%" & Hex(Asc(sBuffer)))
    End If
  Next lCounter
  HTTPSafeString = sReturn
End Function
```

