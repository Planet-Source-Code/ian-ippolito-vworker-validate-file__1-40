<div align="center">

## Validate\_File


</div>

### Description

Determines if a file exists

Improved version--detects hidden files too!
 
### More Info
 
filename--file to validate

'returns true or false


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Unknown
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\)
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-validate-file__1-40/archive/master.zip)





### Source Code

```
Function Validate_File (ByVal FileName As String) As Integer
Dim fileFile As Integer
  'attempt to open file
  fileFile = FreeFile
  On Error Resume Next
  Open FileName For Input As fileFile
  'check for error
  If Err Then
    Validate_File = False
  Else
    'file exists
    'close file
    Close fileFile
    Validate_File = True
  End If
End Function
```

