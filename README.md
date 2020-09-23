<div align="center">

## A simple Wait Function


</div>

### Description

Ur code waits for X seconds without stopping complete VB like most others (incl. Sleep Api)
 
### More Info
 
'Set the API Declaration and the Code into a Module

'to use the function Wait(x) x stands for the amount of seconds to wait


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[PrixM](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/prixm.md)
**Level**          |Beginner
**User Rating**    |4.3 (39 globes from 9 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/prixm-a-simple-wait-function__1-7915/archive/master.zip)

### API Declarations

```
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
```


### Source Code

```
Public Function Wait(ByVal TimeToWait As Long) 'Time in seconds
 Dim EndTime   As Long
 EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
 Do Until GetTickCount > EndTime
  DoEvents
 Loop
End Function
```

