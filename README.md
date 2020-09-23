<div align="center">

## clsTimer


</div>

### Description

This code keeps a count in milliseconds of how long it takes between calls of StartTimer and StopTimer or Elapsed.
 
### More Info
 
Make a class called clsTimer and paste this in there.

Call order is something like:

Dim t1 as new clsTimer

t1.StartTimer

'do something that takes a while

Debug.Print "Right now, current elapsed = " & t2.Elapsed

'do something else?

t2.StopTimer

Debug.Print "Total elapsed = " & t2.Elapsed

.Elapsed returns the number of milliseconds between calls to .StartTimer and .StopTimer. If .StopTimer hasn't been called, it returns the number of milliseconds since .StartTimer was called.

No side effects.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Lambert](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-lambert.md)
**Level**          |Unknown
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-lambert-clstimer__1-937/archive/master.zip)

### API Declarations

```
Private Declare Function GetTickCount Lib "kernel32" () As Long
```


### Source Code

```
' in clsTimer...
Dim start, finish
Public Sub StopTimer()
  finish = GetTickCount()
End Sub
Public Sub StartTimer()
  start = GetTickCount()
  finish = 0
End Sub
Public Sub DebugTrace(v)
  Debug.Print v & " " & Elapsed()
End Sub
Public Property Get Elapsed()
  If finish = 0 Then
    Elapsed = GetTickCount() - start
  Else
    Elapsed = finish - start
  End If
End Property
```

