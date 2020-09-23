<div align="center">

## CPerformance


</div>

### Description

This class encapsulate QueryPerfomanceXXX

API functions to mesure small time intervals. You can use this class

to mesure how much time your code take. This function can mesure time

intervals near 0.1 ms , 10 times better then timeGetTime() API or

GetTickCount() that have an error of 50ms.

Example:

Dim m_performance As CPerformance

Dim i As integer

Set m_performance = new CPerformance

m_performance.StartCounter()

'Do something

For i = 1 to 1000

next i

m_performance.StopCounter()

Debug.print m_performance.TimeElapsed() 'Time in ms (1/1000) s

'this is a float number

'ex: 1.54 ms
 
### More Info
 
None

Time interval in ms.

The API function maybe not work, but it's very rare.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ricardo Saat](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ricardo-saat.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ricardo-saat-cperformance__1-1088/archive/master.zip)





### Source Code

```
Option Explicit
Private Type LARGE_INTEGER
  lowpart As Long
  highpart As Long
End Type
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private m_PerfFrequency As LARGE_INTEGER
Private m_CounterStart As LARGE_INTEGER
Private m_CounterEnd As LARGE_INTEGER
Private m_crFrequency As Currency
Private m_bEnable As Boolean
'mesure time that the code take jus to call functions
Property Get Delay() As Double
 Dim i As Integer
 Dim crTotalcount As Currency
 For i = 1 To 100
 Me.StartCounter
 Me.StopCounter
 crTotalcount = crTotalcount + (Large2Currency(m_CounterEnd) - Large2Currency(m_CounterStart))
 Next i
 Delay = ((crTotalcount / 100) / m_crFrequency) * 1000#
End Property
Private Function Large2Currency(largeInt As LARGE_INTEGER) As Currency
 If (largeInt.lowpart) > 0& Then
    Large2Currency = largeInt.lowpart
  Else
    Large2Currency = CCur(2 ^ 31) + CCur(largeInt.lowpart And &H7FFFFFFF)
  End If
  Large2Currency = Large2Currency + largeInt.highpart * CCur(2 ^ 32)
End Function
Private Sub Class_Initialize()
  Dim lResp As Long
  m_bEnable = CBool(QueryPerformanceFrequency(m_PerfFrequency))
  If m_bEnable Then
  End If
  m_crFrequency = Large2Currency(m_PerfFrequency)
  Debug.Assert m_bEnable 'Computer does not suport PerfCounter
End Sub
Public Sub StartCounter()
Dim lResp As Long
lResp = QueryPerformanceCounter(m_CounterStart)
End Sub
Public Sub StopCounter()
Dim lResp As Long
lResp = QueryPerformanceCounter(m_CounterEnd)
End Sub
Property Get TimeElapsed() As Double
  Dim crStart As Currency
  Dim crStop As Currency
  Dim crFrequency As Currency
  crStart = Large2Currency(m_CounterStart)
  crStop = Large2Currency(m_CounterEnd)
  TimeElapsed = ((crStop - crStart) / m_crFrequency) * 1000#
End Property
```

