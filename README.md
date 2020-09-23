<div align="center">

## Count Days Between Two Dates \(with or without weekends\)


</div>

### Description

Counts the number of days between two dates with a choice to include/exclude the number weekend days in the total count. This can be used to calculate items like day-based thresholds.

This can modified (by your own modifications) to exclude specific days like holidays.
 
### More Info
 
dtFirstDate

dtSecondDate

fNoWeekend

Integer number of days between two dates.

None known.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Todd Main](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/todd-main.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/todd-main-count-days-between-two-dates-with-or-without-weekends__1-8033/archive/master.zip)





### Source Code

```
Option Explicit
Public Sub Test_CountDays()
'Number of Days between now and 10 days ago, excluding all weekend days
MsgBox CountDays(Now - 10, Now, True)
End Sub
Public Function CountDays( _
          dtFirstDate As Date, _
          dtSecondDate As Date, _
          Optional fNoWeekend As Boolean = True _
          ) As Integer
Dim dtFirstDateTemp   As Date   'Hold date to do calculations with
dtFirstDateTemp = dtFirstDate
Dim intWeekendDays   As Integer 'Holds weekend days
If dtFirstDate > dtSecondDate Then
  Exit Function  'Stops you from messing up this calculation, returns "0"
Else
  If fNoWeekend = True Then
    Do
      If (Weekday(dtFirstDateTemp) Mod 6 = 1) Then
        intWeekendDays = intWeekendDays + 1
      End If
      dtFirstDateTemp = DateAdd("d", 1, dtFirstDateTemp)
    Loop Until DateSerial(Year(dtFirstDateTemp), _
          Month(dtFirstDateTemp), _
          Day(dtFirstDateTemp)) _
          = DateSerial(Year(dtSecondDate), _
          Month(dtSecondDate), _
          Day(dtSecondDate))
    CountDays = CInt(DateDiff("d", dtFirstDate, dtSecondDate - intWeekendDays))
  Else
    CountDays = CInt(DateDiff("d", dtFirstDate, dtSecondDate))
  End If
End If
End Function
```

