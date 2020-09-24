<div align="center">

## RemoveWeekEnds


</div>

### Description

Returns number of business days from a start date and total number of days. There are more pieces coming. Building a routine that calculates holidays and weekends to tell you strictly business days.
 
### More Info
 
strStartDate and totalsdays

total number of business days


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ustes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ustes.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ustes-removeweekends__1-8953/archive/master.zip)





### Source Code

```
Public Function RemoveWeekends(strStartDate As String, intNumberOfDays) As Integer
  Dim i As Integer
  For i = 0 To intNumberOfDays
    Select Case Weekday(DateAdd("d", i, CDate(strStartDate)))
      Case vbSaturday, vbSunday
        intNumberOfDays = intNumberOfDays - 1
    End Select
  Next i
  RemoveWeekends = intNumberOfDays
End Function
```

