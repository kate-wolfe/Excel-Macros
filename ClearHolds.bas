Attribute VB_Name = "ClearKate"
Option Explicit

Sub ClearHoldsKate()

'Make things go fast.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False



'Declare Variables

Dim ClrHlds As Worksheet
Set ClrHlds = ThisWorkbook.Sheets("Clear Holds")

Dim i As Integer


'Sort sheet by clear holdshelf status to group all the "Hold Expired" stuff together.
ClrHlds.Range("A1", Range("E1").End(xlDown)).Sort Key1:=ClrHlds.Range("E1"), Order1:=xlAscending, Header:=xlNo


'Delete HOLD EXPIRED rows

Dim expRng As Range
Set expRng = ClrHlds.Range("E1", Range("E1").End(xlDown))
    
For i = expRng.Cells.Count To 1 Step -1
    If expRng.Item(i).Value = "HOLD EXPIRED" Then
        expRng.Item(i).EntireRow.Delete
    End If
Next i


'Right 4 of barcode and shorten names

ClrHlds.Columns("D").NumberFormat = "0000"

Dim fullName As String
Dim spltName() As String

Dim lastFour As String
Dim firstIni As String

Dim newRng As Range
Set newRng = ClrHlds.Range("E1", Range("E1").End(xlDown))

For i = newRng.Cells.Count To 1 Step -1

    ClrHlds.Cells(i, 4).Value2 = Right(ClrHlds.Cells(i, 4).Value2, 4)

    If InStr(1, ClrHlds.Cells(i, 1).Value2, "(") > 0 Then
        ClrHlds.Cells(i, 1).Replace "*(", vbNullString, xlPart
        ClrHlds.Cells(i, 1).Replace ")", vbNullString, xlPart
    ElseIf InStr(1, ClrHlds.Cells(i, 1).Value2, ",") = 0 Then
        fullName = ClrHlds.Cells(i, 1).Value2
        spltName() = Split(fullName, ",")
        lastFour = UCase(Left(spltName(0), 4))
        ClrHlds.Cells(i, 1).Value2 = lastFour
    Else
        fullName = ClrHlds.Cells(i, 1).Value2
        spltName() = Split(fullName, ",")
        lastFour = UCase(Left(spltName(0), 4))
        firstIni = UCase(Left(spltName(1), 2))
        ClrHlds.Cells(i, 1).Value2 = lastFour & "," & firstIni
    
    End If
    
Next i

'Do Stats

Dim OnShelves As Range
Set OnShelves = ClrHlds.Range("O1")

Dim ToBeCleared As Integer
ToBeCleared = ClrHlds.Range("C2000").End(xlUp).Row

Dim statsWB As Workbook
Dim statsSH As Worksheet

Dim StatsDate As Range
Dim StatsTotal As Range
Dim StatsPull As Range

Set statsWB = Workbooks.Open _
    ("\\coc\Library\Borrower Services\Clear Holds\Clear Holds Stats.xlsx")
Set statsSH = statsWB.Sheets("Stats")

Dim statsLast As Long
statsLast = statsSH.Cells(Rows.Count, 1).End(xlUp).Row

Set StatsDate = statsSH.Range("A" & statsLast).Offset(1)
Set StatsTotal = statsSH.Range("B" & statsLast).Offset(1)
Set StatsPull = statsSH.Range("C" & statsLast).Offset(1)

StatsDate.Value2 = Date
StatsTotal.Value2 = OnShelves.Value2
StatsPull.Value2 = ToBeCleared

statsWB.Save
statsWB.Close

ClrHlds.Activate

'Sort by name

ClrHlds.Range("A1", Range("E1").End(xlDown)).Sort Key1:=ClrHlds.Range("A1"), Order1:=xlAscending, Header:=xlNo


'Format Sheet

With ClrHlds
.PageSetup.LeftHeader = "&A"
.PageSetup.CenterHeader = "&D &T"
.PageSetup.RightHeader = "&P of &N"
End With

Dim SheetEnd As Range

ClrHlds.Columns("A:A").ColumnWidth = 25
ClrHlds.Columns("B:B").ColumnWidth = 30
ClrHlds.Columns("C:E").Font.Size = 10
ClrHlds.Columns("C:C").ColumnWidth = 14.5
    
With ClrHlds.Columns("D:D")
    .ColumnWidth = 5.5
    .HorizontalAlignment = xlLeft
End With
    
ClrHlds.Columns("E:E").ColumnWidth = 18

ClrHlds.Rows(1).Insert shift:=xlDown

ClrHlds.Cells(1, 1).Value2 = "Patron"
ClrHlds.Cells(1, 2).Value2 = "Title"
ClrHlds.Cells(1, 3).Value2 = "Call Number"
ClrHlds.Cells(1, 4).Value2 = "Last4"
ClrHlds.Cells(1, 5).Value2 = "Clear Hold"
ClrHlds.Rows("1:1").Font.Bold = True
ClrHlds.Rows("1:1").Font.Size = 11

Set SheetEnd = ClrHlds.Cells(Rows.Count, 5).End(xlUp)

With ClrHlds.Range("A1", SheetEnd).Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
End With

With ClrHlds.PageSetup
   .Orientation = xlPortrait
   .LeftMargin = Application.InchesToPoints(0.2)
   .RightMargin = Application.InchesToPoints(0.2)
   .TopMargin = Application.InchesToPoints(0.4)
   .BottomMargin = Application.InchesToPoints(0.2)
   .HeaderMargin = Application.InchesToPoints(0.2)
   .FooterMargin = Application.InchesToPoints(0.2)
End With

ClrHlds.PageSetup.printarea = "A:E"

End Sub
