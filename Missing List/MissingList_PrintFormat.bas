Attribute VB_Name = "PrintFormat"
Sub PageSetup()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Mezz.Columns("D").ColumnWidth = 13.33
Mezz.Columns("D").WrapText = True

'Loop

Dim SheetCount As Integer
Dim i As Integer
Dim LoopSheet As Worksheet

SheetCount = ThisWorkbook.Worksheets.Count

For i = 1 To SheetCount

Set LoopSheet = ThisWorkbook.Worksheets(i)

With LoopSheet.PageSetup
.LeftHeader = "&A"
.CenterHeader = "&D &T"
.RightHeader = "&P of &N"
.Orientation = xlLandscape
.PrintTitleRows = "$1:$1"
.LeftMargin = Application.InchesToPoints(0.25)
.RightMargin = Application.InchesToPoints(0.25)
.TopMargin = Application.InchesToPoints(0.75)
.BottomMargin = Application.InchesToPoints(0.75)
.HeaderMargin = Application.InchesToPoints(0.3)
.FooterMargin = Application.InchesToPoints(0.3)
End With

Dim wA As Double
Dim wB As Double
Dim wC As Double
Dim wD As Double
Dim wF As Double
Dim wG As Double
Dim wH As Double
Dim wI As Double
Dim wJ As Double
Dim AvailWidth As Double

LoopSheet.Columns("A:D").AutoFit
LoopSheet.Columns("F:J").AutoFit
LoopSheet.Columns("L:M").AutoFit

wA = LoopSheet.Columns("A").ColumnWidth
wB = LoopSheet.Columns("B").ColumnWidth
wC = LoopSheet.Columns("C").ColumnWidth
wD = LoopSheet.Columns("D").ColumnWidth
wF = LoopSheet.Columns("F").ColumnWidth
wG = LoopSheet.Columns("G").ColumnWidth
wH = LoopSheet.Columns("H").ColumnWidth
wI = LoopSheet.Columns("I").ColumnWidth
wJ = LoopSheet.Columns("J").ColumnWidth

wALL = wA + wB + wC + wD + wF + wG + wH + wI + wJ

AvailWidth = 121.65 - wALL

LoopSheet.Columns("E").ColumnWidth = 0.5 * AvailWidth
LoopSheet.Columns("K").ColumnWidth = 0.5 * AvailWidth
LoopSheet.Columns("K").WrapText = True

    Next i



Ground.Columns("A:K").Font.Bold = False

ca4.Columns("A:K").Font.Italic = False
ca5.Columns("A:K").Font.Italic = False
ca6.Columns("A:K").Font.Italic = False
ca7.Columns("A:K").Font.Italic = False
ca8.Columns("A:K").Font.Italic = False
ca9.Columns("A:K").Font.Italic = False

End Sub
