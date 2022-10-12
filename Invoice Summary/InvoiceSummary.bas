Attribute VB_Name = "InvSum"
Sub InvoiceSummary()

'This saves memory/time by not having the screen get updated as it goes through the macro

Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic

On Error Resume Next

Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

'Find last row
Dim INVLastRow As Long

    INVLastRow = Invoice.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
'Declare variables

Dim INVinput As Range
Set INVinput = Invoice.Range("A1:A" & INVLastRow)

Dim INV2red As Range 'Red = Invoice # and Date
Dim INV2blue As Range 'Blue = Account #
Dim INV2green As Range 'Green = $Amount
Dim INV2reg As Range 'Register Date
Dim INV2regf As Range 'Register Date formula
Dim INV2vendor As Range 'Ingram or Midwest

Set INV2red = InvFormulas.Range("A1")
Set INV2blue = InvFormulas.Range("B1")
Set INV2green = InvFormulas.Range("C1")
Set INV2reg = InvFormulas.Range("D1")
Set INV2regf = InvFormulas.Range("E2")
Set INV2vendor = InvFormulas.Range("F2")


'Conditional formatting

INVinput.FormatConditions.Add Type:=xlTextString, String:="Invoice #", TextOperator:=xlBeginsWith
INVinput.FormatConditions(INVinput.FormatConditions.Count).SetFirstPriority
    
With INVinput.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 255
    .TintAndShade = 0
End With

INVinput.FormatConditions(1).StopIfTrue = False
    
INVinput.FormatConditions.Add Type:=xlTextString, String:="Vendor", TextOperator:=xlBeginsWith
INVinput.FormatConditions(INVinput.FormatConditions.Count).SetFirstPriority

With INVinput.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 12611584
    .TintAndShade = 0
End With

INVinput.FormatConditions(1).StopIfTrue = False
    
INVinput.FormatConditions.Add Type:=xlTextString, String:="Invoice Total", TextOperator:=xlBeginsWith
INVinput.FormatConditions(INVinput.FormatConditions.Count).SetFirstPriority

With INVinput.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 5287936
    .TintAndShade = 0
End With
    
INVinput.FormatConditions(1).StopIfTrue = False

'Filter on red and copy and paste to InvFormulas

INVinput.AutoFilter Field:=1, Criteria1:=RGB(255, 0, 0), Operator:=xlFilterCellColor
INVinput.SpecialCells(xlCellTypeVisible).Copy

INV2red.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Filter on blue and copy and paste to InvFormulas

INVinput.AutoFilter Field:=1, Criteria1:=RGB(0, 112, 192), Operator:=xlFilterCellColor
INVinput.SpecialCells(xlCellTypeVisible).Copy

INV2blue.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Filter on green and copy and paste to InvFormulas

INVinput.AutoFilter Field:=1, Criteria1:=RGB(0, 176, 80), Operator:=xlFilterCellColor
INVinput.SpecialCells(xlCellTypeVisible).Copy

INV2green.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Copy Invoice Register Date

INVinput.AutoFilter Field:=1, Criteria1:="=*INVOICE REGISTER*", Operator:=xlAnd
INVinput.SpecialCells(xlCellTypeVisible).Copy
INV2reg.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Invoice.AutoFilterMode = False

'Declare InvFormulas variables

Dim INV2LastRow As Long

    INV2LastRow = InvFormulas.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

Dim INV2acc1 As Range 'Account # formula 1
Dim INV2inv1 As Range 'Invoice # formula 1
Dim INV2date1 As Range 'Date formula 1
Dim INV2amount1 As Range 'Amount formula 1

Dim INV2acc2 As Range 'Account # formula 2
Dim INV2inv2 As Range 'Invoice # formula 2
Dim INV2date2 As Range 'Date formula 2
Dim INV2amount2 As Range 'Amount formula 2

Set INV2acc1 = InvFormulas.Range("J2:J" & INV2LastRow) 'Account # formula 1
Set INV2inv1 = InvFormulas.Range("K2:K" & INV2LastRow) 'Invoice # formula 1
Set INV2date1 = InvFormulas.Range("L2:L" & INV2LastRow) 'Date formula 1
Set INV2amount1 = InvFormulas.Range("M2:M" & INV2LastRow) 'Amount formula 1

Set INV2acc2 = InvFormulas.Range("N2:N" & INV2LastRow) 'Account # formula 2
Set INV2inv2 = InvFormulas.Range("O2:O" & INV2LastRow) 'Invoice # formula 2
Set INV2date2 = InvFormulas.Range("P2:P" & INV2LastRow) 'Date formula 2
Set INV2amount2 = InvFormulas.Range("Q2:Q" & INV2LastRow) 'Amount formula 2

'The MID formulas isolate the data we need from each row.
    
    INV2acc1.FormulaR1C1 = "=MID(RC[-8],SEARCH(""["",RC[-8])+1,SEARCH(""]"",RC[-8])-SEARCH(""["",RC[-8])-1)" 'J
    INV2inv1.FormulaR1C1 = "=MID(RC[-10],SEARCH(""INVOICE #"",RC[-10])+10,8)" 'K
    INV2date1.FormulaR1C1 = "=MID(RC[-11],SEARCH(""INVOICE DATE: "",RC[-11])+14,10)" 'L
    INV2amount1.FormulaR1C1 = "=MID(RC[-10],SEARCH(""$"",RC[-10]),SEARCH("" "",RC[-10],SEARCH(""$"",RC[-10]))-SEARCH(""$"",RC[-10]))" 'M
    
    INV2acc2.FormulaR1C1 = "=IFERROR(RC[-4]*1,RC[-4])" 'N
    INV2inv2.FormulaR1C1 = "=IFERROR(RC[-4]*1,RC[-4])" 'O
    INV2date2.FormulaR1C1 = "=IFERROR(RC[-4]*1,RC[-4])" 'P
    INV2amount2.FormulaR1C1 = "=RC[-4]*1" 'Q
    
    INV2regf.FormulaR1C1 = "=MID(RC[-1],SEARCH(""-"",RC[-1])-2,8)"
    INV2vendor.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(""Ing"",RC[-2])),""Ingram"",IF(ISNUMBER(SEARCH(""Midwest"",RC[-2])),RIGHT(RC[-2],LEN(RC[-2])-SEARCH("" Midwest"",RC[-2])),""?""))"


'Copy to final page

Dim INVcopy As Range
Set INVcopy = InvFormulas.Range("N2:Q" & INV2LastRow)

INVcopy.Copy

InvFinal.Range("A3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Dim INVfinalLastRow As Long

    INVfinalLastRow = InvFinal.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

Dim INVNo1 As Range
Dim INVNo2 As Range

Set INVNo1 = InvFinal.Range("B3")
Set INVNo2 = InvFinal.Range("B" & INVfinalLastRow)

Dim INVsum As Range
Set INVsum = InvFinal.Range("D" & INVfinalLastRow + 1)

INVsum.FormulaR1C1 = "=SUM(R1C:R[-1]C)"
With INVsum.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 4.99893185216834E-02
        .PatternTintAndShade = 0
End With

With INVsum.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
End With
    
    
With InvFinal
    .Range("A1").FormulaR1C1 = "Cambridge Public Library"
    .Range("A2").FormulaR1C1 = "Account #"
    .Range("B2").FormulaR1C1 = "Invoice #"
    .Range("C2").FormulaR1C1 = "Date"
    .Range("D2").FormulaR1C1 = "$ Amount"
    .Range("D1").FormulaR1C1 = INV2vendor
    .Range("D1").HorizontalAlignment = xlRight
    .Range("D1").VerticalAlignment = xlTop
    .Range("D1").Font.Bold = True
    .Columns("C:C").NumberFormat = "m/d/yyyy"
    .Columns("D:D").NumberFormat = "$#,##0.00"
    .Columns("A:A").ColumnWidth = 12.5
    .Columns("B:D").EntireColumn.AutoFit
    .Range("A1").WrapText = True
    .Range("A2:D2").Font.Bold = True
End With


With InvFinal.Range("A2:D" & INVfinalLastRow).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With
    
With InvFinal.Range("A2:D" & INVfinalLastRow).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With
    
With InvFinal.Range("A2:D" & INVfinalLastRow).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With
    
With InvFinal.Range("A2:D" & INVfinalLastRow).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With
    
With InvFinal.Range("A2:D" & INVfinalLastRow).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With

With InvFinal.Range("A2:D" & INVfinalLastRow).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With


InvFinal.Columns("A:A").HorizontalAlignment = xlLeft
InvFinal.Range("A1").Font.Bold = True

With InvFinal.Range("A2:D2").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
End With

InvFinal.Range("A2:D2").Borders(xlDiagonalDown).LineStyle = xlNone
InvFinal.Range("A2:D2").Borders(xlDiagonalUp).LineStyle = xlNone

With InvFinal.Range("A2:D2").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
End With

With InvFinal.Range("A2:D2").Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
End With

With InvFinal.Range("A2:D2").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
End With
    
With InvFinal.Range("A2:D2").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
End With
    
With InvFinal.Range("A2:D2").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With

InvFinal.Range("A2:D2").Borders(xlInsideHorizontal).LineStyle = xlNone
InvFinal.PageSetup.PrintTitleRows = "$2:$2"

Invoice.Activate

'Save to new book

Dim OriginalWB As Workbook
Set OriginalWB = Application.ThisWorkbook

Dim FilePath As String
Dim InvPath As String
Dim VendorPath As String
Dim MyDate As String
Dim DateCreated As String


    FilePath = "S:\Collection Development\Invoice Summaries\" 'Change to suit
    'InvPath = "_InvSummary_"
    VendorPath = INV2vendor
    MyDate = Format(INV2regf, "mm-dd-yy")
    Sum = Int(INVsum)
    'DateCreated = Format(Now(), "yyyymmddhhmmss")

    FileName1 = FilePath & VendorPath & "_" & MyDate & "_" & INVNo1 & "-" & INVNo2 & "_" & Sum


'Save copy
     
OriginalWB.Sheets("InvFinal").Visible = True
OriginalWB.Sheets("InvFinal").Copy
ActiveWorkbook.Sheets("InvFinal").Range("A1").Select

    Application.ActiveWorkbook.SaveAs Filename:=FileName1
    
OriginalWB.Sheets("InvFinal").Visible = False

End Sub
 
