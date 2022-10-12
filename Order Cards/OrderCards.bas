Attribute VB_Name = "Module2"
Sub NewOrder()
Attribute NewOrder.VB_ProcData.VB_Invoke_Func = " \n14"

'Disabling DisplayAlerts means that the "Do you want to save?" box won't pop up.
'Disabling ScreenUpdating saves memory

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Workbooks.Open "H:\My Documents\ingram.xls"
Dim IngramWB As Workbook
Set IngramWB = Application.Workbooks("ingram.xls")

Dim IngramSH As Worksheet
Set IngramSH = IngramWB.Sheets("Titles")

Dim RecordCount As Long

IngramSH.Columns("I:O").Delete Shift:=xlToLeft
IngramSH.Columns("F:F").Delete Shift:=xlToLeft
IngramSH.Columns("A:A").Delete Shift:=xlToLeft
IngramSH.Rows("1:1").ClearContents
    
    
    RecordCount = IngramSH.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
                    
IngramSH.Range("A1") = RecordCount

Dim PageCount As Long
PageCount = Application.RoundUp(RecordCount / 3, 0)

IngramSH.Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

IngramSH.Range("K2") = 1
IngramSH.Range("K2").DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Stop:=PageCount, Trend:=False
        
IngramSH.Range("K2").Offset(PageCount) = 1
IngramSH.Range("K2").Offset(PageCount).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Stop:=PageCount, Trend:=False
        
IngramSH.Range("K2").Offset(PageCount * 2) = 1
IngramSH.Range("K2").Offset(PageCount * 2).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Stop:=PageCount, Trend:=False
        
IngramSH.Range("L2:L" & PageCount + 1).Value = "A"
IngramSH.Range("L2:L" & PageCount + 1).Offset(PageCount).Value = "B"
IngramSH.Range("L2:L" & PageCount + 1).Offset(PageCount * 2).Value = "C"

'Date format for PubDate - column G $ format for Price - column H

    IngramSH.Columns("F:F").NumberFormat = "yyyymmdd           mm/dd/yy"
    IngramSH.Columns("G:G").NumberFormat = "$#,##0.00"
    IngramSH.Columns("J:J").NumberFormat = "$#,##0.00"

'Delete extra rows on top'

    IngramSH.Rows("1:2").ClearContents
    IngramSH.Rows("1:1").Delete Shift:=xlUp
    IngramSH.Range("A1") = "ISBN13"
    IngramSH.Range("B1") = "Title"
    IngramSH.Range("C1") = "Author"
    IngramSH.Range("D1") = "Type"
    IngramSH.Range("E1") = "Publisher"
    IngramSH.Range("F1") = "PubDate"
    IngramSH.Range("G1") = "FullPrice"
    IngramSH.Range("H1") = "Quantity"
    IngramSH.Range("I1") = "Notes"
    IngramSH.Range("J1") = "DiscPrice"
    IngramSH.Range("K1") = "Number"
    IngramSH.Range("L1") = "Section"


'Make sure ISBN is in correct number format

    IngramSH.Columns("A:A").NumberFormat = "0"

'Save and close'

Dim FileName As String
Dim FilePath As String

FilePath = "H:\My Documents\"
FileName = FilePath & "ingram.xls"

IngramWB.SaveAs FileName:=FileName, FileFormat:=xlExcel8
IngramWB.Close

End Sub
