Attribute VB_Name = "Item"
Option Explicit


'This is the macro for setting up the Item Paging.
'It's a lot less work than the Title Paging, so it's nice and short by comparison.
Sub ItemPaging()

'Turn off screen animations so this runs faster.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

Dim IStart As Worksheet, IList As Worksheet
Dim ISR As Range, ILast4 As Range, ICallNos As Range, ITitle As Range
Dim ICode As Range, IListRange As Range
Dim i As Integer

Set IStart = Sheets("Item Paging")
Set IList = Sheets("Item List")

Set ISR = IStart.Columns("A:A")
Set ILast4 = IList.Columns("C:C")
Set ICallNos = IList.Columns("D:D")
Set ITitle = IList.Columns("E:E")
Set ICode = IList.Columns("F:F")
Set IListRange = IList.Range("C2:F1000")

ISR.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
ISR.HorizontalAlignment = xlGeneral

'This moves everything from the email into the "Item List" sheet.
i = 2
Do While IStart.Cells(i, 1).Value2 <> ""
If InStr(1, IStart.Cells(i, 1).Value2, "      TITLE: ") > 0 Then
IList.Cells(i, 5).Value2 = IStart.Cells(i, 1).Value2 'Title
IList.Cells(i, 4).Value2 = IStart.Cells(i, 1).Offset(1, 0).Value2 ' Call No
IList.Cells(i, 6).Value2 = IStart.Cells(i, 1).Offset(2, 0).Value2 'Barcode
IList.Cells(i, 8).Value2 = IStart.Cells(i, 1).Offset(4, 0).Value2 'New status
IList.Cells(i, 3).Value2 = Right(IStart.Cells(i, 1).Offset(2, 0).Value2, 4) 'Last4
End If

i = i + 1
Loop

'This deletes any blank rows.
IList.Range("D2:D20000").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

'This bolds all the new items.
i = 2
Do While IList.Cells(i, 3).Value2 <> ""
If InStr(1, IList.Cells(i, 8).Value2, "CAMBRIDGE/New") > 0 Then
IList.Cells(i, 3).Font.Bold = True
IList.Cells(i, 4).Font.Bold = True
IList.Cells(i, 5).Font.Bold = True
IList.Cells(i, 6).Font.Bold = True
End If
i = i + 1
Loop
IList.Columns(8).Delete



'This cleans up the cells.
With IListRange
.Replace "      BARCODE:  ", vbNullString, xlPart
.Replace "      CALL NO:  ", vbNullString, xlPart
.Replace "      TITLE:    ", vbNullString, xlPart
.Replace "      PICKUP AT:  ", vbNullString, xlPart
End With

'This trims some Call #'s.
With ICallNos
.Replace "[Home & Health]", "[H&H]", xlPart
.Replace "CD CLASSICAL", "CD CLASS", xlPart
.Replace "CD ROCK", "CD POP", xlPart
.Replace "CD FOLK", "CD POP", xlPart
.Replace "CD SNDTRK", "CD POP", xlPart
.Replace "CD COUNTRY", "CD POP", xlPart
.Replace "CD GENERAL", "CD POP", xlPart
.Replace "CD POPULAR", "CD POP", xlPart
.Replace "FICTION", "FIC", xlPart
.Replace "CDB Mystery", "CDB FIC", xlPart
.Replace "CDB SCI FIC", "CDB FIC", xlPart
.Replace "CDB FIC SHORT STORIES", "CDB FIC SHORT", xlPart
.Replace "LP SHORT STORIES", "LP SHORT", xlPart
.Replace "[Great Courses]", "[G C]", xlPart
.Replace "MP3 ", "CDB (MP3) ", xlPart
.Replace "FICTION", "FIC", xlPart
.Replace "SCI FIC", "SCIFI", xlPart
.Replace "MYSTERY", "MYST", xlPart
.Replace "POETRY", "POET", xlPart
.Replace "[Business]", "[Biz]", xlPart
End With


'The below sets up the headers and other formatting.
IList.Activate
Call Base.Headers

With IList
.PageSetup.LeftHeader = "&A" & Chr(10) & "Bold records are for New items"
End With


'This adds the item paging stats into the Paging Stats sheet.
    Dim Paging As Workbook, Stats As Workbook
    Dim StatsSheet As Worksheet
    Dim CompCount As Range, StatsDate As Range, StatsLast As Range
    Dim Total As Integer

    Set Paging = ThisWorkbook
    Total = IList.Range("C500").End(xlUp).Row - 1

    Set Stats = Workbooks.Open _
    ("\\coc\Library\Borrower Services\Paging List\Paging Stats.xlsm")

    Set StatsSheet = Stats.Sheets("Stats")
    'Find last column in Stats to enter the information.
    Set StatsDate = StatsSheet.Cells(StatsSheet.Rows.Count, "A").End(xlUp).Offset(1, 0)
    Set StatsLast = StatsSheet.Cells(StatsSheet.Rows.Count, "B").End(xlUp).Offset(1, 0)

    StatsDate.Value2 = Date
    StatsLast.Value2 = Total
    StatsLast.Offset(0, 1).Value2 = "Item Paging"

    Stats.Save
    Stats.Close False
    
    
IList.Visible = xlSheetVisible
IStart.Visible = xlSheetHidden


End Sub
