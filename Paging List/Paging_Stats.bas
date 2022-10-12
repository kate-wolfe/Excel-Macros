Attribute VB_Name = "Stats"
Option Explicit

Sub OpenStats()

'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
'These are the variables we need to move the stats.
    Dim Paging As Workbook, Stats As Workbook
    Dim ItemC As Worksheet, Complete As Worksheet, StatsSheet As Worksheet, BarSheet As Worksheet
    Dim CompCount As Range, StatsDate As Range, StatsLast As Range, IListCount As Range
    Dim CompNum As Integer, ItemNum As Integer, Total As Integer
    Dim DayCount As Integer
    Dim FolderPath As String
 'These are the ones we need to compare the barcodes.
    Dim CompF As Range, CompI As Range
    Dim BarA As Range, BarC As Range, BarD As Range
    Dim CompareRange As Variant, NewRange As Variant, x As Variant, y As Variant, z As Variant
'Setting Ranges to shorten code
    Set Paging = ThisWorkbook
    Set Complete = ThisWorkbook.Sheets("Complete")
    Set ItemC = ThisWorkbook.Sheets("Item List")

'This finds the last row in Complete with content and subtracts 1 from it to get the total number of items on the list.
'(You subtract 1 due to the header taking up a row).
    Total = Complete.Range("C3000").End(xlUp).Offset(-1, 0).Row - 1
    
    
    FolderPath = Application.ActiveWorkbook.Path & "\Paging Stats.xlsm"
    Set Stats = Workbooks.Open(FolderPath)
    
    Set StatsSheet = Stats.Sheets("Stats")
    Set BarSheet = Stats.Sheets("Barcodes")
    
    
'Find last column in Stats to enter the information.
    Set StatsDate = StatsSheet.Cells(StatsSheet.Rows.Count, "A").End(xlUp).Offset(1, 0)
    Set StatsLast = StatsSheet.Cells(StatsSheet.Rows.Count, "B").End(xlUp).Offset(1, 0)
    
'Set things up to check for duplicates.
Set CompF = Complete.Range("F2:F2000")
Set CompI = Complete.Range("I2:I2000")
Set BarA = BarSheet.Columns("A:A")
Set BarC = BarSheet.Range("C2:C2000")
Set BarD = BarSheet.Range("D2:D2000")
Set CompareRange = BarSheet.Range("A2:C2000")
DayCount = 0


'Add info to Stats.
    StatsDate.Value2 = Date 'Today's dae goes in column A
    StatsLast.Value2 = Total 'The length of the Paging List goes in column B
    
'The below commented-out code highlights items that have appeared on the paging list on the last few days.
'I turned it off when the paging list got enormous during Fall 2020. It may or may not work anymore.

''This next part only happens if this is the first time Paging has been run today.
'If BarSheet.Cells(1, 2).Value2 = Date Then
'DayCount = 1
'End If
'
''Move info from Comp to Stats
'BarC.Value2 = CompF.Value2
'
''Compare the new barcodes to the old.
'
'Dim myCell As Range
'For Each myCell In CompareRange
'    If WorksheetFunction.CountIf(CompareRange, myCell.Value) > 1 Then
'    myCell.Interior.Color = RGB(176, 224, 230)
'    End If
'Next
''
''CompH.Value2 = BarC.Value2
''CompH.Interior.Color = BarC.Interior.Color
'BarSheet.Cells(1, 3).Value2 = Date
'
'
''Delete oldest barcodes
'BarA.Delete Shift:=xlToLeft
'
'BarC.Copy CompF
'
''If DayCount = 0 Then
    Stats.Save
'    'End If
'
    Stats.Close False

Complete.Activate
  
End Sub



