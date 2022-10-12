Attribute VB_Name = "Format"
Option Explicit

Sub Formatting()

Dim sht As Worksheet

For Each sht In ThisWorkbook.Worksheets
    sht.Range("A:I").EntireColumn.AutoFit
  Next sht
  


End Sub
