Attribute VB_Name = "MovePages"
Sub Transfer()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

On Error Resume Next

Dim FilePath As String
Dim FileName1 As String
Dim FileName2 As String
Dim FileName3 As String
Dim Branch As String
Dim Main As String
Dim JuvYA As String
Dim FullWB As Workbook
Dim MyDate As String

FilePath = "S:\Borrower Services\Missing Lists\Create Lists\"
Branch = "_Branch_Missing"
Main = "_Main_Missing"
JuvYA = "_JuvYA_Missing"
MyDate = Format(Now(), "mm-dd-yy")
Set FullWB = Application.ThisWorkbook


'Save the Main Missing List

FullWB.Sheets(Array("Ground", "Mezz", "Stone", "2nd Floor", "L1")).Move
    FileName1 = FilePath & MyDate & Main
    Application.ActiveWorkbook.SaveAs FileName:=FileName1
    
'Move and save the Branch Missing List

FullWB.Sheets(Array("ca4", "ca5", "ca6", "ca7", "ca8", "ca9")).Move

    FileName2 = FilePath & MyDate & Branch
    Application.ActiveWorkbook.SaveAs FileName:=FileName2

'Move and save the Juv/YA Missing List

FullWB.Sheets(Array("Juv", "YA")).Move

    FileName3 = FilePath & MyDate & JuvYA
    Application.ActiveWorkbook.SaveAs FileName:=FileName3

 
FullWB.Close False
 
End Sub
