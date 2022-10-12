Attribute VB_Name = "GetACQ"
Option Explicit

Sub GetAcqs()

Application.ScreenUpdating = False

'Open New Items Report and declare variables

Dim COMPsh As Worksheet
Set COMPsh = ThisWorkbook.ActiveSheet

Dim fd As Office.FileDialog
Dim strFile As String
 
Set fd = Application.FileDialog(msoFileDialogFilePicker)
 
With fd
 
    .Filters.Clear
    .Filters.Add "Excel Files", "*.xlsx?", 1
    .Title = "Choose an Excel file"
    .AllowMultiSelect = False
    .InitialFileName = "H:\Desktop\"
 
    If .Show = True Then
        strFile = .SelectedItems(1)
    End If
 
End With

Dim ACQwb As Workbook
Set ACQwb = Workbooks.Open(strFile)

'Count stuff

Dim Location As Range
Dim SCAT As Range

Set Location = ACQwb.Worksheets("Sheet1").Range("A:A")
Set SCAT = ACQwb.Worksheets("Sheet1").Range("D:D")

'Boudreau

Dim BoudAll As Integer
Dim BoudAdPrint As Integer
Dim BoudJuvPrint1 As Integer
Dim BoudJuvPrint2 As Integer
Dim BoudJuvPrint3 As Integer
Dim BoudJuvPrint4 As Integer
Dim BoudJuvPrintTotal As Integer
Dim BoudAdNon As Integer
Dim BoudJuvNon1 As Integer
Dim BoudJuvNon2 As Integer
Dim BoudJuvNon3 As Integer
Dim BoudJuvNon4 As Integer
Dim BoudJuvNon5 As Integer
Dim BoudJuvNon6 As Integer
Dim BoudJuvNonTotal As Integer

BoudAll = WorksheetFunction.CountIf(Location, "ca4*")
BoudAdPrint = WorksheetFunction.CountIf(Location, "ca4a*")
BoudJuvPrint1 = WorksheetFunction.CountIfs(Location, "ca4y*", SCAT, ">=210", SCAT, "<=219")
BoudJuvPrint2 = WorksheetFunction.CountIfs(Location, "ca4j*", SCAT, ">=228", SCAT, "<=234")
BoudJuvPrint3 = WorksheetFunction.CountIfs(Location, "ca4j*", SCAT, ">=236", SCAT, "<=241")
BoudJuvPrint4 = WorksheetFunction.CountIfs(Location, "ca4j*", SCAT, ">=250", SCAT, "<=268")
BoudJuvPrintTotal = BoudJuvPrint1 + BoudJuvPrint2 + BoudJuvPrint3 + BoudJuvPrint4
BoudAdNon = WorksheetFunction.CountIf(Location, "ca4n*")
BoudJuvNon1 = WorksheetFunction.CountIfs(Location, "ca4j*", SCAT, "208")
BoudJuvNon2 = WorksheetFunction.CountIfs(Location, "cayj*", SCAT, "220")
BoudJuvNon3 = WorksheetFunction.CountIfs(Location, "ca4y*", SCAT, ">=223", SCAT, "<=224")
BoudJuvNon4 = WorksheetFunction.CountIfs(Location, "ca4y*", SCAT, ">=226", SCAT, "<=227")
BoudJuvNon5 = WorksheetFunction.CountIfs(Location, "ca4j*", SCAT, ">=242", SCAT, "<=244")
BoudJuvNon6 = WorksheetFunction.CountIfs(Location, "ca4j*", SCAT, ">=246", SCAT, "<=249")
BoudJuvNonTotal = BoudJuvNon1 + BoudJuvNon2 + BoudJuvNon3 + BoudJuvNon4 + BoudJuvNon5 + BoudJuvNon6

If BoudAll <> BoudAdPrint + BoudJuvPrintTotal + BoudAdNon + BoudJuvNonTotal Then
    MsgBox ("Check Boudreau Juv/YA Acquisitions")
End If

COMPsh.Range("B8") = BoudAdPrint
COMPsh.Range("B9") = BoudJuvPrintTotal
COMPsh.Range("B13") = BoudAdNon
COMPsh.Range("B14") = BoudJuvNonTotal

'CSQ

Dim CSQAll As Integer
Dim CSQAdPrint As Integer
Dim CSQJuvPrint1 As Integer
Dim CSQJuvPrint2 As Integer
Dim CSQJuvPrint3 As Integer
Dim CSQJuvPrint4 As Integer
Dim CSQJuvPrintTotal As Integer
Dim CSQAdNon As Integer
Dim CSQJuvNon1 As Integer
Dim CSQJuvNon2 As Integer
Dim CSQJuvNon3 As Integer
Dim CSQJuvNon4 As Integer
Dim CSQJuvNon5 As Integer
Dim CSQJuvNon6 As Integer
Dim CSQJuvNonTotal As Integer

CSQAll = WorksheetFunction.CountIf(Location, "ca5*")
CSQAdPrint = WorksheetFunction.CountIf(Location, "ca5a*")
CSQJuvPrint1 = WorksheetFunction.CountIfs(Location, "ca5y*", SCAT, ">=210", SCAT, "<=219")
CSQJuvPrint2 = WorksheetFunction.CountIfs(Location, "ca5j*", SCAT, ">=228", SCAT, "<=234")
CSQJuvPrint3 = WorksheetFunction.CountIfs(Location, "ca5j*", SCAT, ">=236", SCAT, "<=241")
CSQJuvPrint4 = WorksheetFunction.CountIfs(Location, "ca5j*", SCAT, ">=250", SCAT, "<=268")
CSQJuvPrintTotal = CSQJuvPrint1 + CSQJuvPrint2 + CSQJuvPrint3 + CSQJuvPrint4
CSQAdNon = WorksheetFunction.CountIf(Location, "ca5n*")
CSQJuvNon1 = WorksheetFunction.CountIfs(Location, "ca5j*", SCAT, "208")
CSQJuvNon2 = WorksheetFunction.CountIfs(Location, "ca5j*", SCAT, "220")
CSQJuvNon3 = WorksheetFunction.CountIfs(Location, "ca5y*", SCAT, ">=223", SCAT, "<=224")
CSQJuvNon4 = WorksheetFunction.CountIfs(Location, "ca5y*", SCAT, ">=226", SCAT, "<=227")
CSQJuvNon5 = WorksheetFunction.CountIfs(Location, "ca5j*", SCAT, ">=242", SCAT, "<=244")
CSQJuvNon6 = WorksheetFunction.CountIfs(Location, "ca5j*", SCAT, ">=246", SCAT, "<=249")
CSQJuvNonTotal = CSQJuvNon1 + CSQJuvNon2 + CSQJuvNon3 + CSQJuvNon4 + CSQJuvNon5 + CSQJuvNon6

If CSQAll <> CSQAdPrint + CSQJuvPrintTotal + CSQAdNon + CSQJuvNonTotal Then
    MsgBox ("Check CSQ Juv/YA Acquisitions")
End If

COMPsh.Range("D8") = CSQAdPrint
COMPsh.Range("D9") = CSQJuvPrintTotal
COMPsh.Range("D13") = CSQAdNon
COMPsh.Range("D14") = CSQJuvNonTotal

'Collins

Dim CollinsAll As Integer
Dim CollinsAdPrint As Integer
Dim CollinsJuvPrint1 As Integer
Dim CollinsJuvPrint2 As Integer
Dim CollinsJuvPrint3 As Integer
Dim CollinsJuvPrint4 As Integer
Dim CollinsJuvPrintTotal As Integer
Dim CollinsAdNon As Integer
Dim CollinsJuvNon1 As Integer
Dim CollinsJuvNon2 As Integer
Dim CollinsJuvNon3 As Integer
Dim CollinsJuvNon4 As Integer
Dim CollinsJuvNon5 As Integer
Dim CollinsJuvNon6 As Integer
Dim CollinsJuvNonTotal As Integer

CollinsAll = WorksheetFunction.CountIf(Location, "ca6*")
CollinsAdPrint = WorksheetFunction.CountIf(Location, "ca6a*")
CollinsJuvPrint1 = WorksheetFunction.CountIfs(Location, "ca6y*", SCAT, ">=210", SCAT, "<=219")
CollinsJuvPrint2 = WorksheetFunction.CountIfs(Location, "ca6j*", SCAT, ">=228", SCAT, "<=234")
CollinsJuvPrint3 = WorksheetFunction.CountIfs(Location, "ca6j*", SCAT, ">=236", SCAT, "<=241")
CollinsJuvPrint4 = WorksheetFunction.CountIfs(Location, "ca6j*", SCAT, ">=250", SCAT, "<=268")
CollinsJuvPrintTotal = CollinsJuvPrint1 + CollinsJuvPrint2 + CollinsJuvPrint3 + CollinsJuvPrint4
CollinsAdNon = WorksheetFunction.CountIf(Location, "ca6n*")
CollinsJuvNon1 = WorksheetFunction.CountIfs(Location, "ca6j*", SCAT, "208")
CollinsJuvNon2 = WorksheetFunction.CountIfs(Location, "ca6j*", SCAT, "220")
CollinsJuvNon3 = WorksheetFunction.CountIfs(Location, "ca6y*", SCAT, ">=223", SCAT, "<=224")
CollinsJuvNon4 = WorksheetFunction.CountIfs(Location, "ca6y*", SCAT, ">=226", SCAT, "<=227")
CollinsJuvNon5 = WorksheetFunction.CountIfs(Location, "ca6j*", SCAT, ">=242", SCAT, "<=244")
CollinsJuvNon6 = WorksheetFunction.CountIfs(Location, "ca6j*", SCAT, ">=246", SCAT, "<=249")
CollinsJuvNonTotal = CollinsJuvNon1 + CollinsJuvNon2 + CollinsJuvNon3 + CollinsJuvNon4 + CollinsJuvNon5 + CollinsJuvNon6

If CollinsAll <> CollinsAdPrint + CollinsJuvPrintTotal + CollinsAdNon + CollinsJuvNonTotal Then
    MsgBox ("Check Collins Juv/YA Acquisitions")
End If

COMPsh.Range("F8") = CollinsAdPrint
COMPsh.Range("F9") = CollinsJuvPrintTotal
COMPsh.Range("F13") = CollinsAdNon
COMPsh.Range("F14") = CollinsJuvNonTotal

'OConnell

Dim OConnAll As Integer
Dim OConnAdPrint As Integer
Dim OConnJuvPrint1 As Integer
Dim OConnJuvPrint2 As Integer
Dim OConnJuvPrint3 As Integer
Dim OConnJuvPrint4 As Integer
Dim OConnJuvPrintTotal As Integer
Dim OConnAdNon As Integer
Dim OConnJuvNon1 As Integer
Dim OConnJuvNon2 As Integer
Dim OConnJuvNon3 As Integer
Dim OConnJuvNon4 As Integer
Dim OConnJuvNon5 As Integer
Dim OConnJuvNon6 As Integer
Dim OConnJuvNonTotal As Integer

OConnAll = WorksheetFunction.CountIf(Location, "ca7*")
OConnAdPrint = WorksheetFunction.CountIf(Location, "ca7a*")
OConnJuvPrint1 = WorksheetFunction.CountIfs(Location, "ca7y*", SCAT, ">=210", SCAT, "<=219")
OConnJuvPrint2 = WorksheetFunction.CountIfs(Location, "ca7j*", SCAT, ">=228", SCAT, "<=234")
OConnJuvPrint3 = WorksheetFunction.CountIfs(Location, "ca7j*", SCAT, ">=236", SCAT, "<=241")
OConnJuvPrint4 = WorksheetFunction.CountIfs(Location, "ca7j*", SCAT, ">=250", SCAT, "<=268")
OConnJuvPrintTotal = OConnJuvPrint1 + OConnJuvPrint2 + OConnJuvPrint3 + OConnJuvPrint4
OConnAdNon = WorksheetFunction.CountIf(Location, "ca7n*")
OConnJuvNon1 = WorksheetFunction.CountIfs(Location, "ca7j*", SCAT, "208")
OConnJuvNon2 = WorksheetFunction.CountIfs(Location, "ca7j*", SCAT, "220")
OConnJuvNon3 = WorksheetFunction.CountIfs(Location, "ca7y*", SCAT, ">=223", SCAT, "<=224")
OConnJuvNon4 = WorksheetFunction.CountIfs(Location, "ca7y*", SCAT, ">=226", SCAT, "<=227")
OConnJuvNon5 = WorksheetFunction.CountIfs(Location, "ca7j*", SCAT, ">=242", SCAT, "<=244")
OConnJuvNon6 = WorksheetFunction.CountIfs(Location, "ca7j*", SCAT, ">=246", SCAT, "<=249")
OConnJuvNonTotal = OConnJuvNon1 + OConnJuvNon2 + OConnJuvNon3 + OConnJuvNon4 + OConnJuvNon5 + OConnJuvNon6

If OConnAll <> OConnAdPrint + OConnJuvPrintTotal + OConnAdNon + OConnJuvNonTotal Then
    MsgBox ("Check OConnell Juv/YA Acquisitions")
End If

COMPsh.Range("J8") = OConnAdPrint
COMPsh.Range("J9") = OConnJuvPrintTotal
COMPsh.Range("J13") = OConnAdNon
COMPsh.Range("J14") = OConnJuvNonTotal

'ONeill

Dim ONeillAll As Integer
Dim ONeillAdPrint As Integer
Dim ONeillJuvPrint1 As Integer
Dim ONeillJuvPrint2 As Integer
Dim ONeillJuvPrint3 As Integer
Dim ONeillJuvPrint4 As Integer
Dim ONeillJuvPrintTotal As Integer
Dim ONeillAdNon As Integer
Dim ONeillJuvNon1 As Integer
Dim ONeillJuvNon2 As Integer
Dim ONeillJuvNon3 As Integer
Dim ONeillJuvNon4 As Integer
Dim ONeillJuvNon5 As Integer
Dim ONeillJuvNon6 As Integer
Dim ONeillJuvNonTotal As Integer

ONeillAll = WorksheetFunction.CountIf(Location, "ca8*")
ONeillAdPrint = WorksheetFunction.CountIf(Location, "ca8a*")
ONeillJuvPrint1 = WorksheetFunction.CountIfs(Location, "ca8y*", SCAT, ">=210", SCAT, "<=219")
ONeillJuvPrint2 = WorksheetFunction.CountIfs(Location, "ca8j*", SCAT, ">=228", SCAT, "<=234")
ONeillJuvPrint3 = WorksheetFunction.CountIfs(Location, "ca8j*", SCAT, ">=236", SCAT, "<=241")
ONeillJuvPrint4 = WorksheetFunction.CountIfs(Location, "ca8j*", SCAT, ">=250", SCAT, "<=268")
ONeillJuvPrintTotal = ONeillJuvPrint1 + ONeillJuvPrint2 + ONeillJuvPrint3 + ONeillJuvPrint4
ONeillAdNon = WorksheetFunction.CountIf(Location, "ca8n*")
ONeillJuvNon1 = WorksheetFunction.CountIfs(Location, "ca8j*", SCAT, "208")
ONeillJuvNon2 = WorksheetFunction.CountIfs(Location, "ca8j*", SCAT, "220")
ONeillJuvNon3 = WorksheetFunction.CountIfs(Location, "ca8y*", SCAT, ">=223", SCAT, "<=224")
ONeillJuvNon4 = WorksheetFunction.CountIfs(Location, "ca8y*", SCAT, ">=226", SCAT, "<=227")
ONeillJuvNon5 = WorksheetFunction.CountIfs(Location, "ca8j*", SCAT, ">=242", SCAT, "<=244")
ONeillJuvNon6 = WorksheetFunction.CountIfs(Location, "ca8j*", SCAT, ">=246", SCAT, "<=249")
ONeillJuvNonTotal = ONeillJuvNon1 + ONeillJuvNon2 + ONeillJuvNon3 + ONeillJuvNon4 + ONeillJuvNon5 + ONeillJuvNon6

If ONeillAll <> ONeillAdPrint + ONeillJuvPrintTotal + ONeillAdNon + ONeillJuvNonTotal Then
    MsgBox ("Check ONeill Juv/YA Acquisitions")
End If

COMPsh.Range("L8") = ONeillAdPrint
COMPsh.Range("L9") = ONeillJuvPrintTotal
COMPsh.Range("L13") = ONeillAdNon
COMPsh.Range("L14") = ONeillJuvNonTotal

'Valente

Dim ValAll As Integer
Dim ValAdPrint As Integer
Dim ValJuvPrint1 As Integer
Dim ValJuvPrint2 As Integer
Dim ValJuvPrint3 As Integer
Dim ValJuvPrint4 As Integer
Dim ValJuvPrintTotal As Integer
Dim ValAdNon As Integer
Dim ValJuvNon1 As Integer
Dim ValJuvNon2 As Integer
Dim ValJuvNon3 As Integer
Dim ValJuvNon4 As Integer
Dim ValJuvNon5 As Integer
Dim ValJuvNon6 As Integer
Dim ValJuvNonTotal As Integer

ValAll = WorksheetFunction.CountIf(Location, "ca9*")
ValAdPrint = WorksheetFunction.CountIf(Location, "ca9a*")
ValJuvPrint1 = WorksheetFunction.CountIfs(Location, "ca9y*", SCAT, ">=210", SCAT, "<=219")
ValJuvPrint2 = WorksheetFunction.CountIfs(Location, "ca9j*", SCAT, ">=228", SCAT, "<=234")
ValJuvPrint3 = WorksheetFunction.CountIfs(Location, "ca9j*", SCAT, ">=236", SCAT, "<=241")
ValJuvPrint4 = WorksheetFunction.CountIfs(Location, "ca9j*", SCAT, ">=250", SCAT, "<=268")
ValJuvPrintTotal = ValJuvPrint1 + ValJuvPrint2 + ValJuvPrint3 + ValJuvPrint4
ValAdNon = WorksheetFunction.CountIf(Location, "ca9n*")
ValJuvNon1 = WorksheetFunction.CountIfs(Location, "ca9j*", SCAT, "208")
ValJuvNon2 = WorksheetFunction.CountIfs(Location, "ca9j*", SCAT, "220")
ValJuvNon3 = WorksheetFunction.CountIfs(Location, "ca9y*", SCAT, ">=223", SCAT, "<=224")
ValJuvNon4 = WorksheetFunction.CountIfs(Location, "ca9y*", SCAT, ">=226", SCAT, "<=227")
ValJuvNon5 = WorksheetFunction.CountIfs(Location, "ca9j*", SCAT, ">=242", SCAT, "<=244")
ValJuvNon6 = WorksheetFunction.CountIfs(Location, "ca9j*", SCAT, ">=246", SCAT, "<=249")
ValJuvNonTotal = ValJuvNon1 + ValJuvNon2 + ValJuvNon3 + ValJuvNon4 + ValJuvNon5 + ValJuvNon6

If ValAll <> ValAdPrint + ValJuvPrintTotal + ValAdNon + ValJuvNonTotal Then
    MsgBox ("Check Valente Juv/YA Acquisitions")
End If

COMPsh.Range("N8") = ValAdPrint
COMPsh.Range("N9") = ValJuvPrintTotal
COMPsh.Range("N13") = ValAdNon
COMPsh.Range("N14") = ValJuvNonTotal

'Main

Dim MainAll As Integer
Dim MainAdPrint As Integer
Dim SRSPrint As Integer
Dim CRPrint As Integer
Dim RefPrint As Integer
Dim MainJuvPrint1 As Integer
Dim MainJuvPrint2 As Integer
Dim MainJuvPrint3 As Integer
Dim MainJuvPrint4 As Integer
Dim MainJuvPrintTotal As Integer
Dim MainAdNon As Integer
Dim MainJuvNon1 As Integer
Dim MainJuvNon2 As Integer
Dim MainJuvNon3 As Integer
Dim MainJuvNon4 As Integer
Dim MainJuvNon5 As Integer
Dim MainJuvNon6 As Integer
Dim MainJuvNonTotal As Integer

MainAll = WorksheetFunction.CountIf(Location, "cam*")
MainAdPrint = WorksheetFunction.CountIf(Location, "cama*")
SRSPrint = WorksheetFunction.CountIf(Location, "ca3a*")
CRPrint = WorksheetFunction.CountIf(Location, "camc*")
RefPrint = WorksheetFunction.CountIf(Location, "camr*")
MainJuvPrint1 = WorksheetFunction.CountIfs(Location, "camy*", SCAT, ">=210", SCAT, "<=219")
MainJuvPrint2 = WorksheetFunction.CountIfs(Location, "camj*", SCAT, ">=228", SCAT, "<=234")
MainJuvPrint3 = WorksheetFunction.CountIfs(Location, "camj*", SCAT, ">=236", SCAT, "<=241")
MainJuvPrint4 = WorksheetFunction.CountIfs(Location, "camj*", SCAT, ">=250", SCAT, "<=268")
MainJuvPrintTotal = MainJuvPrint1 + MainJuvPrint2 + MainJuvPrint3 + MainJuvPrint4
MainAdNon = WorksheetFunction.CountIf(Location, "camn*")
MainJuvNon1 = WorksheetFunction.CountIfs(Location, "camj*", SCAT, "208")
MainJuvNon2 = WorksheetFunction.CountIfs(Location, "camj*", SCAT, "220")
MainJuvNon3 = WorksheetFunction.CountIfs(Location, "camy*", SCAT, ">=223", SCAT, "<=224")
MainJuvNon4 = WorksheetFunction.CountIfs(Location, "camy*", SCAT, ">=226", SCAT, "<=227")
MainJuvNon5 = WorksheetFunction.CountIfs(Location, "camj*", SCAT, ">=242", SCAT, "<=244")
MainJuvNon6 = WorksheetFunction.CountIfs(Location, "camj*", SCAT, ">=246", SCAT, "<=249")
MainJuvNonTotal = MainJuvNon1 + MainJuvNon2 + MainJuvNon3 + MainJuvNon4 + MainJuvNon5 + MainJuvNon6

If MainAll <> MainAdPrint + MainJuvPrintTotal + MainAdNon + MainJuvNonTotal Then
    MsgBox ("Check Main Juv/YA Acquisitions")
End If

COMPsh.Range("H8") = MainAdPrint + SRSPrint + CRPrint + RefPrint
COMPsh.Range("H9") = MainJuvPrintTotal
COMPsh.Range("H13") = MainAdNon
COMPsh.Range("H14") = MainJuvNonTotal

'release the range objects
Set Location = Nothing
Set SCAT = Nothing

ACQwb.Close

End Sub

