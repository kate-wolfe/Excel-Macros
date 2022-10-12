Attribute VB_Name = "Import"
Option Explicit

Sub GetFunds()

Application.ScreenUpdating = False

'Import Fund Report

Dim textFileNum, rowNum, colNum As Integer
Dim textFileLocation, textDelimiter, textData As String
Dim tArray() As String
Dim sArray() As String
    
textFileLocation = Application.GetOpenFilename()
textDelimiter = ","
textFileNum = FreeFile

Open textFileLocation For Input As textFileNum
textData = Input(LOF(textFileNum), textFileNum)
Close textFileNum
tArray() = Split(textData, vbLf)
For rowNum = LBound(tArray) To UBound(tArray) - 1
    If Len(Trim(tArray(rowNum))) <> 0 Then
        sArray = Split(tArray(rowNum), textDelimiter)
        For colNum = LBound(sArray) To UBound(sArray)
            ActiveSheet.Cells(rowNum + 1, colNum + 1) = sArray(colNum)
        Next colNum
    End If
Next rowNum

ThisWorkbook.Sheets("All Library").Range("C3:G56").NumberFormat = "$#,##0.00"

Call DivvyLocations

Call DivvyVendors

Call Formatting

End Sub
