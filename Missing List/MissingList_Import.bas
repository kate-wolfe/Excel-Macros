Attribute VB_Name = "Import"
Sub OpenFile()

Dim filter As String
Dim caption As String
Dim FileName As String
Dim Full As Worksheet

Application.ScreenUpdating = False
Application.Calculation = xlManual

ChDrive "H:\"
    ChDir "H:\My Documents"

filter = "Text files (*.txt),*.txt"
caption = "Please select the current Missing List"
FileName = Application.GetOpenFilename(filter, , caption)
Set Full = ThisWorkbook.Sheets("Full")

' OpenFile Macro

With Full.QueryTables.Add(Connection:= _
        "TEXT;" & FileName, _
        Destination:=Full.Range("$C$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "^"
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
Call Formatting

Call SectionOut

Call PageSetup

Call Transfer

End Sub
