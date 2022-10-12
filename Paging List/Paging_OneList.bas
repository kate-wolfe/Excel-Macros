Attribute VB_Name = "OneList"
Sub listshift()

'This is a simplified macro replaces the more in-depth sorting macros that are used at Main.
'This was designed to be used at the branches, but it'll work at Main too if you just tweak the "CombinedSort" macro accordingly.
'It doesn't sort the materials on the sheet at all, it just moves them all to one central list.


'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp As Worksheet, FullList As Worksheet
    Dim Codes As Range, CallNos As Range, Titles As Range, CopyRange As Range, FullRange As Range
   
    Dim Secret As Worksheet

    Dim i As Integer

    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")
    
    Set FullList = ThisWorkbook.Sheets("Full List")
    Set FullRange = FullList.Range("c2:g2500")
    
'This simply moves the whole sheet into the "Full List" sheet.
    FullRange.Value2 = CopyRange.Value2
'This does some minor call number replacements.


    With FullList.Range("D2:D1000")
        .Replace "Fiction ", "FIC ", xlPart
        .Replace "MYSTERY ", "MYS ", xlPart
        .Replace "SCI FIC ", "SCIFI ", xlPart
        .Replace "DVD J ", "J DVD ", xlPart
        .Replace "CDB J ", "J CDB ", xlPart
        .Replace "CD J ", "J CD ", xlPart
    End With
    
    FullRange.Sort Key1:=FullList.Range("D1"), Order1:=xlAscending, Header:=xlNo

    FullList.Activate
    
    Call Base.Split
    Call Base.Headers
    
    FullList.Columns("G:G").Clear
    
    FullList.Visible = xlSheetVisible
End Sub
