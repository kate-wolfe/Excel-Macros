Attribute VB_Name = "Base"
Sub CombinedSort()

'This is the central macro that does most of the work.


'Turn off screen animations to make things go faster.
'It's important to include this because otherwise it takes a loooooooong time to run.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

On Error Resume Next

'Define variables.

Dim Email As Worksheet, Comp As Worksheet, Instructions As Worksheet
Dim LocalHolds As Worksheet, BranchHolds As Worksheet, BothHolds As Worksheet, GrayBins As Worksheet
Dim Emailrng As Range, CopyRange As Range, loccell As Range, EmailEnd As Range
Dim Last4 As Range, CallNos As Range, Titles As Range, Barcodes As Range
Dim i As Integer, loc As Integer

Set Email = Sheets("Paste Email Here")
Set Comp = Sheets("Complete")
Set Instructions = Sheets("Instructions")
Set LocalHolds = Sheets("Local Holds")
Set BranchHolds = Sheets("Open Branch Holds")
Set BothHolds = Sheets("Local + Branch Holds")
Set GrayBins = Sheets("Gray Bins")

Set CopyRange = Comp.Range("C2:F4000")

Set Last4 = Comp.Columns("C:C")
Set CallNos = Comp.Columns("D:D")
Set Titles = Comp.Columns("E:E")
Set Barcodes = Comp.Columns("F:F")

Set loccell = Instructions.Range("J4") 'This checks the Instructions sheet to see which Cambridge location is selected.

'This determines which branch is running the list.
'This is used to figure out whether a hold is local or for a branch.

If loccell.Value2 = "Main" Then
loc = 1
ElseIf loccell.Value2 = "Boudreau" Then
loc = 5
ElseIf loccell.Value2 = "Central Square" Then
loc = 5
ElseIf loccell.Value2 = "Collins" Then
loc = 6
ElseIf loccell.Value2 = "O'Connell" Then
loc = 7
ElseIf loccell.Value2 = "O'Neill" Then
loc = 8
ElseIf loccell.Value2 = "Valente" Then
loc = 9
End If

'This formats the text from the original email so that it can be easily searched and sorted by later macros.
Set Emailrng = ActiveSheet.Range("A1:A20000")

With Email.Range("A1:K20000")
.ClearFormats
.HorizontalAlignment = xlGeneral
.Font.Size = 10
End With


'This runs through every line of the Email sheet.
'When it finds a barcode, it takes item information from nearb y cells and moves it to the "Complete" sheet.
'Not that using "Value2" and "=" in this way is much faster than copying and pasting the information from one sheet to another.
i = 2
Do While i < 20000
Email.Cells(i, 1).Value2 = Email.Cells(i, 2).Value2 & " " & Email.Cells(i, 5).Value2 & Email.Cells(i, 6).Value2
If InStr(1, ActiveSheet.Cells(i, 1).Value2, "31189") > 0 Then 'If we find a "31189" barcode...
Comp.Cells(i, 6).Value2 = Email.Cells(i, 1).Value2 'Move over the barcode
Comp.Cells(i, 3).Value2 = Right(Email.Cells(i, 1).Value2, 4) 'Barcode Last4
Comp.Cells(i, 7).Value2 = Email.Cells((i + 1), 2).Value2 'Pickup Location. This one is (i,2) because it's the next cell and hasn't been "converted" yet.
Comp.Cells(i, 5).Value2 = Email.Cells((i - 1), 1).Value2 'Title
Comp.Cells(i, 4).Value2 = Email.Cells((i - 2), 1).Value2 'Call No
Comp.Cells(i, 2).Value2 = Email.Cells((i - 3), 1).Value2 'Location
i = i + 2
End If
i = i + 1
Loop

'Running through the above macro leaves a bunch of blank spaces on the "Complete" sheet.
'The below deletes all of those blank spaces so we're left with only item information.
Comp.Range("D2:D20000").SpecialCells(xlCellTypeBlanks).EntireRow.Delete





''This is a long "Do While" Loop that does a lot to sort the sheet.
i = 2
Do While Comp.Cells(i, 4) <> ""

'The next long chunk of code labels CAM and non-CAM Holds so they can be sorted later.
'These need to be sorted differently for each branch, so there's a nested if statement for each of them.
'Here's what these mean:
'1 = Local Pickup
'2 = Branch Pickup
'3 = Gray bin Pickup
'4 = Don't Page (for closed branches). Items with a 4 will not appear on the final list.

'Main edited 1/10/22 to include Boud and Coll open zz
If loc = 1 Then
        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 1
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        Else
            Comp.Cells(i, 8).Value2 = 3
        End If
        
'All of the below ones are in case a branch uses this sheet. You won't need to make any changes to it for Main.
'Boudreau
ElseIf loc = 4 Then
        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 1
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        Else
            Comp.Cells(i, 8).Value2 = 3
        End If
'Central Square
ElseIf loc = 5 Then
        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 1
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        Else
            Comp.Cells(i, 8).Value2 = 3
        End If
'Collins
ElseIf loc = 6 Then
        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 1
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        Else
            Comp.Cells(i, 8).Value2 = 3
        End If
'O'Connell
ElseIf loc = 7 Then
        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 1
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        Else
            Comp.Cells(i, 8).Value2 = 3
        End If
'O'Neill
ElseIf loc = 8 Then
        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 1
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        Else
            Comp.Cells(i, 8).Value2 = 3
        End If
'Valente
ElseIf loc = 9 Then
        If InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/BOUDREAU/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/CENT SQ/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/COLLINS/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/OCONNELL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/ONEILL/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 2
        ElseIf InStr(1, Comp.Cells(i, 7).Value2, "CAMBRIDGE/VALENTE/Pickup") = 1 Then
            Comp.Cells(i, 8).Value2 = 1
        Else
            Comp.Cells(i, 8).Value2 = 3
        End If
End If


'Shorten titles so they fit in the cells better.
Comp.Cells(i, 5).Value2 = Left(Comp.Cells(i, 5).Value2, 50)

'This moves info on new books to the Call Number column, making it easier to sort for them later.
If Comp.Cells(i, 2).Value2 = ("New Books ") Then Comp.Cells(i, 4).Value2 = "New " & Comp.Cells(i, 4).Value2
i = i + 1
Loop

'Column G has the full pickup information. We don't need it now that every item has been assigned "1,2,3,4"
Comp.Columns("G").Delete

Comp.Range("D1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes 'Remove duplicate items from the list.

'This calls the Stats macro, which opens the "Paging Stats" file and saves the information there.
Call Stats.OpenStats


'This formats the "Complete" sheet. We almost never print it, but at least now it looks nice.
Comp.Activate
Call Headers


'If you're running this at Main, the below runs all of the location sorting Macros.
If loc = 1 Then
    Call Main.NewList
    Call Main.MezzList
    Call Main.LOneList
    Call Main.StoneList
    Call Main.SecondList
'    Call Main.JEastList
'    Call Main.JCenterList
'    Call Main.JWestList
    
    CallNos.AutoFilter 'Remove any lingering filters from the above sorting
    Comp.Columns("G:G").Delete ' Delete the "1,2,3,4" from pickup location sorting
    Comp.Range("B2:B20000").ClearContents 'Delete the "New" information for any new books
    Comp.Visible = xlSheetVisible 'Unhide the "Complete" sheet

Else

'If you're running this at a branch, it uses a simplified sort that puts everything on one page.
Call OneList.listshift
End If

'Hide the now-unnecessary sheets.
Email.Visible = xlSheetHidden
ThisWorkbook.Sheets("Instructions").Visible = xlSheetHidden

End Sub

Sub Split()
'This macro splits items onto different sheets based on their pickup location.
'It is called by all of the "sorting" macros like NewBks, Mezz, Listshift, etc.
'It looks at the location number (1,2,3,4) for every item on a given sheet and moves them to the appropriate sheets.

Dim LocalHolds As Worksheet, BranchHolds As Worksheet, BothHolds As Worksheet, GrayBins As Worksheet
Dim LocalEnd As Range, BranchEnd As Range, BothEnd As Range, GrayEnd As Range
Dim i As Integer

    Set LocalHolds = Sheets("Local Holds")
    Set BranchHolds = Sheets("Open Branch Holds")
    Set BothHolds = Sheets("Local + Branch Holds")
    Set GrayBins = Sheets("Gray Bins")
    
    'These ranges with "End" in them find the last row with content in them for each of our sheets.
    Set LocalEnd = LocalHolds.Cells(Rows.Count, 3).End(xlUp)
    Set BranchEnd = BranchHolds.Cells(Rows.Count, 3).End(xlUp)
    Set BothEnd = BothHolds.Cells(Rows.Count, 3).End(xlUp)
    Set GrayEnd = GrayBins.Cells(Rows.Count, 3).End(xlUp)

    With LocalEnd
        .Offset(1, 0).Value2 = ActiveSheet.Name
        .Offset(1, 0).Font.Bold = True
    End With
    
    With BranchEnd
        .Offset(1, 0).Value2 = ActiveSheet.Name
        .Offset(1, 0).Font.Bold = True

    End With
    
    With BothEnd
        .Offset(1, 0).Value2 = ActiveSheet.Name
        .Offset(1, 0).Font.Bold = True
    End With
    
    With GrayEnd
        .Offset(1, 0).Value2 = ActiveSheet.Name
        .Offset(1, 0).Font.Bold = True
    End With
 
      
    ActiveSheet.Range("C1").CurrentRegion.RemoveDuplicates Columns:=Array(1), Header:=xlYes
  
'This "Do While" Loop sorts everything out by location number.
i = 2
Do While ActiveSheet.Cells(i, 7).Value2 <> ""

    Set LocalEnd = LocalHolds.Cells(Rows.Count, 3).End(xlUp)
    Set BranchEnd = BranchHolds.Cells(Rows.Count, 3).End(xlUp)
    Set BothEnd = BothHolds.Cells(Rows.Count, 3).End(xlUp)
    Set GrayEnd = GrayBins.Cells(Rows.Count, 3).End(xlUp)
    
'Local Pickups go on the Local and Both lists.
If ActiveSheet.Cells(i, 7).Value2 = 1 Then
    Range(LocalEnd.Offset(1, 0), LocalEnd.Offset(1, 3)).Value2 = Range(ActiveSheet.Cells(i, 3), ActiveSheet.Cells(i, 6)).Value2
    Range(BothEnd.Offset(1, 0), BothEnd.Offset(1, 3)).Value2 = Range(ActiveSheet.Cells(i, 3), ActiveSheet.Cells(i, 6)).Value2

'Branch Pickups go on the Branch and Both lists.
ElseIf ActiveSheet.Cells(i, 7).Value2 = 2 Then
    Range(BranchEnd.Offset(1, 0), BranchEnd.Offset(1, 3)).Value2 = Range(ActiveSheet.Cells(i, 3), ActiveSheet.Cells(i, 6)).Value2
    Range(BothEnd.Offset(1, 0), BothEnd.Offset(1, 3)).Value2 = Range(ActiveSheet.Cells(i, 3), ActiveSheet.Cells(i, 6)).Value2

'GrayBin pickups just go on their own lists.
ElseIf ActiveSheet.Cells(i, 7).Value2 = 3 Then
    Range(GrayEnd.Offset(1, 0), GrayEnd.Offset(1, 3)).Value2 = Range(ActiveSheet.Cells(i, 3), ActiveSheet.Cells(i, 6)).Value2

End If
i = i + 1
Loop


    Set LocalEnd = LocalHolds.Cells(Rows.Count, 3).End(xlUp)
    Set BranchEnd = BranchHolds.Cells(Rows.Count, 3).End(xlUp)
    Set BothEnd = BothHolds.Cells(Rows.Count, 3).End(xlUp)
    Set GrayEnd = GrayBins.Cells(Rows.Count, 3).End(xlUp)
    
    
    'This inserts page breaks at the end of each section, so that the sheets look better when printed.
    LocalEnd.Offset(1, 0).EntireRow.PageBreak = xlManual
    BranchEnd.Offset(1, 0).EntireRow.PageBreak = xlManual
    BothEnd.Offset(1, 0).EntireRow.PageBreak = xlManual
    GrayEnd.Offset(1, 0).EntireRow.PageBreak = xlManual
    
       

End Sub

Sub Headers()

'This sets up the headers and all the other formatting for each sheet. It's called by all of the "sorting" macros.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False


'This sets up the header.
With ActiveSheet
.PageSetup.LeftHeader = "&A" 'Name of the sheet
.PageSetup.CenterHeader = "&D &T" 'Date and Time the list was run
.PageSetup.RightHeader = "&P of &N" 'Page number
End With

Dim SheetEnd As Range, PickRng As Range, i As Integer

'This formats each of the columns.
    Columns("A:A").ColumnWidth = 5.25
    Columns("B:B").ColumnWidth = 3.75
    
    With Columns("C:C")
    .NumberFormat = "0000"
    .ColumnWidth = 7.25
    .HorizontalAlignment = xlLeft
    End With
    
    Columns("D:D").ColumnWidth = 25.75
    Columns("E:E").ColumnWidth = 40
            
    With Columns("F:F")
    .NumberFormat = "0"
    .ColumnWidth = 15.25
    .Font.Size = 10
    End With
    

    Cells(1, 1).Value2 = "Found"
    Cells(1, 2).Value2 = "NOS"
    Cells(1, 3).Value2 = "Last4"
    Cells(1, 4).Value2 = "Call Number"
    Cells(1, 5).Value2 = "Title"
    Cells(1, 6).Value2 = "Barcode"
    Cells(1, 6).Font.Size = 10
    
    Rows("1:1").Font.Bold = True
    
    'This draws the borders for each cell.
    Set SheetEnd = ActiveSheet.Cells(Rows.Count, 6).End(xlUp)
    With ActiveSheet.Range("A1", SheetEnd).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

'The next few macros correspond with the buttons on the "Paste Email Here" page.

'This is the "Local + Open Branch Holds" button.
Sub BothDisplay()

'Turn off screen animations to make things go faster.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False


'Define variables.
Dim Email As Worksheet, Comp As Worksheet
Dim LocalHolds As Worksheet, BranchHolds As Worksheet, BothHolds As Worksheet, GrayBins As Worksheet

Set Email = Sheets("Paste Email Here")
Set Comp = Sheets("Complete")
Set LocalHolds = Sheets("Local Holds")
Set BranchHolds = Sheets("Open Branch Holds")
Set BothHolds = Sheets("Local + Branch Holds")
Set GrayBins = Sheets("Gray Bins")

Call CombinedSort

BothHolds.Activate
Call Headers

BothHolds.Visible = xlSheetVisible

End Sub


'This is the "Gray Bins" button.

Sub GrayDisplay()

'Turn off screen animations to make things go faster.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False


'Define variables.
Dim Email As Worksheet, Comp As Worksheet
Dim LocalHolds As Worksheet, BranchHolds As Worksheet, BothHolds As Worksheet, GrayBins As Worksheet

Set Email = Sheets("Paste Email Here")
Set Comp = Sheets("Complete")
Set LocalHolds = Sheets("Local Holds")
Set BranchHolds = Sheets("Open Branch Holds")
Set BothHolds = Sheets("Local + Branch Holds")
Set GrayBins = Sheets("Gray Bins")

Call CombinedSort

GrayBins.Activate
Call Headers

GrayBins.Visible = xlSheetVisible
End Sub



'This is the "Everything" button.
'Not that the "Local Holds" and "Branch Holds stuff is commented out because we never use it.
Sub CompleteDisplay()

'Turn off screen animations to make things go faster.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False


'Define variables.
Dim Email As Worksheet, Comp As Worksheet
Dim LocalHolds As Worksheet, BranchHolds As Worksheet, BothHolds As Worksheet, GrayBins As Worksheet

Set Email = Sheets("Paste Email Here")
Set Comp = Sheets("Complete")
Set LocalHolds = Sheets("Local Holds")
Set BranchHolds = Sheets("Open Branch Holds")
Set BothHolds = Sheets("Local + Branch Holds")
Set GrayBins = Sheets("Gray Bins")

Call CombinedSort

BothHolds.Activate
Call Headers

GrayBins.Activate
Call Headers


BothHolds.Visible = xlSheetVisible
GrayBins.Visible = xlSheetVisible
End Sub



