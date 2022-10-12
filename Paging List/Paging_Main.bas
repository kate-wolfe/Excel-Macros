Attribute VB_Name = "Main"
'These macros each sort and format the materials for a different part of the building.
'They mostly follow the same structure, so I've annotated the one for New Books and made notes on the others where they differ.


Sub NewList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Define variables.
    Dim Comp As Worksheet, NewW As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim NewRange, nd As Range, NewEnd As Range
    Dim Secret As Worksheet
    Dim i As Integer

    Set Secret = ThisWorkbook.Sheets("Secret")
    Set NewEnd = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set NewW = ThisWorkbook.Sheets("New")
    Set NewRange = NewW.Range("C2:G2500")
    Set nd = NewW.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")
    
'Move New stuff to its own sheet.

'First we make sure that the "Complete" sheet isn't already being filtered.
Comp.AutoFilterMode = False

'Next we filter Complete based on Call #, using the "New Books" column of the hidden "Secret" sheet for reference.
CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("E2", NewEnd), Unique:=False
        
'Now we copy the filtered cells from Complete to the New Books sheet.
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=NewRange

'We then sort the New items by their Call #.
NewRange.Sort Key1:=nd, Order1:=xlAscending, Header:=xlNo

'The below edits the Call #s in the New sheet for brevity.
'When you see "Replace" code like this it replaces the first quoted text with the second quoted text.
'"xlPart" means it only changes that specific part of the text and not the whole cell.

'"vbnullstring" is Excel's way of saying "nothing"
'So when you see something being replaced with vbnullstring, it means it's just deleting that text without replacing it.

With NewW.Range("D2:D2500")
.Replace "New ", vbNullString, xlPart
.Replace "[Express] ", "[Exp] ", xlPart
.Replace "[EXPRESS PB] ", "[PB] ", xlPart
.Replace "[Express] FICTION", "[Exp] FIC", xlPart
.Replace "MYSTERY ", "MYST ", xlPart
.Replace "SCI FIC ", "SCIFI ", xlPart
.Replace "FICTION ", "FIC ", xlPart
.Replace "FIC SHORT STORIES ", "FIC SHORT ", xlPart
.Replace "MYST SHORT STORIES ", "MYST SHORT ", xlPart
.Replace "[Exp] J *", vbNullString, xlPart
End With

'This un-filters Complete so all materials are visible again.
CallNos.AutoFilter


NewW.Activate 'Make Excel look at the New Books sheet
Call Base.Split 'Run the Split macro to sort the items by pickup location
Call Base.Headers 'Run the Headers macro so it is formatted correctly.

NewW.Columns("G:G").Clear 'Delete the pickup location info
NewW.Visible = xlSheetVisible 'Make New Books visible


End Sub

Sub MezzList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp, Mezz As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim MezzRange As Range, md As Range, MezzEnd As Range

    Dim Secret As Worksheet
    Dim i As Integer


    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set Mezz = ThisWorkbook.Sheets("Mezzanine")
    Set MezzRange = Mezz.Range("C2:G2500")
    Set md = Mezz.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")


Comp.AutoFilterMode = False
'Move Mezzanine over

Set MezzEnd = Secret.Cells(Rows.Count, 2).End(xlUp)

    CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("B2", MezzEnd), Unique:=False
        
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=MezzRange


'Call Nos
With Mezz.Range("D2:D1000")
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

'Here we're deleting the call numbers for the J AV items that end up on the list.
.Replace "DVD J *", vbNullString, xlPart
.Replace "CDB J *", vbNullString, xlPart
.Replace "CD J *", vbNullString, xlPart
.Replace "BOP J *", vbNullString, xlPart
End With


'We now delete all of the J AV items by looking for items with no call #.
i = 2
Do While Mezz.Cells(i, 3).Value2 <> ""
If Mezz.Cells(i, 4).Value2 = "" Then
Mezz.Cells(i, 3).EntireRow.Delete
Else: i = i + 1
End If
Loop

'Titles
With Mezz.Range("E2:E1000")
.Replace "[videorecording]", vbNullString, xlPart
.Replace "[sound recording]", vbNullString, xlPart
.Replace "(Musical group)", vbNullString, xlPart
.Replace "[a novel]", vbNullString, xlPart
End With

MezzRange.Sort Key1:=md, Order1:=xlAscending, Header:=xlNo


Mezz.Activate
Call Base.Split
Call Base.Headers

Mezz.Columns("G:G").Clear
Mezz.Visible = xlSheetVisible

End Sub

Sub LOneList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp, LOne As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim LOneRange As Range, ld As Range
   
    Dim Secret As Worksheet

    Dim i As Integer

    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set LOne = ThisWorkbook.Sheets("L1")
    Set LOneRange = LOne.Range("C2:G2500")
    Set ld = LOne.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")


Comp.AutoFilterMode = False
'Move L1 over

Set LOneEnd = Secret.Cells(Rows.Count, 3).End(xlUp)

    CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("C2", LOneEnd), Unique:=False
        
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=LOneRange


'Call Nos
With LOne.Range("D2:D1000")
.Replace "FICTION", "FIC", xlPart
.Replace "SHORT STORIES", "SHORT", xlPart
.Replace "[PB] ROMANCE", "[PB] ROM", xlPart
.Replace "GRAPHIC", "GRAPH", xlPart
.Replace "     ", vbNullString, xlPart
End With

''This adds authors' full names to the Call Nos for Romance books.
'i = 2
'Do While LOne.Cells(i, 4) <> ""
'If InStr(1, LOne.Cells(i, 4).Value2, "[PB] ROM") > 0 Then
'LOne.Cells(i, 10).Value2 = LOne.Cells(i, 5).Value2
'LOne.Cells(i, 10).Replace "*/", vbNullString, xlPart
'LOne.Cells(i, 4).Value2 = Left(LOne.Cells(i, 4).Value2, 10) & " " & LOne.Cells(i, 10).Value2
'LOne.Cells(i, 10).Clear
'End If
'i = i + 1
'Loop

LOneRange.Sort Key1:=ld, Order1:=xlAscending, Header:=xlNo

LOne.Activate
Call Base.Split
Call Base.Headers

LOne.Columns("G:G").Clear
LOne.Visible = xlSheetVisible

End Sub


Sub StoneList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp, Stone As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim StoneRange As Range, sd As Range, StoneEnd As Range
  
    Dim Secret As Worksheet

    Dim i As Integer

    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set Stone = ThisWorkbook.Sheets("Stone")
    Set StoneRange = Stone.Range("C2:G2500")
    Set sd = Stone.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")


Comp.AutoFilterMode = False


'Move Stone over

Set StoneEnd = Secret.Cells(Rows.Count, 1).End(xlUp)

    CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("A2", StoneEnd), Unique:=False
        
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=StoneRange

'Call No cleanup
Stone.Range("D2:D400").Replace "MYSTERY", "MYST", xlPart

StoneRange.Sort Key1:=sd, Order1:=xlAscending, Header:=xlNo


Stone.Activate
Call Base.Split
Call Base.Headers

Stone.Columns("G:G").Clear
Stone.Visible = xlSheetVisible

End Sub

Sub SecondList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp, Second As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim SecondRange As Range, sd As Range, SecondEnd As Range
   
    Dim Secret As Worksheet
    Dim se As Range

    Dim i As Integer


    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set Second = ThisWorkbook.Sheets("2nd Floor")
    Set SecondRange = Second.Range("C2:G2500")
    Set sd = Second.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")


Comp.AutoFilterMode = False


'Move Stone over

Set SecondEnd = Secret.Cells(Rows.Count, 4).End(xlUp)

    CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("D2", SecondEnd), Unique:=False
        
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=SecondRange

With Second.Range("D2:D800")
.Replace "POETRY", "POET", xlPart
.Replace "[Business]", "[Biz]", xlPart
.Replace "[Home & Health] ", vbNullString, xlPart
End With


SecondRange.Sort Key1:=sd, Order1:=xlAscending, Header:=xlNo

Second.Activate
Call Base.Split
Call Base.Headers

Second.Columns("G:G").Clear
Second.Visible = xlSheetVisible

End Sub

Sub JEastList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp, JEast As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim JEastRange As Range, jed As Range, JEastEnd As Range
    
    Dim Secret As Worksheet

    Dim i As Integer

    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set JEast = ThisWorkbook.Sheets("J East")
    Set JEastRange = JEast.Range("C2:G2500")
    Set jed = JEast.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")

Comp.AutoFilterMode = False

'Move JEAST over

Set JEastEnd = Secret.Cells(Rows.Count, 6).End(xlUp)

    CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("F2", JEastEnd), Unique:=False
        
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=JEastRange


JEast.Range("D2:D500").Replace "[Express]", "[Exp]", xlPart

JEastRange.Sort Key1:=jed, Order1:=xlAscending, Header:=xlNo


JEast.Activate
Call Base.Split
Call Base.Headers

JEast.Columns("G:G").Clear
JEast.Visible = xlSheetVisible

End Sub

Sub JCenterList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp, JCenter As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim JCenterRange As Range, jcd As Range, JCenterEnd As Range
   
    Dim Secret As Worksheet
    Dim se As Range

    Dim i As Integer

    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set JCenter = ThisWorkbook.Sheets("J Center")
    Set JCenterRange = JCenter.Range("C2:G2500")
    Set jcd = JCenter.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")

Comp.AutoFilterMode = False

'Move JCenter over

Set JCenterEnd = Secret.Cells(Rows.Count, 7).End(xlUp)

    CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("G2", JCenterEnd), Unique:=False
        
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=JCenterRange

JCenterRange.Sort Key1:=jcd, Order1:=xlAscending, Header:=xlNo


JCenter.Activate
Call Base.Split
Call Base.Headers

JCenter.Columns("G:G").Clear
JCenter.Visible = xlSheetVisible

End Sub


Sub JWestList()
'Turn off screen animations to speed things up.
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayStatusBar = False

'Defining objects to save memory
    Dim Comp, JCenter As Worksheet
    Dim Codes, CallNos, Titles, CopyRange As Range
    Dim JWestRange As Range, jwd As Range, JWestEnd As Range
   
    Dim Secret As Worksheet
    Dim se As Range

    Dim i As Integer

    Set Secret = ThisWorkbook.Sheets("Secret")
    Set se = Secret.Cells(Rows.Count, 5).End(xlUp)

'Setting worksheets and ranges to shorten code
    Set Comp = ThisWorkbook.Sheets("Complete")
    Set JWest = ThisWorkbook.Sheets("J West")
    Set JWestRange = JWest.Range("C2:G2500")
    Set jwd = JWest.Range("D2")
    Set CallNos = Comp.Columns("D:D")
    Set CopyRange = Comp.Range("C2:G2500")


Comp.AutoFilterMode = False


'Move JWest over

Set JWestEnd = Secret.Cells(Rows.Count, 8).End(xlUp)

    CallNos.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:= _
        Sheets("Secret").Range("H2", JWestEnd), Unique:=False
        
  CopyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=JWestRange


JWestRange.Sort Key1:=jwd, Order1:=xlAscending, Header:=xlNo

JWest.Activate
Call Base.Split
Call Base.Headers

JWest.Columns("G:G").Clear
JWest.Visible = xlSheetVisible

End Sub

