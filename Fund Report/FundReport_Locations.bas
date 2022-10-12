Attribute VB_Name = "Locations"
Option Explicit

Sub DivvyLocations()

'Declare worksheet variables

Dim AllLib, MainLib, BranchLib, locArrWS As Worksheet

Set AllLib = ThisWorkbook.Sheets("All Library")
Set MainLib = ThisWorkbook.Sheets("Main")
Set BranchLib = ThisWorkbook.Sheets("Branches")
Set locArrWS = ThisWorkbook.Sheets("LOCarrays")

Dim currentYear As Integer
Dim curFY As Variant

currentYear = IIf(Month(Now) <= 6, Year(Now), Year(Now) + 1)
curFY = "FY" & Right(currentYear, 2)

'Declare criteria variables

Dim camLROW As Long
Dim ca4LROW As Long
Dim ca5LROW As Long
Dim ca6LROW As Long
Dim ca7LROW As Long
Dim ca8LROW As Long
Dim ca9LROW As Long
Dim blaLROW As Long

camLROW = locArrWS.Cells(Rows.Count, 1).End(xlUp).Row
ca4LROW = locArrWS.Cells(Rows.Count, 2).End(xlUp).Row
ca5LROW = locArrWS.Cells(Rows.Count, 3).End(xlUp).Row
ca6LROW = locArrWS.Cells(Rows.Count, 4).End(xlUp).Row
ca7LROW = locArrWS.Cells(Rows.Count, 5).End(xlUp).Row
ca8LROW = locArrWS.Cells(Rows.Count, 6).End(xlUp).Row
ca9LROW = locArrWS.Cells(Rows.Count, 7).End(xlUp).Row
blaLROW = locArrWS.Cells(Rows.Count, 8).End(xlUp).Row

Dim rngCamArr As Range
Dim camArray As Variant

Dim rngCa4Arr As Range
Dim ca4Array As Variant

Dim rngCa5Arr As Range
Dim ca5Array As Variant

Dim rngCa6Arr As Range
Dim ca6Array As Variant

Dim rngCa7Arr As Range
Dim ca7Array As Variant

Dim rngCa8Arr As Range
Dim ca8Array As Variant

Dim rngCa9Arr As Range
Dim ca9Array As Variant

Dim rngBlaArr As Range
Dim blaArray As Variant

Set rngCamArr = locArrWS.Range("A2:A" & camLROW)
Set rngCa4Arr = locArrWS.Range("B2:B" & ca4LROW)
Set rngCa5Arr = locArrWS.Range("C2:C" & ca5LROW)
Set rngCa6Arr = locArrWS.Range("D2:D" & ca6LROW)
Set rngCa7Arr = locArrWS.Range("E2:E" & ca7LROW)
Set rngCa8Arr = locArrWS.Range("F2:F" & ca8LROW)
Set rngCa9Arr = locArrWS.Range("G2:G" & ca9LROW)
Set rngBlaArr = locArrWS.Range("H2:H" & blaLROW)

camArray = rngCamArr.Value
ca4Array = rngCa4Arr.Value
ca5Array = rngCa5Arr.Value
ca6Array = rngCa6Arr.Value
ca7Array = rngCa7Arr.Value
ca8Array = rngCa8Arr.Value
ca9Array = rngCa9Arr.Value
blaArray = rngBlaArr.Value


'All Library Spent

'NOTE: if any funds are added or taken away, the 50 must change

Dim appallSum As Single
Dim expallSum As Single
Dim encallSum As Single
Dim freeallSum As Single
Dim cashallSum As Single
Dim allSpent As Single

appallSum = AllLib.Range("C50")
expallSum = AllLib.Range("D50")
encallSum = AllLib.Range("E50")
freeallSum = AllLib.Range("F50")
cashallSum = AllLib.Range("G50")

allSpent = (expallSum + encallSum) / appallSum

AllLib.Range("I50") = allSpent
AllLib.Range("I49") = "% Spent"
AllLib.Range("I48") = curFY

'Copy Main over

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(camArray), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

MainLib.Cells(1, 1).PasteSpecial
MainLib.Range("A1") = "Main Library"
MainLib.Range("A1").Interior.Color = vbYellow

Dim mainLast As Long
mainLast = MainLib.Cells(Rows.Count, 1).End(xlUp).Row

Dim mainTotal As Range
Set mainTotal = MainLib.Range("A" & mainLast).Offset(2)

Dim appmainTotal As Range
Dim expmainTotal As Range
Dim encmainTotal As Range
Dim freemainTotal As Range
Dim cashmainTotal As Range
Dim mainSpent As Range

Set appmainTotal = mainTotal.Offset(, 1)
Set expmainTotal = mainTotal.Offset(, 2)
Set encmainTotal = mainTotal.Offset(, 3)
Set freemainTotal = mainTotal.Offset(, 4)
Set cashmainTotal = mainTotal.Offset(, 5)
Set mainSpent = mainTotal.Offset(, 7)

mainTotal = "Total"
appmainTotal.Formula = "=SUM(B2:B" & mainLast & ")"
expmainTotal.Formula = "=SUM(C2:C" & mainLast & ")"
encmainTotal.Formula = "=SUM(D2:D" & mainLast & ")"
freemainTotal.Formula = "=SUM(E2:E" & mainLast & ")"
cashmainTotal.Formula = "=SUM(F2:F" & mainLast & ")"

mainSpent.Formula = "=(" & expmainTotal & "+" & encmainTotal & ")/" & appmainTotal
mainTotal.Offset(-1, 7) = "% Spent"
mainTotal.Offset(-2, 7) = curFY

'Copy Boudreau over

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(ca4Array), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

BranchLib.Cells(1, 1).PasteSpecial
BranchLib.Range("A1") = "Boudreau"
BranchLib.Range("A1").Interior.Color = vbYellow

Dim boudLast As Long
boudLast = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row

Dim boudTotal As Range
Set boudTotal = BranchLib.Range("A" & boudLast).Offset(2)

Dim appboudTotal As Range
Dim expboudTotal As Range
Dim encboudTotal As Range
Dim freeboudTotal As Range
Dim cashboudTotal As Range
Dim boudSpent As Range

Set appboudTotal = boudTotal.Offset(, 1)
Set expboudTotal = boudTotal.Offset(, 2)
Set encboudTotal = boudTotal.Offset(, 3)
Set freeboudTotal = boudTotal.Offset(, 4)
Set cashboudTotal = boudTotal.Offset(, 5)
Set boudSpent = boudTotal.Offset(, 7)

boudTotal = "Total"
appboudTotal.Formula = "=SUM(B2:B" & boudLast & ")"
expboudTotal.Formula = "=SUM(C2:C" & boudLast & ")"
encboudTotal.Formula = "=SUM(D2:D" & boudLast & ")"
freeboudTotal.Formula = "=SUM(E2:E" & boudLast & ")"
cashboudTotal.Formula = "=SUM(F2:F" & boudLast & ")"

boudSpent.Formula = "=(" & expboudTotal & "+" & encboudTotal & ")/" & appboudTotal
boudTotal.Offset(-1, 7) = "% Spent"
boudTotal.Offset(-2, 7) = curFY

'Copy CSQ over

Dim csqTitleRow As Long
Dim csqTitle As Range

csqTitleRow = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row
Set csqTitle = BranchLib.Range("A" & csqTitleRow).Offset(4)

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(ca5Array), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

csqTitle.PasteSpecial
csqTitle = "CSQ"
csqTitle.Interior.Color = vbYellow

Dim csqLast As Long
csqLast = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row

Dim csqTotal As Range
Set csqTotal = BranchLib.Range("A" & csqLast).Offset(2)

Dim appcsqTotal As Range
Dim expcsqTotal As Range
Dim enccsqTotal As Range
Dim freecsqTotal As Range
Dim cashcsqTotal As Range
Dim csqSpent As Range

Set appcsqTotal = csqTotal.Offset(, 1)
Set expcsqTotal = csqTotal.Offset(, 2)
Set enccsqTotal = csqTotal.Offset(, 3)
Set freecsqTotal = csqTotal.Offset(, 4)
Set cashcsqTotal = csqTotal.Offset(, 5)
Set csqSpent = csqTotal.Offset(, 7)

csqTotal = "Total"
appcsqTotal.Formula = "=SUM(B" & csqTitleRow + 5 & ":B" & csqLast & ")"
expcsqTotal.Formula = "=SUM(C" & csqTitleRow + 5 & ":C" & csqLast & ")"
enccsqTotal.Formula = "=SUM(D" & csqTitleRow + 5 & ":D" & csqLast & ")"
freecsqTotal.Formula = "=SUM(E" & csqTitleRow + 5 & ":E" & csqLast & ")"
cashcsqTotal.Formula = "=SUM(F" & csqTitleRow + 5 & ":F" & csqLast & ")"

csqSpent.Formula = "=(" & expcsqTotal & "+" & enccsqTotal & ")/" & appcsqTotal
csqTotal.Offset(-1, 7) = "% Spent"
csqTotal.Offset(-2, 7) = curFY

'Copy Collins over

Dim colTitleRow As Long
Dim colTitle As Range

colTitleRow = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row
Set colTitle = BranchLib.Range("A" & colTitleRow).Offset(4)

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(ca6Array), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

colTitle.PasteSpecial
colTitle = "Collins"
colTitle.Interior.Color = vbYellow

Dim colLast As Long
colLast = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row

Dim colTotal As Range
Set colTotal = BranchLib.Range("A" & colLast).Offset(2)

Dim appcolTotal As Range
Dim expcolTotal As Range
Dim enccolTotal As Range
Dim freecolTotal As Range
Dim cashcolTotal As Range
Dim colSpent As Range

Set appcolTotal = colTotal.Offset(, 1)
Set expcolTotal = colTotal.Offset(, 2)
Set enccolTotal = colTotal.Offset(, 3)
Set freecolTotal = colTotal.Offset(, 4)
Set cashcolTotal = colTotal.Offset(, 5)
Set colSpent = colTotal.Offset(, 7)

colTotal = "Total"
appcolTotal.Formula = "=SUM(B" & colTitleRow + 5 & ":B" & colLast & ")"
expcolTotal.Formula = "=SUM(C" & colTitleRow + 5 & ":C" & colLast & ")"
enccolTotal.Formula = "=SUM(D" & colTitleRow + 5 & ":D" & colLast & ")"
freecolTotal.Formula = "=SUM(E" & colTitleRow + 5 & ":E" & colLast & ")"
cashcolTotal.Formula = "=SUM(F" & colTitleRow + 5 & ":F" & colLast & ")"

colSpent.Formula = "=(" & expcolTotal & "+" & enccolTotal & ")/" & appcolTotal
colTotal.Offset(-1, 7) = "% Spent"
colTotal.Offset(-2, 7) = curFY

'Copy OConnell over

Dim oconnTitleRow As Long
Dim oconnTitle As Range

oconnTitleRow = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row
Set oconnTitle = BranchLib.Range("A" & oconnTitleRow).Offset(4)

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(ca7Array), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

oconnTitle.PasteSpecial
oconnTitle = "OConnell"
oconnTitle.Interior.Color = vbYellow

Dim oconnLast As Long
oconnLast = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row

Dim oconnTotal As Range
Set oconnTotal = BranchLib.Range("A" & oconnLast).Offset(2)

Dim appoconnTotal As Range
Dim expoconnTotal As Range
Dim encoconnTotal As Range
Dim freeoconnTotal As Range
Dim cashoconnTotal As Range
Dim oconnSpent As Range

Set appoconnTotal = oconnTotal.Offset(, 1)
Set expoconnTotal = oconnTotal.Offset(, 2)
Set encoconnTotal = oconnTotal.Offset(, 3)
Set freeoconnTotal = oconnTotal.Offset(, 4)
Set cashoconnTotal = oconnTotal.Offset(, 5)
Set oconnSpent = oconnTotal.Offset(, 7)

oconnTotal = "Total"
appoconnTotal.Formula = "=SUM(B" & oconnTitleRow + 5 & ":B" & oconnLast & ")"
expoconnTotal.Formula = "=SUM(C" & oconnTitleRow + 5 & ":C" & oconnLast & ")"
encoconnTotal.Formula = "=SUM(D" & oconnTitleRow + 5 & ":D" & oconnLast & ")"
freeoconnTotal.Formula = "=SUM(E" & oconnTitleRow + 5 & ":E" & oconnLast & ")"
cashoconnTotal.Formula = "=SUM(F" & oconnTitleRow + 5 & ":F" & oconnLast & ")"

oconnSpent.Formula = "=(" & expoconnTotal & "+" & encoconnTotal & ")/" & appoconnTotal
oconnTotal.Offset(-1, 7) = "% Spent"
oconnTotal.Offset(-2, 7) = curFY

'Copy ONeill over

Dim oneillTitleRow As Long
Dim oneillTitle As Range

oneillTitleRow = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row
Set oneillTitle = BranchLib.Range("A" & oneillTitleRow).Offset(4)

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(ca8Array), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

oneillTitle.PasteSpecial
oneillTitle = "ONeill"
oneillTitle.Interior.Color = vbYellow

Dim oneillLast As Long
oneillLast = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row

Dim oneillTotal As Range
Set oneillTotal = BranchLib.Range("A" & oneillLast).Offset(2)

Dim apponeillTotal As Range
Dim exponeillTotal As Range
Dim enconeillTotal As Range
Dim freeoneillTotal As Range
Dim cashoneillTotal As Range
Dim oneillSpent As Range

Set apponeillTotal = oneillTotal.Offset(, 1)
Set exponeillTotal = oneillTotal.Offset(, 2)
Set enconeillTotal = oneillTotal.Offset(, 3)
Set freeoneillTotal = oneillTotal.Offset(, 4)
Set cashoneillTotal = oneillTotal.Offset(, 5)
Set oneillSpent = oneillTotal.Offset(, 7)

oneillTotal = "Total"
apponeillTotal.Formula = "=SUM(B" & oneillTitleRow + 5 & ":B" & oneillLast & ")"
exponeillTotal.Formula = "=SUM(C" & oneillTitleRow + 5 & ":C" & oneillLast & ")"
enconeillTotal.Formula = "=SUM(D" & oneillTitleRow + 5 & ":D" & oneillLast & ")"
freeoneillTotal.Formula = "=SUM(E" & oneillTitleRow + 5 & ":E" & oneillLast & ")"
cashoneillTotal.Formula = "=SUM(F" & oneillTitleRow + 5 & ":F" & oneillLast & ")"

oneillSpent.Formula = "=(" & exponeillTotal & "+" & enconeillTotal & ")/" & apponeillTotal
oneillTotal.Offset(-1, 7) = "% Spent"
oneillTotal.Offset(-2, 7) = curFY

'Copy Valente over

Dim valTitleRow As Long
Dim valTitle As Range

valTitleRow = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row
Set valTitle = BranchLib.Range("A" & valTitleRow).Offset(4)

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(ca9Array), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

valTitle.PasteSpecial
valTitle = "Valente"
valTitle.Interior.Color = vbYellow

Dim valLast As Long
valLast = BranchLib.Cells(Rows.Count, 1).End(xlUp).Row

Dim valTotal As Range
Set valTotal = BranchLib.Range("A" & valLast).Offset(2)

Dim appvalTotal As Range
Dim expvalTotal As Range
Dim encvalTotal As Range
Dim freevalTotal As Range
Dim cashvalTotal As Range
Dim valSpent As Range

Set appvalTotal = valTotal.Offset(, 1)
Set expvalTotal = valTotal.Offset(, 2)
Set encvalTotal = valTotal.Offset(, 3)
Set freevalTotal = valTotal.Offset(, 4)
Set cashvalTotal = valTotal.Offset(, 5)
Set valSpent = valTotal.Offset(, 7)

valTotal = "Total"
appvalTotal.Formula = "=SUM(B" & valTitleRow + 5 & ":B" & valLast & ")"
expvalTotal.Formula = "=SUM(C" & valTitleRow + 5 & ":C" & valLast & ")"
encvalTotal.Formula = "=SUM(D" & valTitleRow + 5 & ":D" & valLast & ")"
freevalTotal.Formula = "=SUM(E" & valTitleRow + 5 & ":E" & valLast & ")"
cashvalTotal.Formula = "=SUM(F" & valTitleRow + 5 & ":F" & valLast & ")"

valSpent.Formula = "=(" & expvalTotal & "+" & encvalTotal & ")/" & appvalTotal
valTotal.Offset(-1, 7) = "% Spent"
valTotal.Offset(-2, 7) = curFY

End Sub
