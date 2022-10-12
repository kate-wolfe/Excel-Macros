Attribute VB_Name = "Vendors"
Option Explicit

Sub DivvyVendors()

'Declare worksheet variables

Dim AllLib, Vend, venarrWS As Worksheet

Set AllLib = ThisWorkbook.Sheets("All Library")
Set Vend = ThisWorkbook.Sheets("Vendors")
Set venarrWS = ThisWorkbook.Sheets("VENarrays")

Dim currentYear As Integer
Dim curFY As Variant

currentYear = IIf(Month(Now) <= 6, Year(Now), Year(Now) + 1)
curFY = "FY" & Right(currentYear, 2)

'Declare Ingram variables

Dim ingcritLROW As Long
Dim rngIArr As Range
Dim ingArray As Variant

ingcritLROW = venarrWS.Cells(Rows.Count, 1).End(xlUp).Row
Set rngIArr = venarrWS.Range("A2:A" & ingcritLROW)
ingArray = rngIArr.Value


'NOTE: AllLib.Range("B2:G50") -> the G50 should be changed if the number of funds changes.
'I tried with using a "last row" variable, but it wouldn't copy all the third party over when running
'It did work when stepping through, though. :/


'Copy Ingram over

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(ingArray), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

Vend.Cells(1, 1).PasteSpecial
Vend.Range("A1") = "Ingram/55116"
Vend.Range("A1").Interior.Color = vbYellow

Dim ingLast As Long
ingLast = Vend.Cells(Rows.Count, 1).End(xlUp).Row

Dim ingTotal As Range
Set ingTotal = Vend.Range("A" & ingLast).Offset(2)

Dim appiTotal As Range
Dim expiTotal As Range
Dim enciTotal As Range
Dim freeiTotal As Range
Dim cashiTotal As Range
Dim ingSpent As Range

Set appiTotal = ingTotal.Offset(, 1)
Set expiTotal = ingTotal.Offset(, 2)
Set enciTotal = ingTotal.Offset(, 3)
Set freeiTotal = ingTotal.Offset(, 4)
Set cashiTotal = ingTotal.Offset(, 5)
Set ingSpent = ingTotal.Offset(, 7)

ingTotal = "Total"
appiTotal.Formula = "=SUM(B2:B" & ingLast & ")"
expiTotal.Formula = "=SUM(C2:C" & ingLast & ")"
enciTotal.Formula = "=SUM(D2:D" & ingLast & ")"
freeiTotal.Formula = "=SUM(E2:E" & ingLast & ")"
cashiTotal.Formula = "=SUM(F2:F" & ingLast & ")"

ingSpent.Formula = "=(" & expiTotal & "+" & enciTotal & ")/" & appiTotal
ingTotal.Offset(-1, 7) = "% Spent"
ingTotal.Offset(-2, 7) = curFY

'Declare Midwest variables

Dim mwtArray As Variant
Dim rngMArr As Range
Dim mwtcritLROW As Long
Dim mwtTitleRow As Long
Dim mwtTitle As Range

mwtcritLROW = venarrWS.Cells(Rows.Count, 2).End(xlUp).Row
Set rngMArr = venarrWS.Range("B2:B" & mwtcritLROW)
mwtArray = rngMArr.Value
mwtTitleRow = Vend.Cells(Rows.Count, 1).End(xlUp).Row
Set mwtTitle = Vend.Range("A" & mwtTitleRow).Offset(4)


'Copy Midwest over

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(mwtArray), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

mwtTitle.PasteSpecial
mwtTitle = "Midwest/55122"
mwtTitle.Interior.Color = vbYellow

Dim mwtLast As Long
mwtLast = Vend.Cells(Rows.Count, 1).End(xlUp).Row

Dim mwtTotal As Range
Set mwtTotal = Vend.Range("A" & mwtLast).Offset(2)

Dim appmTotal As Range
Dim expmTotal As Range
Dim encmTotal As Range
Dim freemTotal As Range
Dim cashmTotal As Range
Dim mwtSpent As Range

Set appmTotal = mwtTotal.Offset(, 1)
Set expmTotal = mwtTotal.Offset(, 2)
Set encmTotal = mwtTotal.Offset(, 3)
Set freemTotal = mwtTotal.Offset(, 4)
Set cashmTotal = mwtTotal.Offset(, 5)
Set mwtSpent = mwtTotal.Offset(, 7)

mwtTotal = "Total"
appmTotal.Formula = "=SUM(B" & mwtTitleRow + 5 & ":B" & mwtLast & ")"
expmTotal.Formula = "=SUM(C" & mwtTitleRow + 5 & ":C" & mwtLast & ")"
encmTotal.Formula = "=SUM(D" & mwtTitleRow + 5 & ":D" & mwtLast & ")"
freemTotal.Formula = "=SUM(E" & mwtTitleRow + 5 & ":E" & mwtLast & ")"
cashmTotal.Formula = "=SUM(F" & mwtTitleRow + 5 & ":F" & mwtLast & ")"

mwtSpent.Formula = "=(" & expmTotal & "+" & encmTotal & ")/" & appmTotal
mwtTotal.Offset(-1, 7) = "% Spent"
mwtTotal.Offset(-2, 7) = curFY


'Declare Third Party variables

Dim thirdArray As Variant
Dim rngThirdArr As Range
Dim thirdcritLROW As Long
Dim thirdTitleRow As Long
Dim thirdTitle As Range

thirdcritLROW = venarrWS.Cells(Rows.Count, 3).End(xlUp).Row
Set rngThirdArr = venarrWS.Range("C2:C" & thirdcritLROW)
thirdArray = rngThirdArr.Value
thirdTitleRow = Vend.Cells(Rows.Count, 1).End(xlUp).Row
Set thirdTitle = Vend.Range("A" & thirdTitleRow).Offset(4)


'Copy Third Party over

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(thirdArray), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

thirdTitle.PasteSpecial
thirdTitle = "Third Party/Cont/55120"
thirdTitle.Interior.Color = vbYellow

Dim thirdLast As Long
thirdLast = Vend.Cells(Rows.Count, 1).End(xlUp).Row

Dim thirdTotal As Range
Set thirdTotal = Vend.Range("A" & thirdLast).Offset(2)

Dim apptTotal As Range
Dim exptTotal As Range
Dim enctTotal As Range
Dim freetTotal As Range
Dim cashtTotal As Range
Dim thirdSpent As Range

Set apptTotal = thirdTotal.Offset(, 1)
Set exptTotal = thirdTotal.Offset(, 2)
Set enctTotal = thirdTotal.Offset(, 3)
Set freetTotal = thirdTotal.Offset(, 4)
Set cashtTotal = thirdTotal.Offset(, 5)
Set thirdSpent = thirdTotal.Offset(, 7)

thirdTotal = "Total"
apptTotal.Formula = "=SUM(B" & thirdTitleRow + 5 & ":B" & thirdLast & ")"
exptTotal.Formula = "=SUM(C" & thirdTitleRow + 5 & ":C" & thirdLast & ")"
enctTotal.Formula = "=SUM(D" & thirdTitleRow + 5 & ":D" & thirdLast & ")"
freetTotal.Formula = "=SUM(E" & thirdTitleRow + 5 & ":E" & thirdLast & ")"
cashtTotal.Formula = "=SUM(F" & thirdTitleRow + 5 & ":F" & thirdLast & ")"

thirdSpent.Formula = "=(" & exptTotal & "+" & enctTotal & ")/" & apptTotal
thirdTotal.Offset(-1, 7) = "% Spent"
thirdTotal.Offset(-2, 7) = curFY


'Declare Steam variables

Dim steamArray As Variant
Dim rngsteamArr As Range
Dim steamcritLROW As Long
Dim steamTitleRow As Long
Dim steamTitle As Range

steamcritLROW = venarrWS.Cells(Rows.Count, 4).End(xlUp).Row
Set rngsteamArr = venarrWS.Range("D2:D" & steamcritLROW)
steamArray = rngsteamArr.Value
steamTitleRow = Vend.Cells(Rows.Count, 1).End(xlUp).Row
Set steamTitle = Vend.Range("A" & steamTitleRow).Offset(4)

'Copy Steam Kits over

AllLib.Range("B2").AutoFilter Field:=2, Criteria1:=Application.Transpose(steamArray), Operator:=xlFilterValues
AllLib.Range("B2:G50").SpecialCells(xlCellTypeVisible).Copy

steamTitle.PasteSpecial
steamTitle = "Steam/55121"
steamTitle.Interior.Color = vbYellow

Dim steamLast As Long
steamLast = Vend.Cells(Rows.Count, 1).End(xlUp).Row

Dim steamTotal As Range
Set steamTotal = Vend.Range("A" & steamLast).Offset(2)

Dim appsTotal As Range
Dim expsTotal As Range
Dim encsTotal As Range
Dim freesTotal As Range
Dim cashsTotal As Range
Dim steamSpent As Range

Set appsTotal = steamTotal.Offset(, 1)
Set expsTotal = steamTotal.Offset(, 2)
Set encsTotal = steamTotal.Offset(, 3)
Set freesTotal = steamTotal.Offset(, 4)
Set cashsTotal = steamTotal.Offset(, 5)
Set steamSpent = steamTotal.Offset(, 7)

steamTotal = "Total"
appsTotal.Formula = "=SUM(B" & steamTitleRow + 5 & ":B" & steamLast & ")"
expsTotal.Formula = "=SUM(C" & steamTitleRow + 5 & ":C" & steamLast & ")"
encsTotal.Formula = "=SUM(D" & steamTitleRow + 5 & ":D" & steamLast & ")"
freesTotal.Formula = "=SUM(E" & steamTitleRow + 5 & ":E" & steamLast & ")"
cashsTotal.Formula = "=SUM(F" & steamTitleRow + 5 & ":F" & steamLast & ")"

steamSpent.Formula = "=(" & expsTotal & "+" & encsTotal & ")/" & appsTotal
steamTotal.Offset(-1, 7) = "% Spent"
steamTotal.Offset(-2, 7) = curFY

AllLib.AutoFilterMode = False

End Sub
