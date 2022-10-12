Attribute VB_Name = "Separate"
Sub SectionOut()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim ca4Sheet As Range
Dim ca5Sheet As Range
Dim ca6Sheet As Range
Dim ca7Sheet As Range
Dim ca8Sheet As Range
Dim ca9Sheet As Range
Dim JuvSheet As Range
Dim YASheet As Range
Dim GroundSheet As Range
Dim L1Sheet As Range
Dim MezzSheet As Range
Dim StoneSheet As Range
Dim TwoSheet As Range

Set ca4Sheet = ca4.Range("A1")
Set ca5Sheet = ca5.Range("A1")
Set ca6Sheet = ca6.Range("A1")
Set ca7Sheet = ca7.Range("A1")
Set ca8Sheet = ca8.Range("A1")
Set ca9Sheet = ca9.Range("A1")
Set JuvSheet = Juv.Range("A1")
Set YASheet = YA.Range("A1")
Set GroundSheet = Ground.Range("A1")
Set L1Sheet = Lower1.Range("A1")
Set MezzSheet = Mezz.Range("A1")
Set StoneSheet = Stone.Range("A1")
Set TwoSheet = Two.Range("A1")

Dim LastRow As Long
    
    LastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

Dim CopyRange As Range
Set CopyRange = Full.Range("A1:K" & LastRow)

Dim LocColumn As Range
Set LocColumn = Full.Columns("G")

Dim SectionColumn As Range
Set SectionColumn = Full.Columns("L")

LocColumn.AutoFilter Field:=1, Criteria1:="ca4*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy ca4Sheet

LocColumn.AutoFilter Field:=1, Criteria1:="ca5*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy ca5Sheet

LocColumn.AutoFilter Field:=1, Criteria1:="ca6*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy ca6Sheet

LocColumn.AutoFilter Field:=1, Criteria1:="ca7*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy ca7Sheet

LocColumn.AutoFilter Field:=1, Criteria1:="ca8*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy ca8Sheet

LocColumn.AutoFilter Field:=1, Criteria1:="ca9*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy ca9Sheet

LocColumn.AutoFilter

SectionColumn.AutoFilter Field:=1, Criteria1:="Juv", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy JuvSheet

SectionColumn.AutoFilter Field:=1, Criteria1:="YA", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy YASheet

SectionColumn.AutoFilter Field:=1, Criteria1:="*Ground*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy GroundSheet

SectionColumn.AutoFilter Field:=1, Criteria1:="*Stone*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy StoneSheet

SectionColumn.AutoFilter Field:=1, Criteria1:="*2nd Floor*", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy TwoSheet

SectionColumn.AutoFilter Field:=1, Criteria1:="Mezz", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy MezzSheet

SectionColumn.AutoFilter Field:=1, Criteria1:="L1", Operator:=xlFilterValues
CopyRange.SpecialCells(xlCellTypeVisible).Copy L1Sheet

SectionColumn.AutoFilter

End Sub
