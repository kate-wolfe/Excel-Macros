Attribute VB_Name = "FormatTabs"
Sub Formatting()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Declare the last row of the range and other stuff

Dim LastRow As Long
    
    LastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

Dim SectionFormula As Variant
Dim ErrorFormula As Variant

Dim NonHeaderRange As Range
Set NonHeaderRange = Full.Range("A2:M" & LastRow)

'Declare the formula that decides the sections

SectionFormula = "=IF(ISNUMBER(SEARCH(""j"",G2))=TRUE,""Juv""," _
& "IF(ISNUMBER(SEARCH(""y"",G2))=TRUE,""YA""," _
& "IF(ISNUMBER(SEARCH(""al"",G2)),""Mezz""," _
& "IF(OR(I2=4,I2=5,I2=10)=TRUE,""Ground""," _
& "IF(OR(H2=130,H2=143,H2=144,AND(H2>=148,H2<=179),H2=220),""Mezz""," _
& "IF(OR(AND(H2>0,H2<100),H2=104,H2=109,H2=113,H2=114,H2=116,H2=117,H2=119,H2=139,H2=140,H2=141),""2nd Floor""," _
& "IF(OR(H2=102,H2=103,AND(H2=106,OR(LEFT(D2,7)=""MYSTERY"",LEFT(D2,7)=""SCI FIC""))),""Stone""," _
& "IF(OR(H2=106,LEFT(D2,7)=""FICTION""),""L1""," _
& "IF(OR(H2=101,H2=107,H2=108,H2=121,H2=122,H2=124),""L1"",""Other"")))))))))"


'Declare the formula that finds errors in the Location and IType (and sometimes SCAT)

ErrorFormula = "=IF(AND(I2=0,OR(G2=""cama"",G2=""ca4a"",G2=""ca5a"",G2=""ca6a"",G2=""ca7a"",G2=""ca8a"",G2=""ca9a"")),""Ok""," _
& "IF(AND(I2=1,ISNUMBER(SEARCH(""ap"",G2))=TRUE,L2=""L1""),""Ok""," _
& "IF(AND(I2=2,ISNUMBER(SEARCH(""al"",G2))=TRUE),""Ok""," _
& "IF(AND(I2=3,OR(G2=""camr"",G2=""camh"",G2=""camc""),H2=139),""Ok""," _
& "IF(AND(I2=6,G2=""ca3al""),""Ok""," _
& "IF(AND(I2=7,H2=115),""Ok""," _
& "IF(AND(OR(I2=4,I2=5),OR(G2=""caman"",G2=""camas"",G2=""ca4a"",G2=""ca5a"",G2=""ca6a"",G2=""ca7a"",G2=""ca8a"",G2=""ca9a"")),""Ok""," _
& "IF(AND(I2=9,ISNUMBER(SEARCH(""ae"",G2))=TRUE,L2=""Mezz""),""Ok""," _
& "IF(AND(I2=10,OR(G2=""cama"",G2=""ca4a"",G2=""ca5a"",G2=""ca6a"",G2=""ca7a"",G2=""ca8a"",G2=""ca9a"")),""Ok""," _
& "IF(AND(I2=12,H2=116,OR(G2=""cama"",G2=""ca4a"",G2=""ca5a"",G2=""ca6a"",G2=""ca7a"",G2=""ca8a"",G2=""ca9a"")),""Ok""," _
& "IF(AND(I2=51,H2=202,OR(G2=""camn"",G2=""ca4n"",G2=""ca5n"",G2=""ca6n"",G2=""ca7n"",G2=""ca8n"",G2=""ca9n"")),""Ok""," _
& "IF(AND(I2>=19,I2<=52,OR(G2=""camn"",G2=""ca4n"",G2=""ca5n"",G2=""ca6n"",G2=""ca7n"",G2=""ca8n"",G2=""ca9n""),L2=""Mezz""),""Ok""," _
& "IF(AND(I2>=100,I2<=133,ISNUMBER(SEARCH(""y"",G2))=TRUE),""Ok""," _
& "IF(AND(I2>=150,I2<=183,ISNUMBER(SEARCH(""j"",G2))=TRUE),""Ok"",""Error""))))))))))))))"


'General formatting and filling down the formulas

With Full
.Range("A1").FormulaR1C1 = "FND"
.Range("B1").FormulaR1C1 = "NOS"
.Range("C1").FormulaR1C1 = "Barcode"
.Range("D1").FormulaR1C1 = "Call #"
.Range("E1").FormulaR1C1 = "Title"
.Range("F1").FormulaR1C1 = "Date"
.Range("G1").FormulaR1C1 = "Loc"
.Range("H1").FormulaR1C1 = "SCAT"
.Range("I1").FormulaR1C1 = "IType"
.Range("J1").FormulaR1C1 = "Status"
.Range("K1").FormulaR1C1 = "Msg"
.Range("L1").FormulaR1C1 = "Section"
.Range("M1").FormulaR1C1 = "Errors"
.Range("A1:M1").Interior.ThemeColor = xlThemeColorLight1
.Range("A1:M1").Font.ThemeColor = xlThemeColorDark1
.Columns("C:C").NumberFormat = "0"
.Columns("H:I").NumberFormat = "0"
.Columns("G:G").Replace " ", "", xlPart
.Range("L2:L" & LastRow).Formula = SectionFormula
.Range("M2:M" & LastRow).Formula = ErrorFormula
.Range("A1:M" & LastRow).Borders.LineStyle = xlContinuous
.Range("A1:M" & LastRow).Font.Size = "10"
End With


'Shorten the call numbers

With Full.Columns("D:D")
.Replace "[Home & Health]", "[H&H]", xlPart
.Replace "CD CLASSICAL", "CD CLASS", xlPart
.Replace "CD ROCK", "CD POP", xlPart
.Replace "CD FOLK", "CD POP", xlPart
.Replace "CD SNDTRK", "CD POP", xlPart
.Replace "CD COUNTRY", "CD POP", xlPart
.Replace "FICTION", "FIC", xlPart
.Replace "CDB MYSTERY", "CDB FIC", xlPart
.Replace "CDB SCI FIC", "CDB FIC", xlPart
.Replace "SCI FIC", "SCIFI", xlPart
.Replace "MYSTERY", "MYS", xlPart
.Replace "[Business]", "[Biz]", xlPart
.Replace "[Great Courses]", "[GC]", xlPart
.Replace "MP3 ", "CDB (MP3)", xlPart
.Replace "[Express View]", "[Exp]", xlPart
.Replace "[Express", "[Exp", xlPart
End With


'Bold new books, italicize branch items, and highlight errors

Full.Columns("I").AutoFilter Field:=1, Criteria1:=Array("4", "5"), Operator:=xlFilterValues
NonHeaderRange.SpecialCells(xlCellTypeVisible).Font.Bold = True
Full.Columns("I").AutoFilter

Full.Columns("G").AutoFilter Field:=1, Criteria1:="<>*cam*", Operator:=xlFilterValues
NonHeaderRange.SpecialCells(xlCellTypeVisible).Font.Italic = True
Full.Columns("G").AutoFilter

Full.Columns("M").AutoFilter Field:=1, Criteria1:="Error", Operator:=xlFilterValues
NonHeaderRange.SpecialCells(xlCellTypeVisible).Interior.Color = 65535
Full.Columns("M").AutoFilter


End Sub
