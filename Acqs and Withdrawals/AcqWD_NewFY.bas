Attribute VB_Name = "NewFY"
Option Explicit

Sub NewFiscalYear()

Application.ScreenUpdating = False

'Declare variables

Dim currentYear As Integer
Dim nextYear As Integer
Dim oldFY As Variant
Dim NewFY As Variant
Dim wbBook As Workbook
Dim wsYTD As Worksheet
Dim wsJULY As Worksheet
Dim wsAUG As Worksheet
Dim wsSEPT As Worksheet
Dim wsOCT As Worksheet
Dim wsNOV As Worksheet
Dim wsDEC As Worksheet
Dim wsJAN As Worksheet
Dim wsFEB As Worksheet
Dim wsMAR As Worksheet
Dim wsAPR As Worksheet
Dim wsMAY As Worksheet
Dim wsJUNE As Worksheet

Set wbBook = ThisWorkbook
With wbBook
    Set wsYTD = .Worksheets("YTD")
    Set wsJULY = .Worksheets("JULY")
    Set wsAUG = .Worksheets("AUGUST")
    Set wsSEPT = .Worksheets("SEPTEMBER")
    Set wsOCT = .Worksheets("OCTOBER")
    Set wsNOV = .Worksheets("NOVEMBER")
    Set wsDEC = .Worksheets("DECEMBER")
    Set wsJAN = .Worksheets("JANUARY")
    Set wsFEB = .Worksheets("FEBRUARY")
    Set wsMAR = .Worksheets("MARCH")
    Set wsAPR = .Worksheets("APRIL")
    Set wsMAY = .Worksheets("MAY")
    Set wsJUNE = .Worksheets("JUNE")

End With


'Set years

currentYear = Year(Now)
nextYear = Year(Now) + 1
oldFY = "FY" & Right(currentYear, 2)
NewFY = "FY" & Right(nextYear, 2)

wsYTD.Range("B6").Value = NewFY
wsYTD.Range("C6").Value = oldFY

wsJULY.Range("J2").Value = currentYear
wsJAN.Range("J2").Value = nextYear

'Move fiscal year to old spot and clear out new spot

Dim ws As Worksheet

For Each ws In Sheets
    If ws.Name <> "YTD" Then
        ws.Range("C8:C9").Value = ws.Range("B8:B9").Value
        ws.Range("E8:E9").Value = ws.Range("D8:D9").Value
        ws.Range("G8:G9").Value = ws.Range("F8:F9").Value
        ws.Range("I8:I9").Value = ws.Range("H8:H9").Value
        ws.Range("K8:K9").Value = ws.Range("J8:J9").Value
        ws.Range("M8:M9").Value = ws.Range("L8:L9").Value
        ws.Range("O8:O9").Value = ws.Range("N8:N9").Value
        ws.Range("C13:C14").Value = ws.Range("B13:B14").Value
        ws.Range("E13:E14").Value = ws.Range("D13:D14").Value
        ws.Range("G13:G14").Value = ws.Range("F13:F14").Value
        ws.Range("I13:I14").Value = ws.Range("H13:H14").Value
        ws.Range("K13:K14").Value = ws.Range("J13:J14").Value
        ws.Range("M13:M14").Value = ws.Range("L13:L14").Value
        ws.Range("O13:O14").Value = ws.Range("N13:N14").Value
        ws.Range("C25:C26").Value = ws.Range("B25:B26").Value
        ws.Range("E25:E26").Value = ws.Range("D25:D26").Value
        ws.Range("G25:G26").Value = ws.Range("F25:F26").Value
        ws.Range("I25:I26").Value = ws.Range("H25:H26").Value
        ws.Range("K25:K26").Value = ws.Range("J25:J26").Value
        ws.Range("M25:M26").Value = ws.Range("L25:L26").Value
        ws.Range("O25:O26").Value = ws.Range("N25:N26").Value
        ws.Range("C30:C31").Value = ws.Range("B30:B31").Value
        ws.Range("E30:E31").Value = ws.Range("D30:D31").Value
        ws.Range("G30:G31").Value = ws.Range("F30:F31").Value
        ws.Range("I30:I31").Value = ws.Range("H30:H31").Value
        ws.Range("K30:K31").Value = ws.Range("J30:J31").Value
        ws.Range("M30:M31").Value = ws.Range("L30:L31").Value
        ws.Range("O30:O31").Value = ws.Range("N30:N31").Value
        
        ws.Range("B8:B9,D8:D9,F8:F9,H8:H9,J8:J9,L8:L9,N8:N9,B13:B14,D13:D14,F13:F14,H13:H14,J13:J14,L13:L14,N13:N14,B25:B26,D25:D26,F25:F26,H25:H26,J25:J26,L25:L26,N25:N26,B30:B31,D30:D31,F30:F31,H30:H31,J30:J31,L30:L31,N30:N31").ClearContents

    End If
Next ws


End Sub


