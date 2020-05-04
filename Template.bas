Attribute VB_Name = "Module2"

Sub Template()

'Generate formatting
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Blank Count"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Fill Count"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Fill %"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "Distinct Count"
    Range("C1").Select
    Dim btch As Worksheet
    Set btch = Sheets.Add(After:=Sheets(Sheets.Count))
    btch.Name = "Batch Metadata"
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Layout"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = _
        "=INDIRECT(""'Batch Metadata'!B""&MATCH(R1C6,'Batch Metadata'!R[1]C[-2]:R[1048575]C[-2],0)+1)"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = _
        "=INDIRECT(""'Batch Metadata'!C""&MATCH(R1C6,'Batch Metadata'!R[1]C[-3]:R[1048575]C[-3],0)+1)"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = _
        "=INDIRECT(""'Batch Metadata'!D""&MATCH(R1C6,'Batch Metadata'!R[1]C[-4]:R[1048575]C[-4],0)+1)"
    Range("C1:F1").Select
    Selection.NumberFormat = "@"
    Range("E3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(TRIM(INDIRECT(RC[-1]&""!A2""))="""",INDIRECT(RC[-1]&""!B2""),0)"
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=SUM(INDIRECT(RC[-2]&""!B2:B1048576""))-RC[-1]"
    Range("G3").Select
    ActiveCell.Formula = "=F3/$E$1"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=COUNT(INDIRECT(RC[-4]&""!B2:B1048576""))"
    Sheets("Batch Metadata").Select
    Cells.Select
    Selection.NumberFormat = "@"
    Sheets("Layout").Select
    Range("A1:B2,C2:H2").Select
    Range("C2").Activate
    Selection.Style = "Note"
    Selection.Font.Bold = True
    Selection.Font.Underline = xlUnderlineStyleSingle
    With Selection.Font
        .Name = "Courier New"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleSingle
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A1:B1").Select
    Selection.Font.Size = 12
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("G3:G1048576").NumberFormat = "0.00%"

    
'Generate Sheets
    Dim rng As Range, cell As Range, x As Integer
    Set rng = Range("'Layout'!D3:D1000")
    x = 3
    For Each cell In rng
        If cell.Value <> "" Then
            Dim ws As Worksheet
            With ThisWorkbook
                Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
                ws.Name = cell.Value
                ws.Columns("A").NumberFormat = "@"
                ws.Cells(1, 1).Value = cell.Value
                ws.Cells(1, 2).Value = "Count"
                ws.ListObjects.Add(xlSrcRange, Range("A1:B2"), , xlYes).Name = cell.Value
                ws.Range("D1").Select
                ActiveCell.Formula = "Back"
                ws.hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
                "Layout!A1"
                x = x + 1
            End With
        End If
    Next cell
    
'Fill formulas
    x = x - 1
    Set SourceRange = Worksheets("Layout").Range("E3:H3")
    Set fillRange = Worksheets("Layout").Range("E3:H" & x)
    SourceRange.AutoFill Destination:=fillRange
End Sub

