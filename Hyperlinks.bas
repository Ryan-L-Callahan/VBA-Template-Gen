Attribute VB_Name = "Module1"
Sub hyperlinks()

    Dim rng As Range, cell As Range, x As Integer
    Set rng = Range("'Layout'!D3:D1000")
    x = 3
    For Each cell In rng
        If cell.Value <> "" Then
            cell.Select
            Worksheets("Layout").hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
            "Layout!A1"
        x = x + 1
        End If
    Next cell
End Sub



