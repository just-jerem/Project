Sub Test()
Dim Plage As Range
Dim column As Integer
Dim Char As String
Dim Sheet As String
    Sheet = "Sheet1"
    Char = "ok"
    column = 6
 
    With ThisWorkbook.Worksheets(Sheet)
        ' plage des données
        Set Plage = .Cells(2, 1).Resize(.UsedRange.Rows.Count - 1, .UsedRange.Columns.Count)
 
        With .Range("A1")
            ' retrait des filtres s'il y en a
            .AutoFilter
            ' application du filtre
            .AutoFilter  column, Char
            On Error Resume Next
                ' tentative de suppression des résultats
                Plage.SpecialCells(xlCellTypeVisible).EntireRow.Delete
                ' s'il n'y avait pas de résultat : on l'indique
                If Err <> 0 Then MsgBox "Pas de résultat"
            On Error GoTo 0
            ' suppression des filtres
            .AutoFilter
        End With
    End With
End Sub
