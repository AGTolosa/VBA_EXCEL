Sub VerificarTipoDato()
    Dim hoja As Worksheet
    Dim fila As Long
    Dim filMax As Long
    Dim columna As Long
    
    Set hoja = ActiveSheet
    filMax = ActiveSheet.UsedRange.Rows.Count
    columna = 1
    
    For fila = 2 To filMax
        With hoja.Cells(fila, columna)
            If Not IsNumeric(.Value) Then
                .Interior.Color = RGB(251, 226, 213)
            End If
        End With
    Next fila
End Sub
