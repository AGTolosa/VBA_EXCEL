Sub VerificarTipo(hoja As Worksheet, columna As Long)

Dim filMax As Long
filMax = ActiveSheet.UsedRange.Rows.Count
    
    For fila = 2 To filMax
        With hoja.Cells(fila, columna)
            If Not IsNumeric(.Value) Then
                .Interior.Color = RGB(251, 226, 213)
            End If
        End With
    Next fila

End Sub

Sub verificar()

VerificarTipo ActiveSheet, 1

End Sub
