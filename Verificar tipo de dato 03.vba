'Declaración de una variable global que puede ser usada por todos los procedimientos Sub
Dim filMax As Long

Sub VerificarTipo(hoja As Worksheet, columna As Long)
    
    For fila = 2 To filMax
        With hoja.Cells(fila, columna)
            If Not IsNumeric(.Value) Then
                .Interior.Color = RGB(251, 226, 213)
            End If
        End With
    Next fila

End Sub

Sub verificar()

    filMax = ActiveSheet.UsedRange.Rows.Count

    VerificarTipo ActiveSheet, 1

End Sub
