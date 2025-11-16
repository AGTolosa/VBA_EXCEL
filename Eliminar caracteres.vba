Sub LimpiarCaracteres(hoja As Worksheet, fila As Long, columna As Long)
    For Each caracter In Array("-", " "):
        With hoja.Cells(fila, columna)
            .Value = Replace(CStr(.Value), caracter, "")
        End With
    Next caracter
End Sub

Sub Limpiar()

Dim Celda As Long
Dim filMax As Long
filMax = ActiveSheet.UsedRange.Rows.Count
    
    For Celda = 2 To filMax
        LimpiarCaracteres ActiveSheet, Celda, 2
    Next Celda

End Sub
