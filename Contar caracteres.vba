Sub ContarCaracteres()
    Dim filMax As Long
    Dim colMax As Long
    Dim fila As Long
    Dim columna As Long
    Dim MaxCaracteres As Long
    Dim MaxCelda As String
    Dim NumCaracteres As Long

    ' Inicializar variables
    MaxCaracteres = 0
    MaxCelda = ""

    ' Obtener el rango usado
    filMax = ActiveSheet.UsedRange.Rows.Count
    colMax = ActiveSheet.UsedRange.Columns.Count

    ' Recorrer todas las celdas usadas
    For columna = 1 To colMax
        For fila = 1 To filMax
            NumCaracteres = Len(ActiveSheet.Cells(fila, columna).Value)
            If NumCaracteres > MaxCaracteres Then
                MaxCaracteres = NumCaracteres
                MaxCelda = ActiveSheet.Cells(fila, columna).Address
            End If
        Next fila
    Next columna

    ' Mostrar el resultado
    MsgBox "La celda con el mayor número de caracteres es " & MaxCelda & " con " & MaxCaracteres & " caracteres.", vbInformation, "Resultado"
End Sub
