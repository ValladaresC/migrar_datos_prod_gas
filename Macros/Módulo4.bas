Attribute VB_Name = "Módulo4"
'Macro que inserta los datos de Planes de Producción sin filtro desde excel hacia SQL Server en la hoja Menu-Inserción Total
Sub insertarPlanesProd()
    Dim rng As Variant
    Dim uf As Long 'para asegurar que en un futuro hayan muchas filas para insertar, cambiamos Integer por Long
    Dim cadena As String
    Dim i As Long 'para asegurar que en un futuro hayan muchas filas para insertar, cambiamos Integer por Long
    Dim batchSize As Integer
    Dim batchCount As Integer

    uf = Hoja2.Range("A" & Rows.Count).End(xlUp).Row
    rng = Hoja2.UsedRange
    batchSize = 500 ' Ajusta según el tamaño recomendado
    batchCount = 0
    cadena = ""

    For i = 2 To uf
        Dim valor1 As String
        Dim valor2 As String
        Dim valor3 As String
        Dim valor4 As String
        Dim fechaFormateada As String
        
        valor1 = rng(i, 1)
        valor2 = rng(i, 2)
        valor3 = rng(i, 3)
        valor4 = rng(i, 4)
        
        ' Formatear la fecha
        If IsDate(valor2) Then
            fechaFormateada = Format(valor2, "yyyy-mm-dd")
        Else
            MsgBox "Fecha inválida en fila " & i
            Exit Sub
        End If
        
        ' Construir la cadena de valores
        cadena = cadena & "(" & valor1 & ", '" & fechaFormateada & "', '" & valor3 & "', '" & valor4 & "'),"
        batchCount = batchCount + 1
        
        ' Si alcanzamos el tamaño del lote, ejecutamos la inserción
        If batchCount = batchSize Or i = uf Then
            ' Eliminar la última coma si no es el último lote
            If Len(cadena) > 0 Then
                cadena = Left(cadena, Len(cadena) - 1)
            End If
            
            ' Crear la sentencia SQL
            Dim sql As String
            sql = "INSERT INTO [ProdGas].[dbo].[planes_prod] VALUES " & cadena
            
            Debug.Print sql ' Debug: Imprime la consulta SQL
            
            ' Ejecutar la sentencia SQL
            Set CONEXION = New CONEXION_DB
            Call IniciarDatos.IniciarDatos
            On Error GoTo ErrorHandler
            CONEXION.Ejecucion_SQL (sql)
            On Error GoTo 0
            
            ' Reiniciar la cadena y el contador
            cadena = ""
            batchCount = 0
        End If
    Next i

    MsgBox "Operación realizada con éxito."
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en la ejecución de SQL: " & Err.Description
End Sub
'Macro que inserta los datos de Producción de Gas sin filtros de fechas desde excel hacia SQL Server en la hoja Menu-Inserción Total
Sub insertarProdGas()
    Dim rng As Variant
    Dim uf As Long 'fue cambiada de Integer a Long porque Long maneja valores mas grandes, en cambio Integer maneja hasta 32.767
    Dim cadena As String
    Dim i As Long 'fue cambiada de Integer a Long porque Long maneja valores mas grandes, en cambio Integer maneja hasta 32.767
    Dim batchSize As Integer
    Dim batchCount As Integer

    uf = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
    rng = Hoja1.UsedRange
    batchSize = 500 ' Ajusta según el tamaño recomendado
    batchCount = 0
    cadena = ""

    For i = 2 To uf
        Dim valor1 As String
        Dim valor2 As String
        Dim valor3 As String
        Dim valor4 As String
        Dim fechaFormateada As String
        
        valor1 = rng(i, 1)
        valor2 = rng(i, 2)
        valor3 = rng(i, 3)
        valor4 = rng(i, 4)
        
        ' Formatear la fecha
        If IsDate(valor2) Then
            fechaFormateada = Format(valor2, "yyyy-mm-dd")
        Else
            MsgBox "Fecha inválida en fila " & i
            Exit Sub
        End If
        
        ' Construir la cadena de valores
        cadena = cadena & "(" & valor1 & ", '" & fechaFormateada & "', '" & valor3 & "', '" & valor4 & "'),"
        batchCount = batchCount + 1
        
        ' Si alcanzamos el tamaño del lote, ejecutamos la inserción
        If batchCount = batchSize Or i = uf Then
            ' Eliminar la última coma si no es el último lote
            If Len(cadena) > 0 Then
                cadena = Left(cadena, Len(cadena) - 1)
            End If
            
            ' Crear la sentencia SQL
            Dim sql As String
            sql = "INSERT INTO [ProdGas].[dbo].[produc_gas] VALUES " & cadena
            
            Debug.Print sql ' Debug: Imprime la consulta SQL
            
            ' Ejecutar la sentencia SQL
            Set CONEXION = New CONEXION_DB
            Call IniciarDatos.IniciarDatos
            On Error GoTo ErrorHandler
            CONEXION.Ejecucion_SQL (sql)
            On Error GoTo 0
            
            ' Reiniciar la cadena y el contador
            cadena = ""
            batchCount = 0
        End If
    Next i

    MsgBox "Operación realizada con éxito."
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en la ejecución de SQL: " & Err.Description
End Sub

