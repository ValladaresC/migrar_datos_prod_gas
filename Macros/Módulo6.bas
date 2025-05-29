Attribute VB_Name = "Módulo6"
'Macro que inserta o migra los datos de producción de gas desde excel hacia SQL Server filtrados por fecha en la hoja Menu-Inserción Diaria
Sub insertarProducDiaria()
    Dim rng As Variant
    Dim uf As Long 'para asegurar que en un futuro hayan muchas filas para insertar, cambiamos Integer por Long
    Dim cadena As String
    Dim i As Long 'para asegurar que en un futuro hayan muchas filas para insertar, cambiamos Integer por Long
    Dim batchSize As Integer
    Dim batchCount As Integer

    uf = Sheets("Menu").Range("B" & Rows.Count).End(xlUp).Row
    rng = Sheets("Menu").UsedRange
    batchSize = 500 ' Ajusta según el tamaño recomendado
    batchCount = 0
    cadena = ""

    For i = 19 To uf
        Dim valor1 As String
        Dim valor2 As String
        Dim valor3 As String
        Dim valor4 As String
        Dim fechaFormateada As String
        
        valor1 = rng(i, 2)
        valor2 = rng(i, 3)
        valor3 = rng(i, 4)
        valor4 = rng(i, 5)
        
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
'Macro que inserta o migra los datos de planes de producción desde excel hacia SQL Server filtrados por fecha en la hoja Menu-Inserción Diaria
Sub insertarPlanDiario()
    Dim rng As Variant
    Dim uf As Long 'para asegurar que en un futuro hayan muchas filas para insertar, cambiamos Integer por Long
    Dim cadena As String
    Dim i As Long 'para asegurar que en un futuro hayan muchas filas para insertar, cambiamos Integer por Long
    Dim batchSize As Integer
    Dim batchCount As Integer

    uf = Sheets("Menu").Range("G" & Rows.Count).End(xlUp).Row
    rng = Sheets("Menu").UsedRange
    batchSize = 500 ' Ajusta según el tamaño recomendado
    batchCount = 0
    cadena = ""

    For i = 19 To uf
        Dim valor1 As String
        Dim valor2 As String
        Dim valor3 As String
        Dim valor4 As String
        Dim fechaFormateada As String
        
        valor1 = rng(i, 7)
        valor2 = rng(i, 8)
        valor3 = rng(i, 9)
        valor4 = rng(i, 10)
        
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

