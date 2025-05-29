Attribute VB_Name = "Módulo8"
'Macro que actualiza los datos de producción de gas en SQL Server, filtrados por fecha en la hoja Menu-Inserción Diaria
Sub actualizarProducDiaria()
    Dim rng As Variant
    Dim uf As Long
    Dim cadenaUpdate As String
    Dim i As Long
    Dim batchSize As Integer
    Dim batchCount As Integer

    ' Determinar la última fila con datos en la hoja "Menu"
    uf = Sheets("Menu").Range("B" & Rows.Count).End(xlUp).Row
    ' Obtener el rango utilizado
    rng = Sheets("Menu").UsedRange
    batchSize = 500
    batchCount = 0

    Dim updateSQL As String

    For i = 19 To uf
        Dim idProduc As String
        Dim fechaValor As String
        Dim volumenProd As String
        Dim idCampo As String
        Dim fechaFormateada As String

        idProduc = rng(i, 2)    ' idProduc en columna 2
        fechaValor = rng(i, 3)   ' fecha en columna 3
        volumenProd = rng(i, 4)  ' volumen en columna 4
        idCampo = rng(i, 5)      ' idCampo en columna 5

        ' Validar que la fecha sea válida y formatearla
        If IsDate(fechaValor) Then
            fechaFormateada = Format(fechaValor, "yyyy-mm-dd")
        Else
            MsgBox "Fecha inválida en fila " & i
            Exit Sub
        End If

        ' Construir la sentencia UPDATE
        updateSQL = "UPDATE [ProdGas].[dbo].[produc_gas] SET " & _
                    "VolumenProd = '" & volumenProd & "', idCampo = '" & idCampo & "' " & _
                    "WHERE idProduc = " & idProduc & " AND fechaProd = '" & fechaFormateada & "';"

        ' Agregar a la cadena de UPDATE
        cadenaUpdate = cadenaUpdate & updateSQL & " "
        batchCount = batchCount + 1

        ' Ejecutar en lotes
        If batchCount = batchSize Or i = uf Then
            Dim finalSQL As String
            finalSQL = cadenaUpdate

            Debug.Print finalSQL

            ' Ejecutar SQL en la base de datos
            Set CONEXION = New CONEXION_DB
            Call IniciarDatos.IniciarDatos
            On Error GoTo ErrorHandler
            CONEXION.Ejecucion_SQL (finalSQL)
            On Error GoTo 0

            ' Resetear cadenas
            cadenaUpdate = ""
            batchCount = 0
        End If
    Next i

    MsgBox "Operación ACTUALIZAR completada."
    Exit Sub

ErrorHandler:
    MsgBox "Error en la ejecución de SQL: " & Err.Description
End Sub
'Macro que actualiza los datos de planes de producción de gas en SQL Server, filtrados por fecha en la hoja Menu-Inserción Diaria
Sub actualizarPlan()
    Dim rng As Variant
    Dim uf As Long
    Dim cadenaUpdate As String
    Dim i As Long
    Dim batchSize As Integer
    Dim batchCount As Integer

    ' Determinar la última fila con datos en la hoja "Menu"
    uf = Sheets("Menu").Range("G" & Rows.Count).End(xlUp).Row
    ' Obtener el rango utilizado
    rng = Sheets("Menu").UsedRange
    batchSize = 500
    batchCount = 0

    Dim updateSQL As String

    For i = 19 To uf
        Dim idPlan As String
        Dim fechaValor As String
        Dim volumenPlan As String
        Dim idArea As String
        Dim fechaFormateada As String

        idPlan = rng(i, 7)    ' idPlan en columna 7
        fechaValor = rng(i, 8)   ' fecha en columna 8
        volumenPlan = rng(i, 9)  ' volumen en columna 9
        idArea = rng(i, 10)      ' idArea en columna 10

        ' Validar que la fecha sea válida y formatearla
        If IsDate(fechaValor) Then
            fechaFormateada = Format(fechaValor, "yyyy-mm-dd")
        Else
            MsgBox "Fecha inválida en fila " & i
            Exit Sub
        End If

        ' Construir la sentencia UPDATE
        updateSQL = "UPDATE [ProdGas].[dbo].[planes_prod] SET " & _
                    "volumenPlan = '" & volumenPlan & "', idArea = '" & idArea & "' " & _
                    "WHERE idPlan = " & idPlan & " AND fechaPlan = '" & fechaFormateada & "';"

        ' Agregar a la cadena de UPDATE
        cadenaUpdate = cadenaUpdate & updateSQL & " "
        batchCount = batchCount + 1

        ' Ejecutar en lotes
        If batchCount = batchSize Or i = uf Then
            Dim finalSQL As String
            finalSQL = cadenaUpdate

            Debug.Print finalSQL

            ' Ejecutar SQL en la base de datos
            Set CONEXION = New CONEXION_DB
            Call IniciarDatos.IniciarDatos
            On Error GoTo ErrorHandler
            CONEXION.Ejecucion_SQL (finalSQL)
            On Error GoTo 0

            ' Resetear cadenas
            cadenaUpdate = ""
            batchCount = 0
        End If
    Next i

    MsgBox "Operación ACTUALIZAR completada."
    Exit Sub

ErrorHandler:
    MsgBox "Error en la ejecución de SQL: " & Err.Description
End Sub


