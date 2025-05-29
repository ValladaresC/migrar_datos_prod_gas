Attribute VB_Name = "M�dulo3"
'Macro que inserta las �reas desde excel hacia SQL Server en la hoja Menu-Inserci�n Total
Sub insertarAreas()

    Dim rng As Variant
    Dim uf As Integer
    Dim cadena As String
    Dim i As Integer
    
    uf = Hoja4.Range("A" & Rows.Count).End(xlUp).Row
    rng = Hoja4.UsedRange
    
    cadena = ""
    
    For i = 2 To uf
        
        cadena = cadena & "(" & "'" & rng(i, 1) & "'" & "," & "'" & rng(i, 2) & "'" & " )" & ","
        
    Next

    cadena = Left(cadena, Len(cadena) - 1)

    Set CONEXION = New CONEXION_DB
    sql = "INSERT INTO [ProdGas].[dbo].[areas]" _
    & "VALUES" & cadena
        
    Call IniciarDatos.IniciarDatos
    
    CONEXION.Ejecucion_SQL (sql)
    
    MsgBox "Operaci�n realizada con �xito."

End Sub
'Macro que inserta los Campos desde excel hacia SQL Server en la hoja Menu-Inserci�n Total
Sub insertarCampos()

    Dim rng As Variant
    Dim uf As Integer
    Dim cadena As String
    Dim i As Integer
    
    uf = Hoja3.Range("A" & Rows.Count).End(xlUp).Row
    rng = Hoja3.UsedRange
    
    cadena = ""
    
    For i = 2 To uf
        
        cadena = cadena & "(" & "'" & rng(i, 1) & "'" & "," & "'" & rng(i, 2) & "'" & "," & "'" & rng(i, 3) & "'" & " )" & ","
        
    Next

    cadena = Left(cadena, Len(cadena) - 1)

    Set CONEXION = New CONEXION_DB
    sql = "INSERT INTO [ProdGas].[dbo].[campos]" _
    & "VALUES" & cadena
        
    Call IniciarDatos.IniciarDatos
    
    CONEXION.Ejecucion_SQL (sql)
    
    MsgBox "Operaci�n realizada con �xito."

End Sub

