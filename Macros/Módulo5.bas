Attribute VB_Name = "Módulo5"
'Macro que limpia los datos de producción de gas filtrados por fecha en la hoja Menu-Inserción Diaria
Sub LimpiarPROD()
Attribute LimpiarPROD.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpiar Macro
'

    Range("B19").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub
'Macro que filtra los datos de producción de gas por fecha en la hoja Menu-Inserción Diaria
Sub FiltrarPROD()
Attribute FiltrarPROD.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Filtrar Macro
'

    Range("BA19").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(ProducGas,(FechaProd>=R14C3)*(FechaProd<=R14C4),""No Existe"")"
    Range("BA19").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("B19").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B19").Select
    Application.CutCopyMode = False
End Sub
'Macro que limpia los datos de planes de producción filtrados por fecha en la hoja Menu-Inserción Diaria
Sub LimpiarPLAN()
'
' Limpiar Macro
'

    Range("G19").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
End Sub
'Macro que filtra los datos de planes de producción por fecha en la hoja Menu-Inserción Diaria
Sub FiltrarPLAN()
'
' Filtrar Macro
'

    Range("BF19").Select
    ActiveCell.Formula2R1C1 = _
        "=FILTER(PlanesProd,(FechaPlan>=R14C8)*(FechaPlan<=R14C9),""No Existe"")"
    Range("BF19").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("G19").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("G19").Select
    Application.CutCopyMode = False
End Sub
