Attribute VB_Name = "Módulo2"
Sub Limpiar_ingreso_manual()
Attribute Limpiar_ingreso_manual.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpiar_ingreso_manual Macro
'

'
    Range("limpieza_examenes.perfil").Select
    Selection.ClearContents
    Range("II_desplegable").Select
    
End Sub


Sub limpiar_rango_colaboradores()
' limpiar_rango_colaboradores Macro

    Worksheets("Formulario").Unprotect
    Range("rango_colaboradores").Select
    Selection.ClearContents
    Worksheets("Formulario").Protect
End Sub



Sub limpieza_codigo()

    Worksheets("Formulario").Unprotect
    
    Range("limpieza_codigo").Select
    Selection.ClearContents
    
    Worksheets("Formulario").Protect
End Sub



Sub limpiar_zonas()
' limpiar_zonas Macro

    Worksheets("Formulario").Unprotect

    Range("limpieza_zona").Select
    Selection.ClearContents
    Range("celda_provincia").Select

    Worksheets("Formulario").Protect
End Sub



Sub limpiar_rentabilidad()
'
' limpiar_rentabilidad Macro
'
    Worksheets("Formulario").Unprotect

    Range("limpieza_valores").Select
    Selection.ClearContents
    Range("celda_ventaManual").Select
    Selection.ClearContents
    Range("celda_costo").Select
    
    Worksheets("Formulario").Protect
End Sub




Sub Limpiar_nombreComercial()
'
' Limpiar_nombreComercial Macro
'

    Worksheets("Formulario").Unprotect
    Range("rango_cliente").Select
    Selection.ClearContents
    Range("celda_nombreComercial").Select
    Worksheets("Formulario").Protect
    
End Sub






Sub limpiar_perfil_armado()
'
' limpiar_perfil_armado Macro
'
    Worksheets("Formulario").Unprotect
    Range("perfiles_armados").Select
    Selection.ClearContents
    Worksheets("Formulario").Protect
End Sub





