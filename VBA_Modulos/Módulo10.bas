Attribute VB_Name = "Módulo10"
Option Explicit

Public Sub GuardarDatos()
    Dim calcPrev As XlCalculation
    On Error GoTo manejar_error

    ' Optimiza Excel durante la carga
    Application.ScreenUpdating = False
    calcPrev = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "Guardando datos..."

    ' Referencias a hojas y tablas
    Dim wsTemporal As Worksheet
    Dim wsPreos As Worksheet
    Dim wsATC_ANP_VAD As Worksheet
    Dim wsIC_JM As Worksheet

    Set wsTemporal = ThisWorkbook.Sheets("Formulario")
    Set wsPreos = ThisWorkbook.Sheets("PREOS Y ADICIONALES")
    On Error Resume Next
    Set wsATC_ANP_VAD = ThisWorkbook.Sheets("ATC ANP y VAD")
    If wsATC_ANP_VAD Is Nothing Then Set wsATC_ANP_VAD = ThisWorkbook.Sheets("ATC ANP Y VAD")
    On Error GoTo manejar_error
    Set wsIC_JM = ThisWorkbook.Sheets("IC y JM")

    Dim tbTemp As ListObject
    Set tbTemp = wsTemporal.ListObjects("tabla_chequeo")

    ' Constantes de columnas (tabla temporal)
    Const T_ID_SOL As Long = 1
    Const T_ID_PED As Long = 2
    Const T_ID_ATE As Long = 3
    Const T_FECHA As Long = 4
    Const T_VEND As Long = 5
    Const T_NOMCOM As Long = 6
    Const T_RSOC As Long = 7
    Const T_SUBCLI As Long = 8
    Const T_COD_EXA As Long = 9
    Const T_TIPO As Long = 10
    Const T_EXAMEN As Long = 11
    Const T_PUESTO As Long = 12
    Const T_PROV As Long = 13
    Const T_LOC As Long = 14
    Const T_SEDE As Long = 15
    Const T_COSTO As Long = 16
    Const T_VENTA As Long = 17
    Const T_COLAB As Long = 18
    Const T_DNI As Long = 19
    Const T_TEL As Long = 20
    Const T_HORA As Long = 21

    ' Variables de trabajo
    Dim filaTemp As ListRow
    Dim tblDef As ListObject
    Dim nuevaFila As ListRow
    Dim ultima As Long

    ' Variables para IC_JM
    Dim examenCompleto As String
    Dim examen As String, especialidad As String, modalidad As String



    For Each filaTemp In tbTemp.ListRows

        examenCompleto = NzText(filaTemp.Range(1, T_EXAMEN).Value)
        examen = DetectarExamen(examenCompleto)                 ' "Junta Medica" / "Interconsulta" / ""
        especialidad = DetectarEspecialidad(examenCompleto)     ' Clinica, Cardiologica, etc.
        modalidad = DetectarModalidad(examenCompleto)           ' Presencial / Virtual / ""

        If NzText(filaTemp.Range(1, T_TIPO).Value) = "Preocupacional" Then
            Set tblDef = wsPreos.ListObjects("Preos_Adicionales")

        ElseIf NzText(filaTemp.Range(1, T_TIPO).Value) = "Ausentismo" _
           And (NzText(filaTemp.Range(1, T_COD_EXA).Value) = "ATC" _
             Or NzText(filaTemp.Range(1, T_COD_EXA).Value) = "ANP" _
             Or NzText(filaTemp.Range(1, T_COD_EXA).Value) = "VAD") Then
            Set tblDef = wsATC_ANP_VAD.ListObjects("ATC_ANP_VAD")

        Else
            Set tblDef = wsIC_JM.ListObjects("IC_JM")
        End If

        Set nuevaFila = tblDef.ListRows.Add

        ' Asignaciones comunes (IDs y fecha)
        nuevaFila.Range(1, 1).Value = filaTemp.Range(1, T_ID_SOL).Value
        nuevaFila.Range(1, 2).Value = filaTemp.Range(1, T_ID_PED).Value
        nuevaFila.Range(1, 3).Value = filaTemp.Range(1, T_ID_ATE).Value

        With nuevaFila.Range(1, 5)
            .Value = CLng(filaTemp.Range(1, T_FECHA).Value)
            .NumberFormat = "dd/mm/yyyy"
            .Formula = .Value
        End With

        ' Mapeos especificos por tabla
        If tblDef.Name = "ATC_ANP_VAD" Then
            nuevaFila.Range(1, 8).Value = filaTemp.Range(1, T_VEND).Value
            nuevaFila.Range(1, 9).Value = filaTemp.Range(1, T_NOMCOM).Value
            nuevaFila.Range(1, 10).Value = filaTemp.Range(1, T_RSOC).Value
            nuevaFila.Range(1, 11).Value = filaTemp.Range(1, T_SUBCLI).Value
            nuevaFila.Range(1, 12).Value = filaTemp.Range(1, T_COD_EXA).Value
            nuevaFila.Range(1, 13).Value = filaTemp.Range(1, T_TIPO).Value
            nuevaFila.Range(1, 14).Value = filaTemp.Range(1, T_EXAMEN).Value
            nuevaFila.Range(1, 15).Value = filaTemp.Range(1, T_PROV).Value
            nuevaFila.Range(1, 16).Value = filaTemp.Range(1, T_LOC).Value
            nuevaFila.Range(1, 17).Value = filaTemp.Range(1, T_SEDE).Value
            'nuevaFila.Range(1, 18).Value = filaTemp.Range(1, T_COSTO).Value
            'nuevaFila.Range(1, 19).Value = filaTemp.Range(1, T_VENTA).Value
            nuevaFila.Range(1, 22).Value = filaTemp.Range(1, T_COLAB).Value
            nuevaFila.Range(1, 23).Value = filaTemp.Range(1, T_DNI).Value
            nuevaFila.Range(1, 24).Value = filaTemp.Range(1, T_TEL).Value
            nuevaFila.Range(1, 37).Value = filaTemp.Range(1, T_HORA).Value

        ElseIf tblDef.Name = "IC_JM" Then
            nuevaFila.Range(1, 9).Value = filaTemp.Range(1, T_VEND).Value
            nuevaFila.Range(1, 10).Value = filaTemp.Range(1, T_NOMCOM).Value
            nuevaFila.Range(1, 11).Value = filaTemp.Range(1, T_RSOC).Value
            nuevaFila.Range(1, 12).Value = filaTemp.Range(1, T_SUBCLI).Value
            nuevaFila.Range(1, 13).Value = filaTemp.Range(1, T_COD_EXA).Value
            nuevaFila.Range(1, 14).Value = filaTemp.Range(1, T_TIPO).Value
            nuevaFila.Range(1, 15).Value = examen
            nuevaFila.Range(1, 16).Value = especialidad
            nuevaFila.Range(1, 17).Value = modalidad
            nuevaFila.Range(1, 18).Value = filaTemp.Range(1, T_PROV).Value
            nuevaFila.Range(1, 19).Value = filaTemp.Range(1, T_LOC).Value
            'nuevaFila.Range(1, 20).Value = filaTemp.Range(1, T_SEDE).Value
            'nuevaFila.Range(1, 21).Value = filaTemp.Range(1, T_COSTO).Value
            'nuevaFila.Range(1, 22).Value = filaTemp.Range(1, T_VENTA).Value
            nuevaFila.Range(1, 25).Value = filaTemp.Range(1, T_COLAB).Value
            nuevaFila.Range(1, 26).Value = filaTemp.Range(1, T_DNI).Value
            nuevaFila.Range(1, 27).Value = filaTemp.Range(1, T_TEL).Value
            nuevaFila.Range(1, 38).Value = filaTemp.Range(1, T_HORA).Value

        Else
            ' Preos_Adicionales
            nuevaFila.Range(1, 9).Value = filaTemp.Range(1, T_VEND).Value
            nuevaFila.Range(1, 10).Value = filaTemp.Range(1, T_NOMCOM).Value
            nuevaFila.Range(1, 11).Value = filaTemp.Range(1, T_RSOC).Value
            nuevaFila.Range(1, 12).Value = filaTemp.Range(1, T_SUBCLI).Value
            nuevaFila.Range(1, 15).Value = filaTemp.Range(1, T_COD_EXA).Value
            nuevaFila.Range(1, 16).Value = filaTemp.Range(1, T_TIPO).Value
            nuevaFila.Range(1, 17).Value = filaTemp.Range(1, T_EXAMEN).Value
            nuevaFila.Range(1, 19).Value = filaTemp.Range(1, T_PUESTO).Value
            nuevaFila.Range(1, 20).Value = filaTemp.Range(1, T_PROV).Value
            nuevaFila.Range(1, 21).Value = filaTemp.Range(1, T_LOC).Value
            nuevaFila.Range(1, 22).Value = filaTemp.Range(1, T_SEDE).Value
            nuevaFila.Range(1, 23).Value = filaTemp.Range(1, T_COSTO).Value
            nuevaFila.Range(1, 24).Value = filaTemp.Range(1, T_VENTA).Value
            nuevaFila.Range(1, 13).Value = filaTemp.Range(1, T_COLAB).Value
            nuevaFila.Range(1, 14).Value = filaTemp.Range(1, T_DNI).Value
            nuevaFila.Range(1, 32).Value = filaTemp.Range(1, T_HORA).Value
        End If

        ' --- Marca "ASIGNAR" en la col 4 de la ultima fila de la tabla destino ---
        ultima = tblDef.ListRows.Count
        tblDef.DataBodyRange.Cells(ultima, 4).Value = "ASIGNAR"
    Next filaTemp

    ' Limpieza: borrar tabla temporal
    If Not tbTemp.DataBodyRange Is Nothing Then
        On Error Resume Next
        tbTemp.DataBodyRange.Delete
        On Error GoTo manejar_error
    End If

    ' (Opcional) macros de limpieza
    'Call limpiar_zonas
    'Call limpieza_codigo
    'Call limpiar_rentabilidad
    'Call limpiar_perfil_armado
    'Call Limpiar_ingreso_manual
    'Call Limpiar_nombreComercial

    MsgBox "Los datos se han guardado correctamente.", vbInformation
    Application.Goto wsTemporal.Range("A1"), True

finalizar:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.Calculation = calcPrev
    Application.ScreenUpdating = True
    Exit Sub

manejar_error:
    MsgBox "Ocurrio un error en GuardarDatos: " & Err.Number & " - " & Err.Description, vbCritical
    Resume finalizar
End Sub


' Auxiliares

' Devuelve cadena en minusculas, sin espacios repetidos
Private Function NormalizarTexto(ByVal s As String) As String
    NormalizarTexto = LCase$(Application.WorksheetFunction.Trim(Trim$(CStr(s))))
End Function

Private Function NzText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or v = "" Then
        NzText = ""
    Else
        NzText = CStr(v)
    End If
End Function

Private Function DetectarExamen(ByVal examenCompleto As String) As String
    Dim t As String: t = NormalizarTexto(examenCompleto)
    If InStr(1, t, "junta medica", vbTextCompare) > 0 Then
        DetectarExamen = "Junta Medica"
    ElseIf InStr(1, t, "interconsulta", vbTextCompare) > 0 Then
        DetectarExamen = "Interconsulta"
    Else
        DetectarExamen = ""
    End If
End Function

Private Function DetectarModalidad(ByVal examenCompleto As String) As String
    Dim t As String: t = NormalizarTexto(examenCompleto)
    If InStr(1, t, "presencial", vbTextCompare) > 0 Then
        DetectarModalidad = "Presencial"
    ElseIf InStr(1, t, "virtual", vbTextCompare) > 0 Then
        DetectarModalidad = "Virtual"
    Else
        DetectarModalidad = ""
    End If
End Function

Private Function DetectarEspecialidad(ByVal examenCompleto As String) As String
    Dim t As String: t = NormalizarTexto(examenCompleto)
    Dim esp() As String
    Dim i As Long

    esp = Split("clinica,cardiologica,psiquiatrica,traumatologica,otorrino,neurologica,psiquiatrica + clinica,traumatologica + clinica,traumatologica + neurologica", ",")

    For i = LBound(esp) To UBound(esp)
        If InStr(1, t, Trim$(esp(i)), vbTextCompare) > 0 Then
            DetectarEspecialidad = WorksheetFunction.Proper(Trim$(esp(i)))
            Exit Function
        End If
    Next i

    DetectarEspecialidad = ""
End Function


