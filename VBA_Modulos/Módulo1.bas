Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub AgregarDatos()

    Application.ScreenUpdating = False

    ' Referencias a hojas y tablas
    Dim shFormulario As Worksheet
    Dim shPreos As Worksheet
    Dim shATC_ANP_VAD As Worksheet
    Dim shIC_JM As Worksheet
    Dim shCodificacion As Worksheet

    Dim tbChequeo As ListObject
    Dim tbPreos As ListObject
    Dim tbATC As ListObject
    Dim tbICJM As ListObject

    Dim tbAusentismo As ListObject
    Dim tbExamenes As ListObject
    Dim tbDrogas As ListObject
    Dim tbLaboratorio As ListObject
    Dim tbPerfiles As ListObject

    Set shFormulario = ThisWorkbook.Sheets("Formulario")
    Set shPreos = ThisWorkbook.Sheets("PREOS Y ADICIONALES")
    Set shATC_ANP_VAD = ThisWorkbook.Sheets("ATC ANP Y VAD")
    Set shIC_JM = ThisWorkbook.Sheets("IC y JM")
    Set shCodificacion = ThisWorkbook.Sheets("conexion_CodificacionyPerfiles")

    Set tbChequeo = shFormulario.ListObjects("tabla_chequeo")
    Set tbPreos = shPreos.ListObjects("Preos_Adicionales")
    Set tbATC = shATC_ANP_VAD.ListObjects("ATC_ANP_VAD")
    Set tbICJM = shIC_JM.ListObjects("IC_JM")

    Set tbAusentismo = shCodificacion.ListObjects("tabla_ausentismo")
    Set tbExamenes = shCodificacion.ListObjects("tabla_examenes")
    Set tbDrogas = shCodificacion.ListObjects("tabla_drogas")
    Set tbLaboratorio = shCodificacion.ListObjects("tabla_laboratorio")
    Set tbPerfiles = shCodificacion.ListObjects("tabla_perfiles")

    ' Lectura de campos del Formulario
    Dim nombreCliente As String
    Dim vendedor As String
    Dim razonSocial As String
    Dim tipoServicio As String
    Dim puesto As String
    Dim localidad As String
    Dim sede As String
    Dim costo As Double
    Dim provincia As String
    Dim subCliente As String

    Dim fechaSolicitudNumero As Double
    Dim horaSolicitudTexto As String
    Dim cantidadIngresada As Long

    nombreCliente = shFormulario.Range("celda_nombreComercial").Value
    vendedor = shFormulario.Range("vendedor").Value
    razonSocial = shFormulario.Range("razon_social").Value
    tipoServicio = shFormulario.Range("tipo_servicio").Value
    puesto = shFormulario.Range("celda_puesto").Value
    localidad = shFormulario.Range("celda_localidad").Value
    sede = shFormulario.Range("sede").Value
    costo = shFormulario.Range("celda_costo").Value
    provincia = shFormulario.Range("celda_provincia").Value
    subCliente = shFormulario.Range("celda_SubCliente").Value

    fechaSolicitudNumero = CLng(Date)
    horaSolicitudTexto = Format(Time, "hh:mm")
    cantidadIngresada = Val(shFormulario.Range("celda_cantidad").Value)

    ' Validacion minimo
    If cantidadIngresada <= 0 Then
        MsgBox "La cantidad ingresada debe ser mayor a cero.", vbExclamation, "Validacion"
        GoTo salir
    End If
    If Len(Trim(nombreCliente)) = 0 Then
        MsgBox "Falta el Nombre del Cliente.", vbExclamation, "Validación"
        GoTo salir
    End If

    ' Calcular IDs
    Dim idSolicitud As Long, idPedido As Long, idAtencion As Long
    CalcularIDs tbChequeo, tbPreos, tbATC, tbICJM, nombreCliente, fechaSolicitudNumero, _
                idSolicitud, idPedido, idAtencion

    ' Precio de venta
    Dim precioVenta As Double
    precioVenta = TomarPrecioVenta(shFormulario)

    ' Resolver EXAMEN/PERFIL
    Dim examenTexto As String
    examenTexto = DeterminarExamenDesdeFormulario(shFormulario, tipoServicio)

    ' Buscar el codigo del examen/perfil
    Dim codigoExamen As String, codigoEncontrado As Boolean
    codigoExamen = ""
    codigoEncontrado = False

    BuscarCodigoExamen examenTexto, tbAusentismo, tbExamenes, tbDrogas, tbLaboratorio, tbPerfiles, _
                       codigoExamen, codigoEncontrado

    ' Insertar las filas en la tabla_chequeo
    Call AgregarFilasEnChequeo(tbChequeo, cantidadIngresada, _
                               idSolicitud, idPedido, idAtencion, _
                               fechaSolicitudNumero, horaSolicitudTexto, _
                               vendedor, nombreCliente, razonSocial, _
                               tipoServicio, puesto, provincia, localidad, sede, _
                               costo, precioVenta, examenTexto, subCliente, _
                               codigoEncontrado, codigoExamen)

    ' Volver a la celda
    Application.Goto shFormulario.Range("M1"), True

salir:
    Application.ScreenUpdating = True
End Sub

' FUNCIONES AUXILIARES

' Calcula ID_SOLICITUD, ID_PEDIDO, ID_ATENCION
Private Sub CalcularIDs( _
    ByVal tbChequeo As ListObject, _
    ByVal tbPreos As ListObject, _
    ByVal tbATC As ListObject, _
    ByVal tbICJM As ListObject, _
    ByVal nombreCliente As String, _
    ByVal fechaNumero As Double, _
    ByRef idSolicitud As Long, _
    ByRef idPedido As Long, _
    ByRef idAtencion As Long)

    Dim maxIdSolicitud As Long: maxIdSolicitud = 0
    Dim maxIdAtencion As Long: maxIdAtencion = 0
    Dim solicitudExiste As Boolean: solicitudExiste = False
    Dim maxIdPedidoCliente As Long: maxIdPedidoCliente = 0
    Dim u As Long

    ' ID_SOLICITUD = maximo de las 3 tablas + 1
    If tbPreos.ListRows.Count > 0 Then _
        maxIdSolicitud = WorksheetFunction.Max(tbPreos.DataBodyRange.Columns(1))
    If tbATC.ListRows.Count > 0 Then _
        maxIdSolicitud = WorksheetFunction.Max(maxIdSolicitud, WorksheetFunction.Max(tbATC.DataBodyRange.Columns(1)))
    If tbICJM.ListRows.Count > 0 Then _
        maxIdSolicitud = WorksheetFunction.Max(maxIdSolicitud, WorksheetFunction.Max(tbICJM.DataBodyRange.Columns(1)))

    idSolicitud = maxIdSolicitud + 1

    ' ID_PEDIDO: usa la misma solicitud si existe: mismo cliente y mismo mes
    ' PREOS (cliente col 10, fecha col 5, pedido col 2)
    For u = 1 To tbPreos.ListRows.Count
        If tbPreos.DataBodyRange.Cells(u, 10).Value = nombreCliente _
        And Month(tbPreos.DataBodyRange.Cells(u, 5).Value) = Month(fechaNumero) Then
            solicitudExiste = True
            idSolicitud = tbPreos.DataBodyRange.Cells(u, 1).Value
            maxIdPedidoCliente = Application.Max(maxIdPedidoCliente, tbPreos.DataBodyRange.Cells(u, 2).Value)
        End If
    Next u

    ' ATC_ANP_VAD (cliente col 9)
    For u = 1 To tbATC.ListRows.Count
        If tbATC.DataBodyRange.Cells(u, 9).Value = nombreCliente _
        And Month(tbATC.DataBodyRange.Cells(u, 5).Value) = Month(fechaNumero) Then
            solicitudExiste = True
            idSolicitud = tbATC.DataBodyRange.Cells(u, 1).Value
            maxIdPedidoCliente = Application.Max(maxIdPedidoCliente, tbATC.DataBodyRange.Cells(u, 2).Value)
        End If
    Next u

    ' IC_JM (cliente col 10)
    For u = 1 To tbICJM.ListRows.Count
        If tbICJM.DataBodyRange.Cells(u, 10).Value = nombreCliente _
        And Month(tbICJM.DataBodyRange.Cells(u, 5).Value) = Month(fechaNumero) Then
            solicitudExiste = True
            idSolicitud = tbICJM.DataBodyRange.Cells(u, 1).Value
            maxIdPedidoCliente = Application.Max(maxIdPedidoCliente, tbICJM.DataBodyRange.Cells(u, 2).Value)
        End If
    Next u

    If solicitudExiste Then
        idPedido = maxIdPedidoCliente + 1
    Else
        idPedido = 1
    End If

    ' ID_ATENCION = maximo (temporal + 3 definitivas) + 1
    If tbChequeo.ListRows.Count > 0 Then _
        maxIdAtencion = WorksheetFunction.Max(tbChequeo.DataBodyRange.Columns(3))
    If tbPreos.ListRows.Count > 0 Then _
        maxIdAtencion = WorksheetFunction.Max(maxIdAtencion, WorksheetFunction.Max(tbPreos.DataBodyRange.Columns(3)))
    If tbATC.ListRows.Count > 0 Then _
        maxIdAtencion = WorksheetFunction.Max(maxIdAtencion, WorksheetFunction.Max(tbATC.DataBodyRange.Columns(3)))
    If tbICJM.ListRows.Count > 0 Then _
        maxIdAtencion = WorksheetFunction.Max(maxIdAtencion, WorksheetFunction.Max(tbICJM.DataBodyRange.Columns(3)))

    idAtencion = maxIdAtencion + 1
End Sub

' Toma precio de venta
Private Function TomarPrecioVenta(ByVal sh As Worksheet) As Double
    If IsNumeric(sh.Range("celda_venta").Value) And sh.Range("celda_venta").Value > 0 Then
        TomarPrecioVenta = sh.Range("celda_venta").Value
    ElseIf IsNumeric(sh.Range("celda_ventaManual").Value) And sh.Range("celda_ventaManual").Value > 0 Then
        TomarPrecioVenta = sh.Range("celda_ventaManual").Value
    Else
        TomarPrecioVenta = 0
    End If
End Function

' Resuelve el texto del examen/perfil, priorizando "celda_examen_codigo" y despues ingreso manual.
Private Function DeterminarExamenDesdeFormulario(ByVal sh As Worksheet, ByVal tipoServicio As String) As String
    Dim examen As String
    examen = ""

    ' 1) Codigo directo
    If Trim(sh.Range("celda_examen_codigo").Value) <> "" Then
        examen = sh.Range("celda_examen_codigo").Value
    End If

    ' 2) Ingreso manual si no hay codigo
    If examen = "" Then
        If tipoServicio = "Ausentismo" Then
            examen = sh.Range("II_desplegable").Value
            If Trim(sh.Range("III_desplegable").Value) <> "" Then
                examen = examen & " " & sh.Range("III_desplegable").Value & " " & sh.Range("IIII_desplegable").Value
            End If
        ElseIf tipoServicio = "Preocupacional" And sh.Range("II_desplegable").Value = "Examen" Then
            If sh.Range("III_desplegable").Value = "Drogas" Or sh.Range("III_desplegable").Value = "Laboratorio" Then
                examen = sh.Range("drogas_laboratorio").Cells(1, 1).Value & sh.Range("drogas_laboratorio").Cells(1, 2).Value
            Else
                examen = sh.Range("III_desplegable").Value
            End If
        ElseIf tipoServicio = "Preocupacional" And sh.Range("II_desplegable").Value = "Perfil" Then
            examen = sh.Range("perfiles_armados").Value
        End If
    End If

    DeterminarExamenDesdeFormulario = Trim(examen)
End Function

' Busca el codigo del examen/perfil recorriendo las 5 tablas.
Private Sub BuscarCodigoExamen( _
    ByVal examen As String, _
    ByVal tbAusentismo As ListObject, _
    ByVal tbExamenes As ListObject, _
    ByVal tbDrogas As ListObject, _
    ByVal tbLaboratorio As ListObject, _
    ByVal tbPerfiles As ListObject, _
    ByRef codigo As String, _
    ByRef encontrado As Boolean)

    Dim j As Long, lastRow As Long

    codigo = ""
    encontrado = False
    If Len(Trim(examen)) = 0 Then Exit Sub

    ' 1) tabla_ausentismo: concatenar col 3 + 4 + 5 y comparar con examen
    lastRow = tbAusentismo.ListRows.Count
    For j = 1 To lastRow
        If Trim(tbAusentismo.DataBodyRange.Cells(j, 3).Value & " " & _
                 tbAusentismo.DataBodyRange.Cells(j, 4).Value & " " & _
                 tbAusentismo.DataBodyRange.Cells(j, 5).Value) = Trim(examen) Then
            codigo = tbAusentismo.DataBodyRange.Cells(j, 1).Value
            encontrado = True
            Exit Sub
        End If
    Next j

    ' 2) tabla_examenes: comparar col 4
    lastRow = tbExamenes.ListRows.Count
    For j = 1 To lastRow
        If Trim(tbExamenes.DataBodyRange.Cells(j, 4).Value) = Trim(examen) Then
            codigo = tbExamenes.DataBodyRange.Cells(j, 1).Value
            encontrado = True
            Exit Sub
        End If
    Next j

    ' 3) tabla_drogas: concatenar col 4 + 5
    lastRow = tbDrogas.ListRows.Count
    For j = 1 To lastRow
        If Trim(tbDrogas.DataBodyRange.Cells(j, 4).Value & " " & _
                 tbDrogas.DataBodyRange.Cells(j, 5).Value) = Trim(examen) Then
            codigo = tbDrogas.DataBodyRange.Cells(j, 1).Value
            encontrado = True
            Exit Sub
        End If
    Next j

    ' 4) tabla_laboratorio: concatenar col 4 + 5
    lastRow = tbLaboratorio.ListRows.Count
    For j = 1 To lastRow
        If Trim(tbLaboratorio.DataBodyRange.Cells(j, 4).Value & " " & _
                 tbLaboratorio.DataBodyRange.Cells(j, 5).Value) = Trim(examen) Then
            codigo = tbLaboratorio.DataBodyRange.Cells(j, 1).Value
            encontrado = True
            Exit Sub
        End If
    Next j

    ' 5) tabla_perfiles: comparar col 4
    lastRow = tbPerfiles.ListRows.Count
    For j = 1 To lastRow
        If Trim(tbPerfiles.DataBodyRange.Cells(j, 4).Value) = Trim(examen) Then
            codigo = tbPerfiles.DataBodyRange.Cells(j, 1).Value
            encontrado = True
            Exit Sub
        End If
    Next j
End Sub

' Agrega filas a la tabla de chequeo
Private Sub AgregarFilasEnChequeo( _
    ByVal tbChequeo As ListObject, _
    ByVal cantidad As Long, _
    ByVal idSolicitud As Long, _
    ByVal idPedido As Long, _
    ByVal idAtencionInicial As Long, _
    ByVal fechaNumero As Double, _
    ByVal horaTexto As String, _
    ByVal vendedor As String, _
    ByVal nombreCliente As String, _
    ByVal razonSocial As String, _
    ByVal tipoServicio As String, _
    ByVal puesto As String, _
    ByVal provincia As String, _
    ByVal localidad As String, _
    ByVal sede As String, _
    ByVal costo As Double, _
    ByVal precioVenta As Double, _
    ByVal examen As String, _
    ByVal subCliente As String, _
    ByVal codigoEncontrado As Boolean, _
    ByVal codigoExamen As String)

    Dim i As Long, idxUltima As Long, idAtencion As Long
    idAtencion = idAtencionInicial

    tbChequeo.Parent.Unprotect

    For i = 1 To cantidad
        tbChequeo.ListRows.Add
        idxUltima = tbChequeo.ListRows.Count

        ' Columnas
        tbChequeo.DataBodyRange.Cells(idxUltima, 1).Value = idSolicitud        ' ID_SOLICITUD
        tbChequeo.DataBodyRange.Cells(idxUltima, 2).Value = idPedido           ' ID_PEDIDO
        tbChequeo.DataBodyRange.Cells(idxUltima, 3).Value = idAtencion         ' ID_ATENCION

        With tbChequeo.DataBodyRange.Cells(idxUltima, 4)                        ' FECHA
            .Value = fechaNumero
            .NumberFormat = "dd/mm/yyyy"
            .Formula = .Value       ' fuerza reconocimiento como fecha
        End With

        tbChequeo.DataBodyRange.Cells(idxUltima, 5).Value = vendedor           ' Vendedor
        tbChequeo.DataBodyRange.Cells(idxUltima, 6).Value = nombreCliente      ' Nombre del Cliente
        tbChequeo.DataBodyRange.Cells(idxUltima, 7).Value = razonSocial        ' Razon Social
        tbChequeo.DataBodyRange.Cells(idxUltima, 8).Value = subCliente         ' Sub Cliente
        tbChequeo.DataBodyRange.Cells(idxUltima, 9).Value = IIf(codigoEncontrado, codigoExamen, "Sin codigo") ' Codigo examen
        tbChequeo.DataBodyRange.Cells(idxUltima, 10).Value = tipoServicio      ' Tipo de Servicio
        tbChequeo.DataBodyRange.Cells(idxUltima, 11).Value = examen            ' Examen
        tbChequeo.DataBodyRange.Cells(idxUltima, 12).Value = puesto            ' Puesto
        tbChequeo.DataBodyRange.Cells(idxUltima, 13).Value = provincia         ' Provincia
        tbChequeo.DataBodyRange.Cells(idxUltima, 14).Value = localidad         ' Localidad
        tbChequeo.DataBodyRange.Cells(idxUltima, 15).Value = sede              ' Sede
        tbChequeo.DataBodyRange.Cells(idxUltima, 16).Value = costo             ' Costo
        tbChequeo.DataBodyRange.Cells(idxUltima, 17).Value = precioVenta       ' Venta
        tbChequeo.DataBodyRange.Cells(idxUltima, 21).Value = horaTexto         ' Hora del sistema

        idAtencion = idAtencion + 1
    Next i

    tbChequeo.Parent.Protect UserInterfaceOnly:=True
End Sub


