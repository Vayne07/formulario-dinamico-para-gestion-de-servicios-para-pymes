Attribute VB_Name = "Módulo4"
Option Explicit

' Agrega el examen seleccionado al perfil armado (celda "perfiles_armados")
Public Sub AgregarExamenAperfil()
    ' defino hoja
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Formulario")

    ' valido que este en modo Perfil
    If ws.Range("II_desplegable").Value <> "Perfil" Then Exit Sub

    ' tomo seleccion actual
    Dim tercer As String
    Dim itemUno As String
    Dim itemsLista As String
    tercer = Trim$(CStr(ws.Range("III_desplegable").Value))   ' Drogas / Laboratorio / otro
    itemUno = Trim$(CStr(ws.Range("item").Value))             ' item individual
    itemsLista = Trim$(CStr(ws.Range("items").Value))         ' lista separada por coma

    ' construyo el texto del examen segun logica
    Dim examen As String
    If tercer = "" And itemUno = "" And itemsLista = "" Then
        MsgBox "No hay ningun examen seleccionado para agregar al perfil.", vbExclamation, "Aviso"
        Exit Sub
    End If

    If tercer = "Drogas" Or tercer = "Laboratorio" Then
        ' caso drogas/laboratorio: concateno item + items
        If itemUno <> "" And itemsLista <> "" Then
            examen = itemUno & " " & itemsLista
        ElseIf itemUno <> "" Then
            examen = itemUno
        ElseIf itemsLista <> "" Then
            examen = itemsLista
        Else
            MsgBox "Seleccione al menos un item en Drogas/Laboratorio.", vbExclamation, "Aviso"
            Exit Sub
        End If
    Else
        ' caso general: uso el tercer desplegable
        If tercer <> "" Then
            examen = tercer
        Else
            MsgBox "No hay examen seleccionado en el tercer desplegable.", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If

    ' leo perfil actual
    Dim perfil As String
    perfil = CStr(ws.Range("perfiles_armados").Value)

    ' paso perfil a array usando separador " + "
    Dim examenes() As String
    examenes = PerfilATabla(perfil)

    ' evito duplicados comparando normalizado
    If ExisteExamen(examenes, examen) Then
        MsgBox "El examen '" & examen & "' ya existe en el perfil y no se puede agregar nuevamente.", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' agrego examen al array
    examenes = AgregarAlArray(examenes, examen)

    ' ordeno alfabeticamente (case-insensitive)
    If ArrayTieneDatos(examenes) Then
        QuickSortTexto examenes, LBound(examenes), UBound(examenes)
    End If

    ' muevo "Basico de Ley" al inicio si existe
    examenes = BasicoDeLeyPrimero(examenes)

    ' escribo resultado
    ws.Unprotect
    ws.Range("perfiles_armados").Value = TablaAPerfil(examenes)
    ws.Protect
End Sub

' auxiliares de perfil
' convierte el texto del perfil a array usando " + " como separador
Private Function PerfilATabla(ByVal perfil As String) As String()
    Dim arr() As String, tmp As Variant
    Dim i As Long, k As Long

    If Len(Trim$(perfil)) = 0 Then
        ReDim arr(-1 To -1)
        PerfilATabla = arr
        Exit Function
    End If

    tmp = Split(perfil, " + ")
    ReDim arr(LBound(tmp) To UBound(tmp))
    k = LBound(arr)

    For i = LBound(tmp) To UBound(tmp)
        Dim pieza As String
        pieza = Trim$(CStr(tmp(i)))
        If pieza <> "" Then
            arr(k) = pieza
            k = k + 1
        End If
    Next i

    If k = LBound(arr) Then
        ReDim arr(-1 To -1)
    Else
        ReDim Preserve arr(LBound(arr) To k - 1)
    End If

    PerfilATabla = arr
End Function

' convierte el array al texto de perfil usando " + " como separador
Private Function TablaAPerfil(ByRef examenes() As String) As String
    If Not ArrayTieneDatos(examenes) Then
        TablaAPerfil = ""
    Else
        TablaAPerfil = Join(examenes, " + ")
    End If
End Function

' verifica si el array tiene datos validos
Private Function ArrayTieneDatos(ByRef A() As String) As Boolean
    On Error Resume Next
    ArrayTieneDatos = ((Not Not A) <> 0 And UBound(A) >= LBound(A))
End Function

' true si ya existe el examen (comparacion normalizada)
Private Function ExisteExamen(ByRef examenes() As String, ByVal examen As String) As Boolean
    Dim i As Long, needle As String
    needle = NormalizarTexto(examen)
    If Not ArrayTieneDatos(examenes) Then Exit Function
    For i = LBound(examenes) To UBound(examenes)
        If NormalizarTexto(examenes(i)) = needle Then
            ExisteExamen = True
            Exit Function
        End If
    Next i
End Function

' agrega un elemento al final (crea array si estaba vacio)
Private Function AgregarAlArray(ByRef examenes() As String, ByVal examen As String) As String()
    Dim res() As String, i As Long, n As Long
    If Not ArrayTieneDatos(examenes) Then
        ReDim res(0 To 0)
        res(0) = examen
    Else
        n = UBound(examenes) + 1
        ReDim res(LBound(examenes) To n)
        For i = LBound(examenes) To UBound(examenes)
            res(i) = examenes(i)
        Next i
        res(n) = examen
    End If
    AgregarAlArray = res
End Function

' mueve "Basico de Ley" al inicio si esta presente
Private Function BasicoDeLeyPrimero(ByRef examenes() As String) As String()
    If Not ArrayTieneDatos(examenes) Then
        BasicoDeLeyPrimero = examenes
        Exit Function
    End If

    Dim i As Long, idx As Long: idx = -1
    For i = LBound(examenes) To UBound(examenes)
        If NormalizarTexto(examenes(i)) = "basico de ley" Then
            idx = i
            Exit For
        End If
    Next i

    If idx = -1 Then
        BasicoDeLeyPrimero = examenes
        Exit Function
    End If

    If UBound(examenes) = LBound(examenes) Then
        BasicoDeLeyPrimero = examenes
        Exit Function
    End If

    ' quito el elemento y lo inserto al inicio
    Dim res() As String, k As Long
    ReDim res(LBound(examenes) To UBound(examenes))
    res(LBound(res)) = examenes(idx)

    k = LBound(res) + 1
    For i = LBound(examenes) To UBound(examenes)
        If i <> idx Then
            res(k) = examenes(i)
            k = k + 1
        End If
    Next i

    BasicoDeLeyPrimero = res
End Function

' helpers de texto y orden

' normaliza para comparar (minusculas y trim de espacios repetidos)
Private Function NormalizarTexto(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(CStr(s)))
    t = Application.WorksheetFunction.Trim(t)
    NormalizarTexto = t
End Function

' quicksort case-insensitive para strings
Private Sub QuickSortTexto(ByRef A() As String, ByVal lo As Long, ByVal hi As Long)
    If Not ArrayTieneDatos(A) Then Exit Sub
    If lo >= hi Then Exit Sub

    Dim i As Long, j As Long
    Dim pivot As String
    Dim temp As String

    i = lo
    j = hi
    pivot = LCase$(A((lo + hi) \ 2))

    Do While i <= j
        Do While LCase$(A(i)) < pivot
            i = i + 1
        Loop
        Do While LCase$(A(j)) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = A(i)
            A(i) = A(j)
            A(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If lo < j Then QuickSortTexto A, lo, j
    If i < hi Then QuickSortTexto A, i, hi
End Sub

