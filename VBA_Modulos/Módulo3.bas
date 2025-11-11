Attribute VB_Name = "Módulo3"
Option Explicit

' Agrega el item seleccionado a la celda "items" si el tercer desplegable es Drogas o Laboratorio
Public Sub AgregarItemACelda()
    ' defino hoja y rangos
    Dim ws As Worksheet
    Dim celdaDestino As Range
    Set ws = ThisWorkbook.Sheets("Formulario")
    Set celdaDestino = ws.Range("items")

    ' tomo valores de los desplegables
    Dim itemSeleccionado As String
    Dim tercerDesplegable As String
    itemSeleccionado = Trim$(CStr(ws.Range("IIII_desplegable").Value))   ' item a agregar
    tercerDesplegable = Trim$(CStr(ws.Range("III_desplegable").Value))   ' Drogas / Laboratorio / otros

    ' valido tipo de tercer desplegable
    If tercerDesplegable <> "Drogas" And tercerDesplegable <> "Laboratorio" Then
        MsgBox "El boton solo esta disponible para Drogas o Laboratorio.", vbExclamation
        Exit Sub
    End If

    ' valido item seleccionado
    If itemSeleccionado = "" Then
        MsgBox "Por favor, seleccione un item antes de agregarlo.", vbExclamation
        Exit Sub
    End If

    ' leo texto existente en la celda destino
    Dim textoExistente As String
    textoExistente = Trim$(CStr(celdaDestino.Value))

    ' separo items existentes en un array
    Dim items() As String
    items = TextoAColeccion(textoExistente)    ' puede devolver array vacio

    ' normalizo item para comparar
    Dim itemNormalizado As String
    itemNormalizado = NormalizarTexto(itemSeleccionado)

    ' verifico si el item ya existe
    If ExisteItem(items, itemNormalizado) Then
        MsgBox "El item ya esta en la lista.", vbExclamation
        ws.Range("IIII_desplegable").Value = ""   ' limpio el desplegable
        Exit Sub
    End If

    ' agrego el item nuevo al array
    items = AgregarItemAlArray(items, itemSeleccionado)

    ' ordeno los items alfabeticamente
    If LBound(items) <= UBound(items) Then
        QuickSortTexto items, LBound(items), UBound(items)
    End If

    ' reconstruyo el texto concatenado
    celdaDestino.Value = ColeccionATexto(items)

    ' limpio el desplegable para permitir nueva seleccion
    ws.Range("IIII_desplegable").Value = ""
End Sub

' auxiliares de texto

' normaliza texto para comparar
Private Function NormalizarTexto(ByVal s As String) As String
    ' paso a minusculas y quito espacios duplicados
    Dim t As String
    t = LCase$(Trim$(CStr(s)))
    t = Application.WorksheetFunction.Trim(t)
    NormalizarTexto = t
End Function

' convierte texto "a, b, c" en array de items (sin vacios)
Private Function TextoAColeccion(ByVal texto As String) As String()
    Dim arr() As String
    Dim tmp As Variant
    Dim i As Long, k As Long

    If Len(Trim$(texto)) = 0 Then
        ReDim arr(-1 To -1)
        TextoAColeccion = arr
        Exit Function
    End If

    tmp = Split(texto, ",")
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

    TextoAColeccion = arr
End Function

' convierte array de items a texto "a, b, c"
Private Function ColeccionATexto(ByRef items() As String) As String
    If (Not Not items) = 0 Then
        ColeccionATexto = ""
    ElseIf UBound(items) < LBound(items) Then
        ColeccionATexto = ""
    Else
        ColeccionATexto = Join(items, ", ")
    End If
End Function

' verifica existencia de un item normalizado dentro del array (compara normalizado)
Private Function ExisteItem(ByRef items() As String, ByVal itemNorm As String) As Boolean
    Dim i As Long
    If (Not Not items) = 0 Then Exit Function
    For i = LBound(items) To UBound(items)
        If NormalizarTexto(items(i)) = itemNorm Then
            ExisteItem = True
            Exit Function
        End If
    Next i
End Function

' agrega al final del array (crea array si esta vacio)
Private Function AgregarItemAlArray(ByRef items() As String, ByVal nuevo As String) As String()
    Dim res() As String
    Dim n As Long

    If (Not Not items) = 0 Or UBound(items) < LBound(items) Then
        ReDim res(0 To 0)
        res(0) = nuevo
    Else
        n = UBound(items) + 1
        ReDim res(LBound(items) To n)
        Dim i As Long
        For i = LBound(items) To UBound(items)
            res(i) = items(i)
        Next i
        res(n) = nuevo
    End If

    AgregarItemAlArray = res
End Function

' ordenamiento

' quicksort case-insensitive para strings
Private Sub QuickSortTexto(ByRef A() As String, ByVal lo As Long, ByVal hi As Long)
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

