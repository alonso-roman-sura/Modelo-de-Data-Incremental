Option Explicit

' ============================================================
' modGenerarID
' Genera identificadores alfanumericos de 8 caracteres para
' registros que llegan sin documento valido.
' ============================================================

Private Const CHARS_LETRAS  As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const CHARS_TODOS   As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
Private Const ID_LEN        As Integer = 8
Private Const MAX_INTENTOS  As Integer = 10000

Public Function GenerarIDUnico(ByVal pool As Object) As String
    Dim candidato As String
    Dim intentos As Integer

    intentos = 0

    Do
        candidato = GenerarCandidato()
        intentos = intentos + 1

        If intentos > MAX_INTENTOS Then
            Err.Raise vbObjectError + 1001, "modGenerarID", _
                "No se pudo generar un ID unico tras " & MAX_INTENTOS & " intentos."
        End If
    Loop While pool.Exists(candidato)

    pool.Add candidato, True
    GenerarIDUnico = candidato
End Function

Private Function GenerarCandidato() As String
    Dim i As Integer
    Dim posLetra As Integer
    Dim resultado() As String
    Dim tieneLetra As Boolean

    ReDim resultado(1 To ID_LEN)

    For i = 1 To ID_LEN
        resultado(i) = Mid(CHARS_TODOS, Int(Rnd() * Len(CHARS_TODOS)) + 1, 1)
    Next i

    tieneLetra = False
    For i = 1 To ID_LEN
        If resultado(i) >= "A" Then
            tieneLetra = True
            Exit For
        End If
    Next i

    If Not tieneLetra Then
        posLetra = Int(Rnd() * ID_LEN) + 1
        resultado(posLetra) = Mid(CHARS_LETRAS, Int(Rnd() * Len(CHARS_LETRAS)) + 1, 1)
    End If

    GenerarCandidato = Join(resultado, "")
End Function

Public Function InicializarPool() As Object
    Dim pool As Object
    Dim lo As ListObject
    Dim col As ListColumn
    Dim arr As Variant
    Dim i As Long
    Dim val As String

    Set pool = CreateObject("Scripting.Dictionary")
    pool.CompareMode = vbTextCompare

    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(SH_OUTPUT).ListObjects(LO_OUTPUT)
    On Error GoTo 0

    If lo Is Nothing Then
        Set InicializarPool = pool
        Exit Function
    End If

    On Error Resume Next
    Set col = lo.ListColumns(BD_ID)
    On Error GoTo 0

    If col Is Nothing Then
        Set InicializarPool = pool
        Exit Function
    End If

    If col.DataBodyRange Is Nothing Then
        Set InicializarPool = pool
        Exit Function
    End If

    arr = col.DataBodyRange.Value

    If IsArray(arr) Then
        For i = 1 To UBound(arr, 1)
            val = Trim(CStr(arr(i, 1)))
            If val <> "" And Not pool.Exists(val) Then
                pool.Add val, True
            End If
        Next i
    End If

    Set InicializarPool = pool
End Function

Public Function EsDocumentoValido(ByVal doc As String) As Boolean
    Dim d As String
    d = UCase(Trim(doc))
    EsDocumentoValido = (d <> "" And _
                         d <> SRC_SIN_ID And _
                         d <> SRC_NOAPLICA And _
                         d <> "NAN")
End Function

Public Sub RegistrarEnPool(ByVal pool As Object, ByVal doc As String)
    If Not pool.Exists(doc) Then
        pool.Add doc, True
    End If
End Sub