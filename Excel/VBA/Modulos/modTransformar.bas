Option Explicit

' ============================================================
' modTransformar
' Punto de entrada del boton "Generar BDNE".
' Requisito: ambas quincenas deben estar cargadas.
' Genera un archivo Excel externo y una hoja BDNE_OUTPUT interna.
' Ambas salidas usan el mismo formato.
'
' MAPEO (formato actualizado):
'   LAFT -> BDNE:
'     Apellidos = Apellido Paterno + " " + Apellido Materno
'     Nombre    = Nombres
'     Source    = Tipo de Lista
'     Remark    = Entidad si Source = "PEP"; sino "NOAPLICA"
'     Pais      = "PERU" (con acento)
'     Program   = formula =CONCATENATE("BDNEG"," ","INSACO")
'     Sigla1    = "INSACO" (fijo)
'     id        = Documento si valido; sino alfanumerico de 8
'   PEP RELACIONADOS -> BDNE:
'     Apellidos = Ap. Paterno relacion + " " + Ap. Materno relacion
'     Nombre    = NOMBRES
'     Source    = "PARIENTE PEP " + NOMBRE_PEP
'     Remark    = "NOAPLICA"
'     Sigla1    = "INSACO" (fijo)
'     id        = DNIREL si valido; sino alfanumerico de 8
' ============================================================

Private Const BD_COLS As Integer = 11

Public Sub GenerarBDNE()
    Dim paso As String
    Dim pool As Object
    Dim filas() As Variant
    Dim nFilas As Long
    Dim wbOut As Workbook
    Dim wsOut As Worksheet
    Dim rutaSalida As String
    Dim rutaSeleccionada As Variant
    Dim nombreSugerido As String
    Dim rutaInicial As String
    Dim maxFilas As Long
    Dim hayQ1 As Boolean
    Dim hayQ2 As Boolean
    Dim msg As String
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    paso = "Validar quincenas cargadas"

    hayQ1 = HojaExiste(SH_Q1_LAFT)
    hayQ2 = HojaExiste(SH_Q2_LAFT)

    If Not hayQ1 Or Not hayQ2 Then
        msg = "Se requieren ambas quincenas para generar el BDNE." & vbCrLf & vbCrLf

        If Not hayQ1 Then msg = msg & Chr(10007) & "  Quincena 1: no cargada" & vbCrLf
        If Not hayQ2 Then msg = msg & Chr(10007) & "  Quincena 2: no cargada" & vbCrLf

        MsgBox msg, vbExclamation, "Datos incompletos"
        Exit Sub
    End If

    paso = "Seleccionar ruta de guardado"

    nombreSugerido = ConstruirNombreArchivo()
    rutaInicial = NombreDirectorioLibro() & nombreSugerido

    rutaSeleccionada = Application.GetSaveAsFilename( _
        InitialFileName:=rutaInicial, _
        FileFilter:="Libro de Excel (*.xlsx), *.xlsx", _
        Title:="Guardar BDNE como...")

    If VarType(rutaSeleccionada) = vbBoolean Then Exit Sub

    rutaSalida = CStr(rutaSeleccionada)

    If LCase(Right(rutaSalida, 5)) <> ".xlsx" Then
        rutaSalida = rutaSalida & ".xlsx"
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    paso = "Inicializar pool de IDs"
    Randomize
    Set pool = InicializarPool()

    paso = "Dimensionar buffer"
    maxFilas = ContarFilasDisponibles(True, True) + 10
    ReDim filas(1 To maxFilas, 1 To BD_COLS)
    nFilas = 0

    paso = "Procesar LAFT Q1"
    ProcesarLAFT LO_Q1_LAFT, filas, nFilas, pool

    paso = "Procesar REL Q1"
    If HojaExiste(SH_Q1_REL) Then ProcesarREL LO_Q1_REL, filas, nFilas, pool

    paso = "Procesar LAFT Q2"
    ProcesarLAFT LO_Q2_LAFT, filas, nFilas, pool

    paso = "Procesar REL Q2"
    If HojaExiste(SH_Q2_REL) Then ProcesarREL LO_Q2_REL, filas, nFilas, pool

    If nFilas = 0 Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "Los archivos cargados no contienen registros.", vbExclamation, "Sin datos"
        Exit Sub
    End If

    paso = "Crear workbook de salida"

    Set wbOut = Workbooks.Add
    Set wsOut = wbOut.Sheets(1)
    wsOut.Name = "Hoja1"

    Application.DisplayAlerts = False

    For Each ws In wbOut.Worksheets
        If ws.Name <> "Hoja1" Then ws.Delete
    Next ws

    Application.DisplayAlerts = True

    paso = "Escribir cabecera"
    EscribirCabecera wsOut

    paso = "Escribir filas"
    ' Forzar formato Texto en TODAS las celdas de datos ANTES de escribir,
    ' para preservar ceros a la izquierda en DNI/id y evitar que Excel
    ' interprete documentos como numeros. (La columna H se reescribe como
    ' formula despues, en AplicarFormatoBDNE.)
    wsOut.Range("A2").Resize(nFilas, BD_COLS).NumberFormat = "@"
    wsOut.Range("A2").Resize(nFilas, BD_COLS).Value = Left2D(filas, nFilas, BD_COLS)

    paso = "Aplicar formato al archivo exportado"
    AplicarFormatoBDNE wsOut, nFilas

    paso = "Crear BDNE_OUTPUT"
    CrearTablaVerificacion filas, nFilas

    paso = "Guardar archivo"

    Application.DisplayAlerts = False
    wbOut.SaveAs Filename:=rutaSalida, FileFormat:=51
    Application.DisplayAlerts = True
    wbOut.Close False

    paso = "Actualizar estado output"

    On Error Resume Next
    ThisWorkbook.Names(ST_OUT_ESTADO).RefersToRange.Value = _
        Chr(10003) & " " & Format(Now, "DD/MM/YYYY HH:MM") & _
        " - " & nFilas & " registros" & vbCrLf & _
        "(" & NombreDesdeRutaT(rutaSalida) & ")"
    On Error GoTo ErrHandler

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "BDNE generada correctamente." & vbCrLf & vbCrLf & _
           "Registros:  " & nFilas & vbCrLf & _
           "Archivo:    " & NombreDesdeRutaT(rutaSalida), _
           vbInformation, "Proceso completado"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close False
    On Error GoTo 0

    MsgBox "Error en paso: " & paso & vbCrLf & vbCrLf & _
           "Descripcion: " & Err.Description & vbCrLf & _
           "Numero: " & Err.Number, _
           vbCritical, "Error al generar BDNE"
End Sub

Private Function ConstruirNombreArchivo() As String
    Dim iniQ1 As String
    Dim finQ2 As String
    Dim periodo As String

    iniQ1 = ExtraerFechaInicio(g_NombreArchivoQ1)
    finQ2 = ExtraerFechaFin(g_NombreArchivoQ2)
    periodo = Format(Now, "MM_YY")

    If iniQ1 <> "" And finQ2 <> "" Then
        ConstruirNombreArchivo = "FORMATOBDNED_" & iniQ1 & "_AL_" & finQ2 & "_" & periodo & ".xlsx"
    Else
        ConstruirNombreArchivo = "FORMATOBDNED_" & Format(Now, "MM_YY") & ".xlsx"
    End If
End Function

Private Function ExtraerFechaInicio(ByVal nombre As String) As String
    Dim parts() As String
    Dim i As Integer

    nombre = Replace(nombre, ".xlsx", "")
    nombre = Replace(nombre, ".xls", "")

    parts = Split(nombre, "_")

    For i = 0 To UBound(parts) - 1
        If IsNumeric(parts(i)) And Len(parts(i)) <= 2 Then
            If i + 1 <= UBound(parts) And IsNumeric(parts(i + 1)) And Len(parts(i + 1)) <= 2 Then
                ExtraerFechaInicio = Format(CLng(parts(i)), "00") & "_" & Format(CLng(parts(i + 1)), "00")
                Exit Function
            End If
        End If
    Next i

    ExtraerFechaInicio = ""
End Function

Private Function ExtraerFechaFin(ByVal nombre As String) As String
    Dim parts() As String
    Dim i As Integer
    Dim ultimo As String

    nombre = Replace(nombre, ".xlsx", "")
    nombre = Replace(nombre, ".xls", "")

    parts = Split(nombre, "_")
    ultimo = ""

    For i = 0 To UBound(parts) - 1
        If IsNumeric(parts(i)) And Len(parts(i)) <= 2 Then
            If i + 1 <= UBound(parts) And IsNumeric(parts(i + 1)) And Len(parts(i + 1)) <= 2 Then
                ultimo = Format(CLng(parts(i)), "00") & "_" & Format(CLng(parts(i + 1)), "00")
            End If
        End If
    Next i

    ExtraerFechaFin = ultimo
End Function

Private Function NombreDirectorioLibro() As String
    If ThisWorkbook.Path <> "" Then
        NombreDirectorioLibro = ThisWorkbook.Path & Application.PathSeparator
    Else
        NombreDirectorioLibro = ""
    End If
End Function

Private Function NombreDesdeRutaT(ByVal ruta As String) As String
    Dim pos As Long

    pos = InStrRev(ruta, "\")
    If pos = 0 Then pos = InStrRev(ruta, "/")

    If pos = 0 Then
        NombreDesdeRutaT = ruta
    Else
        NombreDesdeRutaT = Mid(ruta, pos + 1)
    End If
End Function

Private Sub ProcesarLAFT(ByVal nombreLO As String, ByRef filas() As Variant, _
                          ByRef nFilas As Long, ByVal pool As Object)
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim arr As Variant
    Dim nRows As Long
    Dim i As Long

    Dim cTipoLista As Integer
    Dim cDocumento As Integer
    Dim cNombres As Integer
    Dim cApPat As Integer
    Dim cApMat As Integer
    Dim cCargo As Integer
    Dim cEntidad As Integer
    Dim cNoticia As Integer

    Dim tipoLista As String
    Dim documento As String
    Dim nombres As String
    Dim apPat As String
    Dim apMat As String
    Dim entidad As String
    Dim sourceVal As String
    Dim remarkVal As String
    Dim idVal As String

    Set ws = BuscarHojaPorLO(nombreLO)
    If ws Is Nothing Then Exit Sub

    Set lo = ws.ListObjects(nombreLO)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    arr = lo.DataBodyRange.Value
    nRows = UBound(arr, 1)

    cTipoLista = IndiceColumna(lo, SRC_TIPO_LISTA)
    cDocumento = IndiceColumna(lo, SRC_DOCUMENTO)
    cNombres = IndiceColumna(lo, SRC_NOMBRES)
    cApPat = IndiceColumna(lo, SRC_AP_PAT)
    cApMat = IndiceColumna(lo, SRC_AP_MAT)
    cCargo = IndiceColumna(lo, SRC_CARGO)
    cEntidad = IndiceColumna(lo, SRC_ENTIDAD)
    cNoticia = IndiceColumna(lo, SRC_NOTICIA)

    For i = 1 To nRows
        tipoLista = ValorMatriz(arr, i, cTipoLista)
        documento = ValorCeldaTexto(lo, i, cDocumento)
        nombres = ValorMatriz(arr, i, cNombres)
        apPat = ValorMatriz(arr, i, cApPat)
        apMat = ValorMatriz(arr, i, cApMat)
        entidad = ValorMatriz(arr, i, cEntidad)

        ' Source = Tipo de Lista (siempre)
        sourceVal = tipoLista

        ' Remark = Entidad si es PEP con entidad valida; sino "NOAPLICA"
        If UCase(tipoLista) = "PEP" And UCase(entidad) <> SRC_NOAPLICA And entidad <> "" Then
            remarkVal = entidad
        Else
            remarkVal = BD_NOAPLICA
        End If

        ' id = Documento si valido; sino alfanumerico
        If EsDocumentoValido(documento) Then
            RegistrarEnPool pool, documento
            idVal = documento
        Else
            idVal = GenerarIDUnico(pool)
        End If

        nFilas = nFilas + 1

        ' Sigla1 = INSACO (fijo). Program se reescribe como formula en AplicarFormatoBDNE.
        EscribirFila filas, nFilas, _
            JuntarApellidos(apPat, apMat), nombres, sourceVal, remarkVal, _
            PaisValor(), "", "", BD_PROGRAM_PREFIJO & BD_SIGLA_VALOR, BD_SIGLA_VALOR, "", idVal
    Next i
End Sub

Private Sub ProcesarREL(ByVal nombreLO As String, ByRef filas() As Variant, _
                         ByRef nFilas As Long, ByVal pool As Object)
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim arr As Variant
    Dim nRows As Long
    Dim i As Long

    Dim cNombrePEP As Integer
    Dim cDNIRel As Integer
    Dim cApPat As Integer
    Dim cApMat As Integer
    Dim cNombres As Integer

    Dim nombrePEP As String
    Dim dniRel As String
    Dim apPat As String
    Dim apMat As String
    Dim nombresRel As String
    Dim sourceVal As String
    Dim idVal As String

    Set ws = BuscarHojaPorLO(nombreLO)
    If ws Is Nothing Then Exit Sub

    Set lo = ws.ListObjects(nombreLO)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    arr = lo.DataBodyRange.Value
    nRows = UBound(arr, 1)

    cNombrePEP = IndiceColumna(lo, REL_NOMBRE_PEP)
    cDNIRel = IndiceColumna(lo, REL_DNIREL)
    cApPat = IndiceColumna(lo, REL_AP_PAT)
    cApMat = IndiceColumna(lo, REL_AP_MAT)
    cNombres = IndiceColumna(lo, REL_NOMBRES)

    For i = 1 To nRows
        nombrePEP = ValorMatriz(arr, i, cNombrePEP)
        dniRel = ValorCeldaTexto(lo, i, cDNIRel)
        apPat = ValorMatriz(arr, i, cApPat)
        apMat = ValorMatriz(arr, i, cApMat)
        nombresRel = ValorMatriz(arr, i, cNombres)

        ' Source = "PARIENTE PEP " + nombre del titular PEP
        sourceVal = BD_PARIENTE_PREFIJO & nombrePEP

        ' id = DNIREL si valido; sino alfanumerico
        If EsDocumentoValido(dniRel) Then
            RegistrarEnPool pool, dniRel
            idVal = dniRel
        Else
            idVal = GenerarIDUnico(pool)
        End If

        nFilas = nFilas + 1

        ' Remark = "NOAPLICA" fijo. Sigla1 = INSACO.
        EscribirFila filas, nFilas, _
            JuntarApellidos(apPat, apMat), nombresRel, sourceVal, BD_NOAPLICA, _
            PaisValor(), "", "", BD_PROGRAM_PREFIJO & BD_SIGLA_VALOR, BD_SIGLA_VALOR, "", idVal
    Next i
End Sub

Private Sub EscribirFila(ByRef filas() As Variant, ByVal n As Long, _
    ByVal apellidos As String, ByVal nombre As String, _
    ByVal source As String, ByVal remark As String, _
    ByVal pais As String, ByVal dir As String, ByVal ciudad As String, _
    ByVal program As String, ByVal sigla1 As String, _
    ByVal sigla2 As String, ByVal id As String)

    filas(n, 1) = apellidos
    filas(n, 2) = nombre
    filas(n, 3) = source
    filas(n, 4) = remark
    filas(n, 5) = pais
    filas(n, 6) = dir
    filas(n, 7) = ciudad
    filas(n, 8) = program
    filas(n, 9) = sigla1
    filas(n, 10) = sigla2
    filas(n, 11) = id
End Sub

Private Sub EscribirCabecera(ByVal ws As Worksheet)
    ws.Range("A1").Value = BD_APELLIDOS
    ws.Range("B1").Value = BD_NOMBRE
    ws.Range("C1").Value = BD_SOURCE
    ws.Range("D1").Value = BD_REMARK
    ws.Range("E1").Value = BD_PAIS
    ws.Range("F1").Value = BD_DIRECCION
    ws.Range("G1").Value = BD_CIUDAD
    ws.Range("H1").Value = BD_PROGRAM
    ws.Range("I1").Value = BD_SIGLA1
    ws.Range("J1").Value = BD_SIGLA2
    ws.Range("K1").Value = BD_ID
End Sub

Private Sub AplicarFormatoBDNE(ByVal ws As Worksheet, ByVal nFilas As Long)
    Dim rngAll As Range
    Dim rngHdr As Range
    Dim lo As ListObject
    Dim r As Long

    Set rngAll = ws.Range("A1").Resize(nFilas + 1, BD_COLS)
    Set rngHdr = ws.Range("A1").Resize(1, BD_COLS)

    Set lo = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        source:=rngAll, _
        XlListObjectHasHeaders:=xlYes)

    lo.Name = LO_OUTPUT
    lo.TableStyle = ""

    ' Todas las columnas en Sura Sans 9 (datos)
    With rngAll.Font
        .Name = "Sura Sans"
        .Size = 9
        .Bold = False
    End With

    rngAll.VerticalAlignment = xlCenter

    ' Cabecera: Sura Sans 10, blanca sobre negro
    With rngHdr
        .Font.Name = "Sura Sans"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .VerticalAlignment = xlCenter
    End With

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    ' Program como formula con la sigla fija como texto literal:
    ' =CONCATENATE("BDNEG"," ","INSACO")
    ' La columna H debe ser General (no Texto) para que la formula se
    ' evalue en vez de mostrarse como texto.
    ws.Range(ws.Cells(2, 8), ws.Cells(nFilas + 1, 8)).NumberFormat = "General"
    For r = 2 To nFilas + 1
        ws.Cells(r, 8).Formula = "=CONCATENATE(""BDNEG"","" "",""" & BD_SIGLA_VALOR & """)"
    Next r

    On Error Resume Next
    ws.Range(ws.Cells(1, 1), ws.Cells(1, BD_COLS)).EntireColumn.AutoFit
    On Error GoTo 0
End Sub

Private Sub CrearTablaVerificacion(ByRef filas() As Variant, ByVal nFilas As Long)
    Dim ws As Worksheet

    EliminarHojasSiExisten Array(SH_OUTPUT)

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = SH_OUTPUT

    EscribirCabecera ws
    ' Forzar formato Texto antes de escribir (preserva ceros a la izquierda)
    ws.Range("A2").Resize(nFilas, BD_COLS).NumberFormat = "@"
    ws.Range("A2").Resize(nFilas, BD_COLS).Value = Left2D(filas, nFilas, BD_COLS)

    AplicarFormatoBDNE ws, nFilas
End Sub

Private Function JuntarApellidos(ByVal pat As String, ByVal mat As String) As String
    If mat = "" Then
        JuntarApellidos = pat
    Else
        JuntarApellidos = Trim(pat & " " & mat)
    End If
End Function

Private Function Limpio(ByVal val As Variant) As String
    Dim s As String

    If IsNull(val) Or IsEmpty(val) Then
        Limpio = ""
        Exit Function
    End If

    s = Trim(CStr(val))

    If UCase(s) = "NAN" Or s = "0" Then
        Limpio = ""
    Else
        Limpio = s
    End If
End Function

Private Function ValorMatriz(ByRef arr As Variant, ByVal fila As Long, ByVal col As Integer) As String
    If col <= 0 Then
        ValorMatriz = ""
    Else
        ValorMatriz = Limpio(arr(fila, col))
    End If
End Function

Private Function ValorCeldaTexto(ByVal lo As ListObject, ByVal fila As Long, ByVal col As Integer) As String
    If col <= 0 Then
        ValorCeldaTexto = ""
    Else
        ValorCeldaTexto = Limpio(lo.DataBodyRange.Cells(fila, col).Text)
    End If
End Function

Private Function BuscarHojaPorLO(ByVal nombreLO As String) As Worksheet
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nombreLO Then
                Set BuscarHojaPorLO = ws
                Exit Function
            End If
        Next lo
    Next ws

    Set BuscarHojaPorLO = Nothing
End Function

Private Function IndiceColumna(ByVal lo As ListObject, ByVal nombreCol As String) As Integer
    Dim lc As ListColumn

    On Error Resume Next
    Set lc = lo.ListColumns(nombreCol)
    On Error GoTo 0

    If lc Is Nothing Then
        IndiceColumna = 0
    Else
        IndiceColumna = lc.Index
    End If
End Function

Private Function ContarFilasDisponibles(ByVal hayQ1 As Boolean, ByVal hayQ2 As Boolean) As Long
    Dim total As Long

    If hayQ1 Then
        total = total + ContarFilasLO(LO_Q1_LAFT) + ContarFilasLO(LO_Q1_REL)
    End If

    If hayQ2 Then
        total = total + ContarFilasLO(LO_Q2_LAFT) + ContarFilasLO(LO_Q2_REL)
    End If

    ContarFilasDisponibles = total
End Function

Private Function ContarFilasLO(ByVal nombreLO As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = BuscarHojaPorLO(nombreLO)

    If ws Is Nothing Then
        ContarFilasLO = 0
        Exit Function
    End If

    Set lo = ws.ListObjects(nombreLO)

    If lo.DataBodyRange Is Nothing Then
        ContarFilasLO = 0
    Else
        ContarFilasLO = lo.DataBodyRange.Rows.count
    End If
End Function

Private Function Left2D(ByRef arr() As Variant, ByVal nFilas As Long, ByVal nCols As Integer) As Variant
    Dim result() As Variant
    Dim i As Long
    Dim j As Integer

    ReDim result(1 To nFilas, 1 To nCols)

    For i = 1 To nFilas
        For j = 1 To nCols
            result(i, j) = arr(i, j)
        Next j
    Next i

    Left2D = result
End Function