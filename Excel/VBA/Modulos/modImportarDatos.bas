Option Explicit

' ============================================================
' modImportarDatos
' Importa el archivo incremental de una quincena al libro.
' Cada archivo tiene dos hojas: LAFT y PEP RELACIONADOS.
' Todos los datos se copian como Texto explicito para preservar
' ceros a la izquierda (DNI, id) y evitar conversiones numericas.
' ============================================================

Public Sub CargarQuincena1()
    CargarQuincena 1
End Sub

Public Sub CargarQuincena2()
    CargarQuincena 2
End Sub

Private Sub CargarQuincena(ByVal nQ As Integer)
    Dim paso As String
    Dim rutaArchivo As String
    Dim wbFuente As Workbook
    Dim wsFuente As Worksheet
    Dim wsRel As Worksheet
    Dim shLAFT As String
    Dim shREL As String
    Dim nLAFT As Long
    Dim nREL As Long

    On Error GoTo ErrHandler

    shLAFT = IIf(nQ = 1, SH_Q1_LAFT, SH_Q2_LAFT)
    shREL = IIf(nQ = 1, SH_Q1_REL, SH_Q2_REL)

    paso = "Seleccionar archivo"
    rutaArchivo = SeleccionarArchivo("Seleccionar incremental quincena " & nQ)
    If rutaArchivo = "" Then Exit Sub

    paso = "Abrir archivo"
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wbFuente = Workbooks.Open(rutaArchivo, ReadOnly:=True, UpdateLinks:=False)

    paso = "Buscar hoja LAFT"
    Set wsFuente = BuscarHoja(wbFuente, SRC_HOJA_LAFT)
    If wsFuente Is Nothing Then
        wbFuente.Close False
        MsgBox "No se encontro la hoja """ & SRC_HOJA_LAFT & """" & _
               " en el archivo." & vbCrLf & rutaArchivo, vbExclamation, "Error de formato"
        GoTo Salir
    End If

    paso = "Validar columnas LAFT"
    If Not ValidarColumnas(wsFuente, SRC_LAFT_HEADER_ROW, ColsRequeridasLAFT()) Then
        wbFuente.Close False
        GoTo Salir
    End If

    paso = "Buscar hoja PEP RELACIONADOS"
    Set wsRel = BuscarHoja(wbFuente, SRC_HOJA_REL)
    If wsRel Is Nothing Then
        wbFuente.Close False
        MsgBox "No se encontro la hoja """ & SRC_HOJA_REL & """" & _
               " en el archivo." & vbCrLf & rutaArchivo, vbExclamation, "Error de formato"
        GoTo Salir
    End If

    paso = "Validar columnas PEP RELACIONADOS"
    If Not ValidarColumnas(wsRel, 1, ColsRequeridasREL()) Then
        wbFuente.Close False
        GoTo Salir
    End If

    paso = "Eliminar hojas previas Q" & nQ
    EliminarHojasSiExisten Array(shLAFT, shREL)

    paso = "Copiar datos LAFT"
    nLAFT = CopiarHojaFuente(wsFuente, shLAFT, LO_LAFT(nQ), SRC_LAFT_HEADER_ROW)

    paso = "Copiar datos PEP RELACIONADOS"
    nREL = CopiarHojaFuente(wsRel, shREL, LO_REL(nQ), 1)

    wbFuente.Close False
    Set wbFuente = Nothing

    paso = "Guardar nombre de archivo"
    If nQ = 1 Then
        g_NombreArchivoQ1 = NombreDesdeRuta(rutaArchivo)
    Else
        g_NombreArchivoQ2 = NombreDesdeRuta(rutaArchivo)
    End If

    paso = "Actualizar estado"
    ActualizarEstado nQ, nLAFT, nREL, NombreDesdeRuta(rutaArchivo)

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Quincena " & nQ & " cargada correctamente." & vbCrLf & vbCrLf & _
           "  LAFT:             " & nLAFT & " registros" & vbCrLf & _
           "  PEP RELACIONADOS: " & nREL & " registros", _
           vbInformation, "Carga exitosa"
    Exit Sub

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    Dim errPaso As String

    errNum = Err.Number
    errDesc = Err.Description
    errPaso = paso

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    On Error Resume Next
    If Not wbFuente Is Nothing Then wbFuente.Close False
    On Error GoTo 0

    MsgBox "Error al cargar quincena " & nQ & vbCrLf & vbCrLf & _
           "Paso:        " & errPaso & vbCrLf & _
           "Descripcion: " & errDesc & vbCrLf & _
           "Numero:      " & errNum, _
           vbCritical, "Error al cargar quincena " & nQ

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Private Function SeleccionarArchivo(ByVal titulo As String) As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = titulo
        .Filters.Clear
        .Filters.Add "Archivos Excel", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False

        If .Show = True Then
            SeleccionarArchivo = .SelectedItems(1)
        Else
            SeleccionarArchivo = ""
        End If
    End With
End Function

Private Function BuscarHoja(ByVal wb As Workbook, ByVal nombre As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If UCase(Trim(ws.Name)) = UCase(Trim(nombre)) Then
            Set BuscarHoja = ws
            Exit Function
        End If
    Next ws

    Set BuscarHoja = Nothing
End Function

Private Function ValidarColumnas(ByVal ws As Worksheet, _
                                  ByVal headerRow As Integer, _
                                  ByVal colsReq As Variant) As Boolean
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim encontrado As Boolean
    Dim cabeceras() As String
    Dim faltantes As String

    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    ReDim cabeceras(1 To lastCol)

    For i = 1 To lastCol
        cabeceras(i) = UCase(Trim(CStr(ws.Cells(headerRow, i).Value)))
    Next i

    faltantes = ""

    For j = LBound(colsReq) To UBound(colsReq)
        encontrado = False

        For i = 1 To lastCol
            If cabeceras(i) = UCase(Trim(CStr(colsReq(j)))) Then
                encontrado = True
                Exit For
            End If
        Next i

        If Not encontrado Then
            faltantes = faltantes & "  - " & colsReq(j) & vbCrLf
        End If
    Next j

    If faltantes <> "" Then
        MsgBox "La hoja """ & ws.Name & """ no contiene las columnas requeridas:" & _
               vbCrLf & faltantes, vbExclamation, "Formato invalido"
        ValidarColumnas = False
    Else
        ValidarColumnas = True
    End If
End Function

Private Function CopiarHojaFuente(ByVal wsFuente As Worksheet, _
                                   ByVal nombreDest As String, _
                                   ByVal nombreLO As String, _
                                   ByVal headerRow As Integer) As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim wsDest As Worksheet
    Dim rngOrig As Range
    Dim rngDest As Range
    Dim loDest As ListObject
    Dim loFuente As ListObject
    Dim loTemp As ListObject
    Dim r As Long
    Dim c As Long

    lastRow = wsFuente.Cells(wsFuente.Rows.count, 1).End(xlUp).Row
    lastCol = wsFuente.Cells(headerRow, wsFuente.Columns.count).End(xlToLeft).Column

    If lastRow <= headerRow Then
        CopiarHojaFuente = 0
        Exit Function
    End If

    Set rngOrig = wsFuente.Range(wsFuente.Cells(headerRow, 1), wsFuente.Cells(lastRow, lastCol))

    Set loFuente = Nothing
    For Each loTemp In wsFuente.ListObjects
        If Not Intersect(loTemp.Range, rngOrig) Is Nothing Then
            Set loFuente = loTemp
            Exit For
        End If
    Next loTemp

    Set wsDest = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsDest.Name = nombreDest
    wsDest.Visible = xlSheetVisible

    Set rngDest = wsDest.Range("A1").Resize(rngOrig.Rows.count, rngOrig.Columns.count)

    ' Forzar formato Texto en TODO el rango destino ANTES de copiar valores,
    ' para preservar ceros a la izquierda (DNI, documentos) y evitar que
    ' Excel convierta documentos en numeros.
    rngDest.NumberFormat = "@"

    ' Copiar valores como texto. Se usa .Text del origen via Value2 sobre
    ' celdas ya marcadas como Texto: el valor entra tal cual sin conversion.
    rngDest.Value2 = rngOrig.Value2

    Set loDest = wsDest.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        source:=rngDest, _
        XlListObjectHasHeaders:=xlYes)

    loDest.Name = nombreLO

    On Error Resume Next
    If Not loFuente Is Nothing Then
        loDest.TableStyle = loFuente.TableStyle
        loDest.ShowTableStyleRowStripes = loFuente.ShowTableStyleRowStripes
        loDest.ShowTableStyleColumnStripes = loFuente.ShowTableStyleColumnStripes
        loDest.ShowTableStyleFirstColumn = loFuente.ShowTableStyleFirstColumn
        loDest.ShowTableStyleLastColumn = loFuente.ShowTableStyleLastColumn
    Else
        loDest.TableStyle = ""
    End If
    On Error GoTo 0

    rngOrig.Copy
    rngDest.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Reaplicar formato Texto despues de pegar formatos (PasteFormats pudo
    ' haber traido un formato numerico del origen).
    rngDest.NumberFormat = "@"

    ' Conservar solo el estado oculto/visible de columnas (no el ancho:
    ' las columnas se autoajustan al contenido mas abajo).
    For c = 1 To lastCol
        wsDest.Columns(c).Hidden = wsFuente.Columns(c).Hidden
    Next c

    For r = 1 To rngOrig.Rows.count
        wsDest.Rows(r).RowHeight = wsFuente.Rows(headerRow + r - 1).RowHeight
        wsDest.Rows(r).Hidden = wsFuente.Rows(headerRow + r - 1).Hidden
    Next r

    wsDest.Cells.Font.Name = wsFuente.Cells.Font.Name

    ' Autofit de las columnas con datos para que se lean completas
    On Error Resume Next
    wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, lastCol)).EntireColumn.AutoFit
    On Error GoTo 0

    CopiarHojaFuente = lastRow - headerRow
End Function

Public Sub EliminarHojasSiExisten(ByVal nombres As Variant)
    Dim i As Integer
    Dim nombres2() As String
    Dim count As Integer

    count = 0

    For i = LBound(nombres) To UBound(nombres)
        If HojaExiste(CStr(nombres(i))) Then
            count = count + 1
            ReDim Preserve nombres2(count - 1)
            nombres2(count - 1) = CStr(nombres(i))
        End If
    Next i

    If count = 0 Then Exit Sub

    On Error GoTo ErrHandler
    Application.DisplayAlerts = False

    For i = 0 To count - 1
        ThisWorkbook.Worksheets(nombres2(i)).Delete
    Next i

    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Public Function HojaExiste(ByVal nombre As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0

    HojaExiste = Not (ws Is Nothing)
End Function

Private Sub ActualizarEstado(ByVal nQ As Integer, ByVal nLAFT As Long, _
                              ByVal nREL As Long, ByVal nombreArchivo As String)
    Dim nmEstado As String
    Dim nm As Name

    nmEstado = IIf(nQ = 1, ST_Q1_ESTADO, ST_Q2_ESTADO)

    On Error Resume Next
    Set nm = ThisWorkbook.Names(nmEstado)
    On Error GoTo 0

    If nm Is Nothing Then Exit Sub

    On Error Resume Next
    nm.RefersToRange.Value = _
        Chr(10003) & " Q" & nQ & " cargada - " & _
        nLAFT & " LAFT / " & nREL & " vinculados" & vbCrLf & _
        "(" & nombreArchivo & ")"
    On Error GoTo 0
End Sub

Private Function NombreDesdeRuta(ByVal ruta As String) As String
    Dim pos As Long

    pos = InStrRev(ruta, "\")
    If pos = 0 Then pos = InStrRev(ruta, "/")

    If pos = 0 Then
        NombreDesdeRuta = ruta
    Else
        NombreDesdeRuta = Mid(ruta, pos + 1)
    End If
End Function

Private Function LO_LAFT(ByVal nQ As Integer) As String
    LO_LAFT = IIf(nQ = 1, LO_Q1_LAFT, LO_Q2_LAFT)
End Function

Private Function LO_REL(ByVal nQ As Integer) As String
    LO_REL = IIf(nQ = 1, LO_Q1_REL, LO_Q2_REL)
End Function