Option Explicit

' ============================================================
' modEliminarDatos
' Punto de entrada del boton "Eliminar Datos".
' Elimina las hojas internas y limpia las celdas de estado.
' ============================================================

Public Sub EliminarDatos()
    Dim paso As String
    Dim hayDatos As Boolean
    Dim resp As Integer
    Dim candidatos(4) As String
    Dim aEliminar() As String
    Dim count As Integer
    Dim i As Integer

    On Error GoTo ErrHandler

    paso = "Verificar datos existentes"

    hayDatos = HojaExiste(SH_Q1_LAFT) Or _
               HojaExiste(SH_Q1_REL) Or _
               HojaExiste(SH_Q2_LAFT) Or _
               HojaExiste(SH_Q2_REL) Or _
               HojaExiste(SH_OUTPUT)

    If Not hayDatos Then
        MsgBox "No hay datos que eliminar.", vbInformation, "Sin datos"
        Exit Sub
    End If

    paso = "Primera confirmacion"
    resp = MsgBox("Esta accion eliminara todos los datos cargados " & _
                  "y el resultado BDNE generado." & vbCrLf & vbCrLf & _
                  Chr(191) & "Deseas continuar?", _
                  vbQuestion + vbYesNo, "Confirmar eliminacion")

    If resp <> vbYes Then Exit Sub

    paso = "Segunda confirmacion"
    resp = MsgBox("Segunda confirmacion requerida." & vbCrLf & vbCrLf & _
                  "Se eliminaran:" & vbCrLf & _
                  "  - Datos Q1 (LAFT + PEP RELACIONADOS)" & vbCrLf & _
                  "  - Datos Q2 (LAFT + PEP RELACIONADOS)" & vbCrLf & _
                  "  - Hoja BDNE_OUTPUT" & vbCrLf & vbCrLf & _
                  Chr(191) & "Confirmas?", _
                  vbExclamation + vbYesNo, "Confirmar eliminacion")

    If resp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    paso = "Recopilar hojas"

    candidatos(0) = SH_Q1_LAFT
    candidatos(1) = SH_Q1_REL
    candidatos(2) = SH_Q2_LAFT
    candidatos(3) = SH_Q2_REL
    candidatos(4) = SH_OUTPUT

    count = 0

    For i = 0 To 4
        If HojaExiste(candidatos(i)) Then
            ReDim Preserve aEliminar(count)
            aEliminar(count) = candidatos(i)
            count = count + 1
        End If
    Next i

    paso = "Eliminar hojas"

    For i = 0 To count - 1
        ThisWorkbook.Worksheets(aEliminar(i)).Delete
    Next i

    paso = "Limpiar estado del panel"

    LimpiarEstado ST_Q1_ESTADO
    LimpiarEstado ST_Q2_ESTADO
    LimpiarEstado ST_OUT_ESTADO

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Datos eliminados correctamente.", vbInformation, "Eliminacion completada"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Error al eliminar datos" & vbCrLf & vbCrLf & _
           "Paso:        " & paso & vbCrLf & _
           "Descripcion: " & Err.Description & vbCrLf & _
           "Numero:      " & Err.Number, _
           vbCritical, "Error al eliminar datos"
End Sub

Private Sub LimpiarEstado(ByVal nmEstado As String)
    On Error Resume Next
    ThisWorkbook.Names(nmEstado).RefersToRange.Value = Chr(10007) & " Sin datos"
    On Error GoTo 0
End Sub