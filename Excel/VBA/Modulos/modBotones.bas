Option Explicit

' ============================================================
' modBotones
' Crea los botones de la hoja Panel y los conecta a sus macros.
' Ejecutar InstalarBotones una sola vez despues de crear el libro,
' o cuando los botones se desconecten por algun motivo.
' ============================================================

Public Sub InstalarBotones()
    Dim ws As Worksheet
    Dim paso As String
    Dim shp As Shape

    On Error GoTo ErrHandler

    paso = "Obtener hoja Panel"
    If Not HojaExiste(SH_PANEL) Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SH_PANEL
    Else
        Set ws = ThisWorkbook.Worksheets(SH_PANEL)
    End If

    Application.ScreenUpdating = False

    paso = "Limpiar botones previos"
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    paso = "Estilo del panel"
    ws.Cells.Clear
    ws.Cells.Font.Name = FuenteCorporativa()
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 6
    ws.Columns("D").ColumnWidth = 55
    ws.Rows.RowHeight = 18

    ' Fondo blanco general
    ws.Cells.Interior.Color = ColorBlanco()
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    paso = "Insertar banner corporativo"
    InsertarBanner ws

    paso = "Subtitulo"
    With ws.Range("B6:D6")
        .Merge
        .Value = "BDNE - Transformacion de Incrementales"
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = ColorNegro()
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("B7:D7")
        .Merge
        .Value = "Fondos SURA SAF S.A.C."
        .Font.Size = 10
        .Font.Color = ColorGrisMedio()
        .HorizontalAlignment = xlCenter
    End With

    EtiquetaSeccion ws, 9, "Carga de datos"
    EtiquetaSeccion ws, 15, "Proceso"
    EtiquetaSeccion ws, 19, "Mantenimiento"

    paso = "Crear Named Ranges de estado"
    CrearNombreEstado ST_Q1_ESTADO, ws, "D10"
    CrearNombreEstado ST_Q2_ESTADO, ws, "D12"
    CrearNombreEstado ST_OUT_ESTADO, ws, "D16"

    ws.Range("D10").Value = Chr(10007) & " Sin datos"
    ws.Range("D12").Value = Chr(10007) & " Sin datos"
    ws.Range("D16").Value = Chr(10007) & " Sin datos"

    With ws.Range("D10,D12,D16")
        .Font.Size = 9
        .Font.Color = ColorGrisMedio()
        .WrapText = True
    End With

    ws.Rows("10").RowHeight = 30
    ws.Rows("12").RowHeight = 30
    ws.Rows("16").RowHeight = 30

    paso = "Crear boton Cargar Q1"
    CrearBoton ws, "B10", "Cargar Quincena 1", "modImportarDatos.CargarQuincena1", _
               ColorAzul(), ColorBlanco()

    paso = "Crear boton Cargar Q2"
    CrearBoton ws, "B12", "Cargar Quincena 2", "modImportarDatos.CargarQuincena2", _
               ColorAzul(), ColorBlanco()

    paso = "Crear boton Generar BDNE"
    CrearBoton ws, "B16", "Generar BDNE", "modTransformar.GenerarBDNE", _
               ColorNegro(), ColorBlanco()

    paso = "Crear boton Eliminar"
    CrearBoton ws, "B20", "Eliminar Datos", "modEliminarDatos.EliminarDatos", _
               RGB(192, 0, 0), ColorBlanco()

    ws.Activate
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox "Panel instalado correctamente." & vbCrLf & _
           "Los cuatro botones estan listos.", _
           vbInformation, "InstalarBotones"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "[" & paso & "] " & Err.Description, vbCritical, "Error al instalar botones"
End Sub

' -- Inserta el banner corporativo SURA en el header --------
' El PNG va embebido en base64 (autocontenido, sin archivos
' externos). Se decodifica a un PNG temporal y se inserta.
Private Sub InsertarBanner(ByVal ws As Worksheet)
    Dim b64 As String
    Dim rutaTmp As String
    Dim pic As Object
    Dim anchoBanner As Double
    Dim altoBanner As Double

    On Error GoTo SinBanner

    b64 = BannerB64()
    rutaTmp = Environ$("TEMP") & "\sura_banner.png"

    ' Decodificar base64 a archivo PNG temporal
    If Not GuardarBase64ComoArchivo(b64, rutaTmp) Then GoTo SinBanner

    ' Banda negra de fondo detras del banner (filas 1-5)
    With ws.Range("A1:E5").Interior
        .Color = ColorNegro()
    End With
    ws.Rows("1:5").RowHeight = 22

    ' Insertar la imagen centrada en la banda
    Set pic = ws.Pictures.Insert(rutaTmp)
    anchoBanner = 360
    altoBanner = anchoBanner * 240 / 900   ' proporcion 900x240
    With pic
        .Width = anchoBanner
        .Height = altoBanner
        .Top = ws.Range("A1").Top + 8
        .Left = ws.Range("A1").Left + (ws.Range("A1:E1").Width - anchoBanner) / 2
        .Placement = xlMoveAndSize
    End With
    Exit Sub

SinBanner:
    ' Si falla la imagen, dejar al menos la banda negra con texto
    On Error Resume Next
    ws.Range("A1:E5").Interior.Color = ColorNegro()
    With ws.Range("B3:D3")
        .Merge
        .Value = "SURA INVESTMENTS"
        .Font.Name = FuenteCorporativa()
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = ColorBlanco()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo 0
End Sub

' -- Decodifica un string base64 y lo guarda como archivo ----
Private Function GuardarBase64ComoArchivo(ByVal b64 As String, ByVal ruta As String) As Boolean
    Dim xmlDoc As Object
    Dim nodo As Object
    Dim stream As Object

    On Error GoTo Falla

    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set nodo = xmlDoc.createElement("b64")
    nodo.DataType = "bin.base64"
    nodo.Text = b64

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1   ' binario
    stream.Open
    stream.Write nodo.nodeTypedValue
    stream.SaveToFile ruta, 2   ' sobrescribir
    stream.Close

    GuardarBase64ComoArchivo = True
    Exit Function

Falla:
    GuardarBase64ComoArchivo = False
End Function

' -- Devuelve el PNG del banner en base64 --------------------
Private Function BannerB64() As String
    Dim s As String
    s = s & "iVBORw0KGgoAAAANSUhEUgAAA4QAAADwCAIAAADXdz0mAABacElEQVR42u3dd3xT5f4H8DOTnOykkwKlyB6FsltANgoyZQjKUMGBXjeiF+d1K66feJUl4ATBsqcspWUre5aWFjqhM0mzc8bvjwdyY6GlQkHG5/3y5aukp8nJk7T5nGd8HzoyMpoCAAAAAPgnMGgCAAAAAEAYBQAAAACEUQAAAAAAhFEAAAAAQBgFAAAAAEAY"
    s = s & "BQAAAACEUQAAAAAAhFEAAAAAQBgFAAAAAEAYBQAAAACEUQAAAAAAhFEAAAAAQBgFAAAAAEAYBQAAAACEUQAAAAAAhFEAAAAAQBgFAAAAAIRRAAAAAACEUQAAAABAGAUAAAAAQBgFAAAAAIRRAAAAAACEUQAAAABAGAUAAAAAQBgFAAAAAIRRAAAAAACEUQAAAABAGAUAAAAAQBgFAAAAAIRRAAAAAACEUQAAAABAGAUAAAAA"
    s = s & "QBgFAAAAAIRRAAAAAEAYBQAAAABAGAUAAAAAhFEAAAAAAIRRAAAAAEAYBQAAAABAGAUAAAAAhFEAAAAAAIRRAAD4h9A0jUYAAIRRAAD4Z0iSxDAMw+ADBQAQRgEA4PpSFEUQNG632+v1cRyHBgEAhFEAALhOWJZ1u90fffTRjBlfhYeHlZaW0jSNLlIAQBgFAIDrRFEUvV7/wAMP/Pbb5kcemej3B1wuF8uymEgKAAijAABw"
    s = s & "PdhsNkmS6tSp8+WX05ctS27fvn1ZWZkoiizLonEAAGEUAACuIVmWz507x7JsIBCQJKl79+7r16/96KMPDQZjWZkNC5sA4LJYnU6PVgAAgCvAMIzX661Vq9aQIYNpmmZZVpZljuM6deo0cOAAh8N+8OChQCCg0WgURUFzAQDCKAAA1HgepX0+77hxY1UqlaIoDMMoiiLLcnh42ODBgxMSWqelpWVmZvI8z3EcIikAIIwCAEBN"
    s = s & "4nm+oOBsmzZtmjRpIssywzBkQb0sy4qiNG7c6P777zcYDHv37istLRUEgaZpRFIAQBgFAICawTCM3+8vKiq6//7RFEXRNE3W0ZMvJElSqVSdOycNGHBPaWnpwYOHJEnCqD0AIIwCAEDNUBRFrVanp6fXq1cvIaF1hUX0wVH7iIiIoUOHxse3PHHieGZmFkbtAQBhFAAAagzDMCkpqd27d69Tp7YoisH+UYqiQkftmzZtMnr0"
    s = s & "KK1W2Ldvv81mEwQNRu0BAGEUAACuFsdxHo9n9eo1zZo1bdSoERmgrxBJyY0ajaZr1679+vUrKio8dOiIoshqtVqWZbQhAMIoAADAFVIUhed5l8uVnLykvLy8TZs2Wq324kgaHLWPioocPnxY06ZNDx06lJ2drdFoWJZFFykAwigAAMCV51GO4xiG2bLlt/Xrf9VqhSZNmqhUqgqRNHTUvnnz5iNHjqAoeu/evW63WxAE5FEA"
    s = s & "hFEAAIArz6MURen1+sLCwuXLV2zcuEkQNE2bNuU47uJISm7R6XS9e/fq3r1bZmbm8eMnOI7jeR6RFABhFAAA4ArJsszzvCAIubl5y5Yt37x5s0ajady4Ec/zFSJpcNS+bt26998/OiIiYu/evcXFxShHCoAwCgAAUI2PEPbS+84riqIoikqlEgQhJydn+fIVGzZspGm6YcMGGo0mNJIGR+1Zlm3fvv2QIYNLSkoOHjwkSSLK"
    s = s & "kQIgjAIAAFRKlmWXy8VxHBmFvzg4BiOpRqPJy8tbtWrVunXrfD5/gwYN9HodiaTUhSF7RaFkWbJarUOGDImPb3X48OEzZ86o1WosbAJAGAUAALhEEjUY9J07dy4qKiouLqYoSqVSMcwlOkpDe0kLCwvXrVu3YsUKm80WF1fPYrEEIynD/G9hU5MmjUePHsUw9J9/7nU6nYIgoMEBEEYBAADOo2laFMXw8Ij169fee+/Q6Ojo"
    s = s & "srKy/Px8l8vFsiyZG1qhO5NEUjKX1G63b9q0ZdmyZQUFZ+vViw0PD6dpmsRQsrW9JEmCIPTs2bN3716ZmVknTpwgd4suUgCEUQAAAIqmKZpmPB7PwIEDmzVr2rVrlzFjHujUqSPP8/n5BcXFRWRL+tB9QUMjKcuyOp3W5XKnpqYuXpx86tSpWrWiY2JiSIQl80fJF7Vr177//tFRUVH79u0rLCzEwiYAhFEAAACKoiiO4xwO"
    s = s & "R+fOSc2aNRVFUa1WN2jQYNCggcOG3dugwR1OpzMvL8/hcDAMc8nhe1lWWJbVarWBQGDPnj2//JJ8+PCRiIjwevXqkYNlWSZr7Wmabteu3dChQ2w22/79B8geTsijAAijAABwW2MYxuPxms3mQYMGkn+SqZ9ms6l9+3bjxo3p0aOHyWQqLCwqKCjw+Xw8z1+8zomMy5Na9wcOHFy8+Jfdu/cYjcYGDRqwLEvG7sn/LRbL4MGD"
    s = s & "4uPjjx49lpWVhYVNAAijAABwu+M49uzZs/fdd59eryeD78FxdoZhateu3adPn/vvH926dWtJknJzc0tLSxXlEuucSKbUagWWZU+cOLF06bKtW1M0GnWDBg3IBk7KeVTTpk1Gjx7FMMzevfvKy8uxsAkAYRQAahipuRjEsixzJc7fCU0zwXxQzRMgDxq6mfg1/OvDstftsSpr7dAGZ1k22HSXb2KaqcGeuSs+jeqdJ0PTzLXo"
    s = s & "ROR5/ty5cxqNpkePHpIkkYhJngtFUbIsy7Ks1WqbN28+YsTwwYMHxcTElJWV5eXlu1wujqu4zol8IQgCz/OZmZnLli3fsGGDLMuNGjUUBIFhGEWRZVkWBKFnzx49evQ4cyb7+PHjWNgEcCt89kVGRqMVAP5BZPkwRVGSJAUCAVEU5Quq+ISlaUpRqMumOEWhGIY2GAzVCXw0TbtcLr/fzzAM2ZVRrVYHK+/UOKfTKYri9Xms"
    s = s & "i5MfRVGiKJIGlySJtPbfSsWKQul0WpVKdWVJKHgawdddkiRJkv7uaVTz0QwG/SXrLl09WZaXL1/apUsXURQ5jruolRQy1E4ePRAIpKamJicv2bRpU05OLsdxWq2WZVmylL7CL4Xb7fb7/U2bNnnwwfGjR4+OiooirxpZla8oypw533z88cd5efkmk4mcCf6eACCMAsDfyyKyLHu9Xp/PxzCM0WgMDw+PioqKiooMCwvTarUc"
    s = s & "xwV7myr8P/QLiiK3h35BvqswDOt0Or/77nu/3191HqVp2ufzDR48uFGjhkVFRdnZOenp6Tk5OZIkGQyG86OkNer++0ebzZaSkuKcnJyMjIzs7BxZlq/RYxGkL9Pr9Xq9XoZhLBZLdHRUrVoxtWrVMpvNWq3A83w1ExsZlV61avXRo0fVanX1T5i87oqi+Hw+r9dLUbTRaIiOjoqJiYmJibFawwRBQ3J5jTxlsgAoEAh89933"
    s = s & "drv94uXtV38p5fP5IiIili5d0qJFc1EUSW/3JTMraTTyz4KCs2vXrlm6dNkff/xJBtw1Gk3wsNBI6vF4vF5vXFzc6NGjHnxwfFxcHImkNE2zLHvmzJl3331v0aLFDMPqdFpJktBLCoAwCgDVSkWBQMDlcqnV6iZNmiQlJXXq1LFlyxa1a9e2WCw1+1g2my0hoS0pAFnF5zTLsjab7ZdfFg8YcA+5paSkZPfuPT/++NPatWtZ"
    s = s & "llWpVDXb87Rz54477qhPvi4uLt6z548ff/xxzZq1PM/zPF+zj0X63srLy1mWbdy40Z133tmlS5eWLVvWrVtHp9Nd8d0++uhjCxYsMJnM1ezQZVlWFEWXy8UwTIMGdyQlJXXu3DkhISE2ti7p27t22rZtT9b91HhWY1nW7XZHRUV9/fVXvXv3oiiKbPVZWQF88soGU+nevfuWLl26Zs3ajIxTiiLrdDqO44KHUeeL4bNer9ft"
    s = s & "dkdHRw8fPuzhhx9u0aI5RVGBQIDneYqiVq9e/fbb7xw6dNhkMrEsex361wEAYRTgZkV6Qx0OR1RU1MCBA+67b2THjh1Jn1DoB3aN9A6SdSQHDhwYMGCQoihkZL+KSGG32+fNmzt8+HBZlsh8PvKtlStXvfDC5JKSEo1GU4MZcf36da1axQcXvpAblyxZOnnyi+Xl5TWVfUkqcjgcarW6b9++48eP6969W2gAvbLWJj2OI0fe"