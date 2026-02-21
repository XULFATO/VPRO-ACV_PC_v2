Attribute VB_Name = "VCA_ESP"
Option Explicit

'==========================================================================
' MÓDULO : VCA_ESP
' FUNCIÓN: Todo el proceso VCA España – generación y validación.
'          100% independiente de Portugal.
' DEPENDE: Comun_Constantes, Comun_Buscador, Comun_GestorPestanas
'==========================================================================
'
' CABECERAS EN "Contabilidad_Cuentas":
'   "ENLACE CONTABLE" → celda combinada → colEnlace
'   "Debe (cliente)"  → nos da filaCabecera y colDebe
'   "Haber (cliente)" → misma fila → colHaber
'   "STANDARD DEBE"   → misma fila → colStdDebe
'   "STANDARD HABER"  → misma fila → colStdHaber
'   Primera fila datos = filaCabecera + 1
'
' Todas las búsquedas usan UCase → no distingue may/min.
'==========================================================================

' Hojas que gestiona este proceso (para el gestor de pestañas)
Private Function ESP_HojasDelProceso() As Variant
    ESP_HojasDelProceso = Array(HOJA_DATOS_ESP, HOJA_LINEAS)
End Function

' Estructura de columnas ESP (se rellena dinámicamente)
Private Type EstructuraESP
    filaCabecera As Long
    filaInicio   As Long
    colEnlace    As Long
    colDebe      As Long
    colHaber     As Long
    colStdDebe   As Long
    colStdHaber  As Long
    valida       As Boolean
End Type


'==========================================================================
' ESP_ObtenerEstructura
' Localiza dinámicamente todas las columnas necesarias en ESP.
'==========================================================================
Private Function ESP_ObtenerEstructura() As EstructuraESP
    Dim ws       As Worksheet
    Dim result   As EstructuraESP
    Dim celdaRef As Range
    Dim celdaAux As Range

    result.valida = False

    Set ws = ObtenerHojaSegura(HOJA_DATOS_ESP, ThisWorkbook)
    If ws Is Nothing Then
        ESP_ObtenerEstructura = result
        Exit Function
    End If

    ' Buscar "Debe (cliente)" → fila de cabecera y colDebe
    Set celdaRef = BuscarCeldaPorTexto(ws, "Debe (cliente)", xlWhole)
    If celdaRef Is Nothing Then
        MsgBox "No se encontró 'Debe (cliente)' en '" & HOJA_DATOS_ESP & "'.", _
               vbCritical, "Cabecera No Encontrada"
        ESP_ObtenerEstructura = result
        Exit Function
    End If
    result.filaCabecera = celdaRef.Row
    result.filaInicio   = result.filaCabecera + 1
    result.colDebe      = ObtenerColumnaReal(celdaRef)

    ' Buscar "Haber (cliente)" en la misma fila
    Set celdaAux = BuscarEnFila(ws, result.filaCabecera, "Haber (cliente)", xlWhole)
    If celdaAux Is Nothing Then
        MsgBox "No se encontró 'Haber (cliente)' en la fila de cabecera.", _
               vbCritical, "Cabecera No Encontrada"
        ESP_ObtenerEstructura = result
        Exit Function
    End If
    result.colHaber = ObtenerColumnaReal(celdaAux)

    ' Buscar "STANDARD DEBE" en la misma fila
    Set celdaAux = BuscarEnFila(ws, result.filaCabecera, "STANDARD DEBE", xlWhole)
    If celdaAux Is Nothing Then
        MsgBox "No se encontró 'STANDARD DEBE' en la fila de cabecera.", _
               vbCritical, "Cabecera No Encontrada"
        ESP_ObtenerEstructura = result
        Exit Function
    End If
    result.colStdDebe = ObtenerColumnaReal(celdaAux)

    ' Buscar "STANDARD HABER" en la misma fila
    Set celdaAux = BuscarEnFila(ws, result.filaCabecera, "STANDARD HABER", xlWhole)
    If celdaAux Is Nothing Then
        MsgBox "No se encontró 'STANDARD HABER' en la fila de cabecera.", _
               vbCritical, "Cabecera No Encontrada"
        ESP_ObtenerEstructura = result
        Exit Function
    End If
    result.colStdHaber = ObtenerColumnaReal(celdaAux)

    ' Buscar "ENLACE CONTABLE" → celda combinada
    Set celdaAux = BuscarCeldaPorTexto(ws, "ENLACE CONTABLE", xlPart)
    If celdaAux Is Nothing Then
        MsgBox "No se encontró 'ENLACE CONTABLE' en '" & HOJA_DATOS_ESP & "'.", _
               vbCritical, "Cabecera No Encontrada"
        ESP_ObtenerEstructura = result
        Exit Function
    End If
    result.colEnlace = ObtenerColumnaReal(celdaAux)

    result.valida = True
    ESP_ObtenerEstructura = result
End Function


'==========================================================================
' ESP_LimpiarComentariosValidacion
' Limpia SOLO las columnas dinámicas ESP (Enlace, Debe, Haber).
'==========================================================================
Public Sub ESP_LimpiarComentariosValidacion()
    Dim ws  As Worksheet
    Dim est As EstructuraESP
    Dim c   As Range

    Set ws = ObtenerHojaSegura(HOJA_DATOS_ESP, ThisWorkbook)
    If ws Is Nothing Then Exit Sub

    est = ESP_ObtenerEstructura()
    If Not est.valida Then Exit Sub

    For Each c In ws.UsedRange
        If Not c.Comment Is Nothing Then
            If Left(c.Comment.Text, Len(PREFIJO_VAL)) = PREFIJO_VAL Then
                c.Comment.Delete
                Select Case c.Column
                    Case est.colEnlace, est.colDebe, est.colHaber
                        c.Interior.Color = xlNone
                End Select
            End If
        End If
    Next c

    ws.Columns(est.colEnlace).Interior.Color = xlNone
    ws.Columns(est.colDebe).Interior.Color   = xlNone
    ws.Columns(est.colHaber).Interior.Color  = xlNone
End Sub


'==========================================================================
' ESP_MarcarError  (privada)
'==========================================================================
Private Sub ESP_MarcarError(ByVal celda   As Range, _
                              ByVal color   As Long, _
                              ByVal mensaje As String)
    If celda.Interior.Color <> vbRed Then
        celda.Interior.Color = color
    End If
    AddOrAppendComment celda, mensaje
End Sub


'==========================================================================
' Generar_LINEASVCA_ESP  (entrada principal)
'==========================================================================
Public Sub Generar_LINEASVCA_ESP()

    Dim wsC       As Worksheet
    Dim wsT       As Worksheet
    Dim est       As EstructuraESP
    Dim cliente   As String, release As String
    Dim filaNueva As Long, contador As Long
    Dim ultFila   As Long, i As Long
    Dim valEnlace As String, valD As String, valF As String
    Dim ruta      As String

    ' 1 ── Verificar y gestionar ejecución anterior ────────────────────────
    If Not GestorPestanas_VerificarAntesDeEjecutar("VCA España", ESP_HojasDelProceso()) Then
        Exit Sub
    End If

    ' 2 ── Hoja de datos y estructura de columnas ──────────────────────────
    Set wsC = ObtenerHojaSegura(HOJA_DATOS_ESP, ThisWorkbook)
    If wsC Is Nothing Then Exit Sub

    est = ESP_ObtenerEstructura()
    If Not est.valida Then Exit Sub

    ' 3 ── Validación previa opcional ──────────────────────────────────────
    If MsgBox("¿Aplicar validaciones antes de generar?", _
              vbYesNo + vbQuestion, "Validaciones ESP") = vbYes Then
        If Not Validar_Contabilidad_ESP() Then
            MsgBox "Corrige los errores marcados y vuelve a intentarlo.", _
                   vbCritical, "Proceso Cancelado"
            Exit Sub
        End If
        MsgBox "Validación correcta. Continuando...", vbInformation, "OK"
    End If

    ' 4 ── Datos del usuario ───────────────────────────────────────────────
    cliente = PedirCliente(PAC_ESP)
    If cliente = "" Then Exit Sub
    release = PedirRelease()
    If release = "" Then Exit Sub

    ' 5 ── Activar modo rendimiento (pantalla oculta desde aquí) ──────────
    PrepararCarpetasBase
    wsC.AutoFilterMode = False
    ActivarModoRendimiento

    ' 6 ── Nueva hoja LINEASVCA ────────────────────────────────────────────
    EliminarHojaSiExiste HOJA_LINEAS, ThisWorkbook
    Set wsT  = ThisWorkbook.Worksheets.Add
    wsT.Name = HOJA_LINEAS
    CrearCabecerasLINEASVCA wsT
    filaNueva = 2
    contador  = 5

    ' 7 ── Bucle de generación ─────────────────────────────────────────────
    ultFila = wsC.Cells(wsC.Rows.Count, est.colEnlace).End(xlUp).Row

    For i = est.filaInicio To ultFila
        valEnlace = Trim(wsC.Cells(i, est.colEnlace).Value)
        If valEnlace <> "" Then
            valD = Trim(wsC.Cells(i, est.colDebe).Value)
            valF = Trim(wsC.Cells(i, est.colHaber).Value)
            If (valD <> "" Or valF <> "") And _
               InStr(valD, " ") = 0 And InStr(valF, " ") = 0 Then
                EscribirLineaVCA wsT, filaNueva, TIPO_ESP, cliente, PAC_ESP, _
                                 release, contador, valEnlace, valD, valF
                contador  = contador + 5
                filaNueva = filaNueva + 1
            End If
        End If
    Next i

    ' 8 ── Formato y reordenar pestañas (pantalla AÚN oculta → sin parpadeo)
    AplicarFormatoTablaLINEASVCA wsT
    GestorPestanas_ReordenarVCA HOJA_ESP, HOJA_DATOS_ESP, RGB(255, 159, 156), HOJA_POR

    ' 9 ── Restaurar DESPUÉS de mover pestañas ────────────────────────────
    RestaurarModoNormal

    ' 10 ── Guardar ────────────────────────────────────────────────────────
    ruta = ObtenerRutaVersionada(RUTA_BASE, "VCA_" & cliente & "_" & release & ".xls")
    GuardarHojaComoXLS wsT, ruta, filaNueva - 2
End Sub


'==========================================================================
' Validar_Contabilidad_ESP
'==========================================================================
Public Function Validar_Contabilidad_ESP() As Boolean

    Dim ws        As Worksheet
    Dim est       As EstructuraESP
    Dim i         As Long, uFilaMax As Long
    Dim valEnlace As String, valD As String, valF As String
    Dim valE      As String, valG As String
    Dim especiales As Variant, v As Variant
    Dim hayError  As Boolean, nErrores As Long, lista As String

    Set ws = ObtenerHojaSegura(HOJA_DATOS_ESP, ThisWorkbook)
    If ws Is Nothing Then
        Validar_Contabilidad_ESP = False
        Exit Function
    End If

    est = ESP_ObtenerEstructura()
    If Not est.valida Then
        Validar_Contabilidad_ESP = False
        Exit Function
    End If

    ESP_LimpiarComentariosValidacion

    especiales = Array("071", "115", "125", "126", "127")
    uFilaMax   = ws.Cells(ws.Rows.Count, est.colEnlace).End(xlUp).Row
    hayError   = False
    nErrores   = 0
    lista      = ""

    For i = est.filaInicio To uFilaMax
        valEnlace = Trim(ws.Cells(i, est.colEnlace).Text)
        valD      = Trim(ws.Cells(i, est.colDebe).Text)
        valF      = Trim(ws.Cells(i, est.colHaber).Text)
        valE      = Trim(ws.Cells(i, est.colStdDebe).Text)
        valG      = Trim(ws.Cells(i, est.colStdHaber).Text)

        ' R1: Sin espacios
        If valD <> "" Or valF <> "" Then
            If InStr(valD, " ") <> 0 Then
                ESP_MarcarError ws.Cells(i, est.colDebe), RGB(255, 189, 180), _
                    "ERROR – No puede contener espacios. Se descartará"
            End If
            If InStr(valF, " ") <> 0 Then
                ESP_MarcarError ws.Cells(i, est.colHaber), RGB(255, 189, 180), _
                    "ERROR – No puede contener espacios. Se descartará"
            End If
            If InStr(valD, " ") <> 0 Or InStr(valF, " ") <> 0 Then
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – Espacios en Debe/Haber." & vbCrLf
            End If
        End If

        ' R2: STANDARD informado → Debe y Haber obligatorios
        If valE <> "" And valG <> "" Then
            If valD = "" Then ESP_MarcarError ws.Cells(i, est.colDebe), RGB(255, 165, 0), _
                "OBLIGATORIO – STANDARD informado: se requiere Debe"
            If valF = "" Then ESP_MarcarError ws.Cells(i, est.colHaber), RGB(255, 165, 0), _
                "OBLIGATORIO – STANDARD informado: se requiere Haber"
            If valD = "" Or valF = "" Then
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – STANDARD: faltan Debe y/o Haber." & vbCrLf
            End If
        End If

        ' R3: Enlace especial → Debe y Haber obligatorios
        Dim esEsp As Boolean: esEsp = False
        For Each v In especiales
            If UCase(valEnlace) = UCase(CStr(v)) Then esEsp = True: Exit For
        Next v
        If esEsp Then
            If valD = "" Then ESP_MarcarError ws.Cells(i, est.colDebe), RGB(173, 216, 230), _
                "OBLIGATORIO – Enlace especial " & valEnlace & ": se requiere Debe"
            If valF = "" Then ESP_MarcarError ws.Cells(i, est.colHaber), RGB(173, 216, 230), _
                "OBLIGATORIO – Enlace especial " & valEnlace & ": se requiere Haber"
            If valD = "" Or valF = "" Then
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – Enlace " & valEnlace & ": faltan Debe/Haber." & vbCrLf
            End If
        End If

        ' R4: Enlace > 500
        If IsNumeric(valEnlace) Then
            If CLng(valEnlace) > 500 Then
                ESP_MarcarError ws.Cells(i, est.colEnlace), RGB(148, 0, 211), _
                    "ERROR – Enlace " & valEnlace & " supera el máximo (500)"
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – Enlace " & valEnlace & " > 500." & vbCrLf
            End If
        End If
    Next i

    If hayError Then
        MsgBox "Se detectaron " & nErrores & " error(es) en '" & HOJA_DATOS_ESP & "':" & _
               vbCrLf & vbCrLf & lista & vbCrLf & _
               "Revisa los comentarios y colores en la hoja.", _
               vbCritical, "Errores de Validación ESP"
        Validar_Contabilidad_ESP = False
    Else
        Validar_Contabilidad_ESP = True
    End If
End Function
