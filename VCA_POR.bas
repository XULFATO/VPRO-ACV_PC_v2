Attribute VB_Name = "VCA_POR"
Option Explicit

'==========================================================================
' MÓDULO : VCA_POR
' FUNCIÓN: Todo el proceso VCA Portugal – generación y validación.
'          100% independiente de España.
' DEPENDE: Comun_Constantes, Comun_Buscador, Comun_GestorPestanas
'==========================================================================
'
' CABECERAS EN "Analisis Conceitos":
'   "ENLACE CONTABLE" → localizado por texto → colBase (dinámica)
'   colBase+1         → Debe Cliente
'   colBase+2         → Haber Cliente
'   "TIPO CONCEPTO"   → localizado por texto → colTC (dinámica)
'
' LÓGICA PRIMERA FILA DE DATOS (via Comun_Buscador):
'   1. Buscar "ENLACE CONTABLE"
'   2. Bajar 1 fila
'   3. Si vacía → bajar 1 más
'   4. Si azul  → es subcabecera, bajar 1 más
'   5. ESA es la primera fila de datos
'==========================================================================

' Hojas que gestiona este proceso
Private Function POR_HojasDelProceso() As Variant
    POR_HojasDelProceso = Array(HOJA_LINEAS)
End Function


'==========================================================================
' POR_Celda_Inicial
' Usa Comun_Buscador para localizar la primera fila de datos.
'==========================================================================
Public Function POR_Celda_Inicial() As Range
    Dim ws       As Worksheet
    Dim celdaRef As Range
    Dim filaDatos As Long
    Dim col       As Long

    Set POR_Celda_Inicial = Nothing

    Set ws = ObtenerHojaSegura(HOJA_DATOS_POR, ThisWorkbook)
    If ws Is Nothing Then Exit Function

    ' Buscar "ENLACE CONTABLE"
    Set celdaRef = BuscarCeldaPorTexto(ws, "ENLACE CONTABLE", xlPart)
    If celdaRef Is Nothing Then
        MsgBox "No se encontró 'ENLACE CONTABLE' en '" & HOJA_DATOS_POR & "'.", _
               vbCritical, "Cabecera No Encontrada"
        Exit Function
    End If

    col = ObtenerColumnaReal(celdaRef)

    ' BuscarPrimeraFilaDatos salta vacías y azules
    filaDatos = BuscarPrimeraFilaDatos(ws, ObtenerFilaReal(celdaRef), col, saltarAzul:=True)

    Set POR_Celda_Inicial = ws.Cells(filaDatos, col)
End Function


'==========================================================================
' POR_LimpiarComentariosValidacion
' Limpia SOLO las columnas dinámicas POR (colBase, +1, +2).
'==========================================================================
Public Sub POR_LimpiarComentariosValidacion()
    Dim ws       As Worksheet
    Dim celdaIni As Range
    Dim colBase  As Long
    Dim c        As Range

    Set ws = ObtenerHojaSegura(HOJA_DATOS_POR, ThisWorkbook)
    If ws Is Nothing Then Exit Sub

    Set celdaIni = POR_Celda_Inicial()
    If celdaIni Is Nothing Then Exit Sub

    colBase = celdaIni.Column

    For Each c In ws.UsedRange
        If Not c.Comment Is Nothing Then
            If Left(c.Comment.Text, Len(PREFIJO_VAL)) = PREFIJO_VAL Then
                c.Comment.Delete
                Select Case c.Column
                    Case colBase, colBase + 1, colBase + 2
                        c.Interior.Color = xlNone
                End Select
            End If
        End If
    Next c

    ws.Columns(colBase).Interior.Color     = xlNone
    ws.Columns(colBase + 1).Interior.Color = xlNone
    ws.Columns(colBase + 2).Interior.Color = xlNone
End Sub


'==========================================================================
' POR_MarcarError  (privada)
'==========================================================================
Private Sub POR_MarcarError(ByVal celda   As Range, _
                              ByVal color   As Long, _
                              ByVal mensaje As String)
    If celda.Interior.Color <> vbRed Then
        celda.Interior.Color = color
    End If
    AddOrAppendComment celda, mensaje
End Sub


'==========================================================================
' Generar_LINEASVCA_POR  (entrada principal)
'==========================================================================
Public Sub Generar_LINEASVCA_POR()

    Dim wsC          As Worksheet
    Dim wsT          As Worksheet
    Dim celdaIni     As Range
    Dim colBase      As Long, ultFila As Long
    Dim cliente      As String, release As String
    Dim filaNueva    As Long, contador As Long
    Dim i            As Long
    Dim ruta         As String
    Dim comboMayoria As Object
    Dim clave        As Variant
    Dim partes()     As String
    Dim valDebe      As String, valHaber As String

    ' 1 ── Verificar y gestionar ejecución anterior ────────────────────────
    If Not GestorPestanas_VerificarAntesDeEjecutar("VCA Portugal", POR_HojasDelProceso()) Then
        Exit Sub
    End If

    ' 2 ── Hoja de datos y celda inicial ───────────────────────────────────
    Set wsC = ObtenerHojaSegura(HOJA_DATOS_POR, ThisWorkbook)
    If wsC Is Nothing Then Exit Sub

    Set celdaIni = POR_Celda_Inicial()
    If celdaIni Is Nothing Then Exit Sub

    colBase = celdaIni.Column
    ultFila = wsC.Cells(wsC.Rows.Count, colBase).End(xlUp).Row

    ' 3 ── Validación previa opcional ──────────────────────────────────────
    If MsgBox("¿Aplicar validaciones antes de generar?", _
              vbYesNo + vbQuestion, "Validaciones POR") = vbYes Then
        If Not Validar_Contabilidad_POR() Then
            MsgBox "Corrige los errores marcados y vuelve a intentarlo.", _
                   vbCritical, "Proceso Cancelado"
            Exit Sub
        End If
        MsgBox "Validación correcta. Continuando...", vbInformation, "OK"
    End If

    ' 4 ── Datos del usuario ───────────────────────────────────────────────
    cliente = PedirCliente(PAC_POR)
    If cliente = "" Then Exit Sub
    release = PedirRelease()
    If release = "" Then Exit Sub

    ' 5 ── Diccionario de mayoría ──────────────────────────────────────────
    Set comboMayoria = ConstruirDiccionarioMayoria(wsC, celdaIni)

    ' 6 ── Activar modo rendimiento (pantalla oculta desde aquí) ──────────
    PrepararCarpetasBase
    wsC.AutoFilterMode = False
    ActivarModoRendimiento

    ' 7 ── Nueva hoja LINEASVCA ────────────────────────────────────────────
    EliminarHojaSiExiste HOJA_LINEAS, ThisWorkbook
    Set wsT  = ThisWorkbook.Worksheets.Add
    wsT.Name = HOJA_LINEAS
    CrearCabecerasLINEASVCA wsT
    filaNueva = 2
    contador  = 5

    ' 8 ── Bucle de generación ─────────────────────────────────────────────
    For i = celdaIni.Row To ultFila
        clave = Trim(wsC.Cells(i, colBase).Value)
        If Trim(CStr(clave)) <> "" Then
            If comboMayoria.Exists(clave) Then
                partes   = Split(Trim(comboMayoria(clave)), "|")
                valDebe  = partes(0)
                valHaber = partes(1)
                If (valDebe <> "" Or valHaber <> "") And _
                   InStr(valDebe, " ") = 0 And InStr(valHaber, " ") = 0 Then
                    EscribirLineaVCA wsT, filaNueva, TIPO_POR, cliente, PAC_POR, _
                                     release, contador, CStr(clave), valDebe, valHaber
                    contador  = contador + 5
                    filaNueva = filaNueva + 1
                End If
                comboMayoria.Remove clave
            End If
        End If
    Next i

    ' 9 ── Formato y reordenar pestañas (pantalla AÚN oculta → sin parpadeo)
    AplicarFormatoTablaLINEASVCA wsT
    GestorPestanas_ReordenarVCA HOJA_POR, HOJA_DATOS_POR, RGB(200, 255, 206), HOJA_ESP

    ' 10 ── Restaurar DESPUÉS de mover pestañas ───────────────────────────
    RestaurarModoNormal

    ' 11 ── Guardar ────────────────────────────────────────────────────────
    ruta = ObtenerRutaVersionada(RUTA_BASE, "VCA_" & cliente & "_" & release & ".xls")
    GuardarHojaComoXLS wsT, ruta, filaNueva - 2
End Sub


'==========================================================================
' Validar_Contabilidad_POR
'==========================================================================
Public Function Validar_Contabilidad_POR() As Boolean

    Dim ws          As Worksheet
    Dim celdaIni    As Range
    Dim celdaTC     As Range
    Dim colBase     As Long, colDebe As Long, colHaber As Long, colTC As Long
    Dim i           As Long, ultFila As Long
    Dim enlace      As String, debe As String, haber As String, tipoConc As String
    Dim comboMay    As Object, comboAct As String, clave As Variant
    Dim especiales  As Variant, v As Variant
    Dim hayError    As Boolean, nErrores As Long, lista As String

    Set ws = ObtenerHojaSegura(HOJA_DATOS_POR, ThisWorkbook)
    If ws Is Nothing Then
        Validar_Contabilidad_POR = False
        Exit Function
    End If

    Set celdaIni = POR_Celda_Inicial()
    If celdaIni Is Nothing Then
        Validar_Contabilidad_POR = False
        Exit Function
    End If

    colBase  = celdaIni.Column
    colDebe  = colBase + 1
    colHaber = colBase + 2
    ultFila  = ws.Cells(ws.Rows.Count, colBase).End(xlUp).Row

    ' Localizar TIPO CONCEPTO (dinámica)
    Set celdaTC = BuscarCeldaPorTexto(ws, "TIPO CONCEPTO", xlPart)
    If celdaTC Is Nothing Then
        MsgBox "No se encontró 'TIPO CONCEPTO' en '" & HOJA_DATOS_POR & "'.", _
               vbCritical, "Cabecera No Encontrada"
        Validar_Contabilidad_POR = False
        Exit Function
    End If
    colTC = ObtenerColumnaReal(celdaTC)

    POR_LimpiarComentariosValidacion

    Set comboMay = ConstruirDiccionarioMayoria(ws, celdaIni)
    especiales   = Array("G- Gestión")
    hayError     = False
    nErrores     = 0
    lista        = ""

    For i = celdaIni.Row To ultFila
        enlace   = Trim(ws.Cells(i, colBase).Text)
        debe     = Trim(ws.Cells(i, colDebe).Text)
        haber    = Trim(ws.Cells(i, colHaber).Text)
        tipoConc = Trim(ws.Cells(i, colTC).Text)

        ' R1: Sin espacios
        If debe <> "" Or haber <> "" Then
            If InStr(debe, " ") <> 0 Then
                POR_MarcarError ws.Cells(i, colDebe), RGB(255, 189, 180), _
                    "ERROR – No puede contener espacios. Se descartará"
            End If
            If InStr(haber, " ") <> 0 Then
                POR_MarcarError ws.Cells(i, colHaber), RGB(255, 189, 180), _
                    "ERROR – No puede contener espacios. Se descartará"
            End If
            If InStr(debe, " ") <> 0 Or InStr(haber, " ") <> 0 Then
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – Espacios en Debe/Haber." & vbCrLf
            End If
        End If

        ' R2: Consistencia vs mayoritaria
        clave    = Trim(ws.Cells(i, colBase).Value)
        comboAct = Trim(ws.Cells(i, colDebe).Value) & "|" & _
                   Trim(ws.Cells(i, colHaber).Value)
        If Trim(CStr(clave)) <> "" And comboMay.Exists(clave) Then
            If Trim(comboAct) <> "|" And comboAct <> comboMay(clave) Then
                ws.Cells(i, colBase).Interior.Color  = vbRed
                ws.Cells(i, colDebe).Interior.Color  = vbRed
                ws.Cells(i, colHaber).Interior.Color = vbRed
                AddOrAppendComment ws.Cells(i, colBase), _
                    "AVISO – Enlace: " & clave & vbCrLf & _
                    "Esta fila : " & Replace(comboAct, "|", " / ") & vbCrLf & _
                    "Mayoritaria: " & Replace(comboMay(clave), "|", " / ")
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – Debe/Haber distinto al mayoritario." & vbCrLf
            End If
        End If

        ' R3: Gestión → Debe y Haber obligatorios
        Dim esGestion As Boolean: esGestion = False
        For Each v In especiales
            If UCase(tipoConc) = UCase(CStr(v)) Then esGestion = True: Exit For
        Next v
        If esGestion Then
            If (debe = "" And haber <> "") Or (debe <> "" And haber = "") Then
                ws.Cells(i, colBase).Interior.Color  = vbYellow
                ws.Cells(i, colDebe).Interior.Color  = vbYellow
                ws.Cells(i, colHaber).Interior.Color = vbYellow
                AddOrAppendComment ws.Cells(i, colBase), _
                    "AVISO – GESTIÓN: deben estar informados Debe y Haber"
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – GESTIÓN: faltan Debe y/o Haber." & vbCrLf
            End If
        End If

        ' R4: Enlace > 500
        If IsNumeric(enlace) Then
            If CLng(enlace) > 500 Then
                ws.Cells(i, colBase).Interior.Color  = vbMagenta
                ws.Cells(i, colDebe).Interior.Color  = vbMagenta
                ws.Cells(i, colHaber).Interior.Color = vbMagenta
                AddOrAppendComment ws.Cells(i, colBase), _
                    "ERROR – Enlace " & enlace & " supera el máximo (500)"
                hayError = True: nErrores = nErrores + 1
                lista = lista & "· Fila " & i & " – Enlace " & enlace & " > 500." & vbCrLf
            End If
        End If
    Next i

    If hayError Then
        MsgBox "Se detectaron " & nErrores & " error(es) en '" & HOJA_DATOS_POR & "':" & _
               vbCrLf & vbCrLf & lista & vbCrLf & _
               "Revisa los comentarios y colores en la hoja.", _
               vbCritical, "Errores de Validación POR"
        Validar_Contabilidad_POR = False
    Else
        Validar_Contabilidad_POR = True
    End If
End Function
