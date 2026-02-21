Attribute VB_Name = "Comun_Constantes"
Option Explicit

'==========================================================================
' MÓDULO : Comun_Constantes
' FUNCIÓN: Constantes globales y utilidades compartidas por todos los módulos.
' DEPENDE: Nada (es la base de todo)
'==========================================================================

' ── Hoja protegida ────────────────────────────────────────────────────────
Public Const HOJA_HOME      As String = "HOME"

' ── Nombres de hojas de proceso ───────────────────────────────────────────
Public Const HOJA_ESP       As String = "VCA_ESP"
Public Const HOJA_POR       As String = "VCA_POR"
Public Const HOJA_DATOS_ESP As String = "Contabilidad_Cuentas"
Public Const HOJA_DATOS_POR As String = "Analisis Conceitos"
Public Const HOJA_LINEAS    As String = "LINEASVCA"
Public Const HOJA_FOTO      As String = "FOTO_VBA"

' ── Parámetros VCA ────────────────────────────────────────────────────────
Public Const RUTA_BASE      As String = "C:\Clientes\VCA\Generados"
Public Const PREFIJO_VAL    As String = "[VALIDACION]"
Public Const MAX_VERSIONES  As Long   = 999
Public Const TIPO_ESP       As String = "18"
Public Const TIPO_POR       As String = "20"
Public Const PAC_ESP        As String = "991"
Public Const PAC_POR        As String = "993"

' ── Hojas que NUNCA se borran (de momento) ────────────────────────────────
' En el futuro VCA_ESP y VCA_POR podrán borrarse desde una opción del menú HOME
Public Function EsHojaProtegida(ByVal nombre As String) As Boolean
    Select Case UCase(Trim(nombre))
        Case UCase(HOJA_HOME), UCase(HOJA_ESP), UCase(HOJA_POR)
            EsHojaProtegida = True
        Case Else
            EsHojaProtegida = False
    End Select
End Function


'==========================================================================
' HojaExiste
'==========================================================================
Public Function HojaExiste(ByVal nombre As String, _
                            ByVal libro  As Workbook) As Boolean
    Dim h As Worksheet
    For Each h In libro.Sheets
        If UCase(Trim(h.Name)) = UCase(Trim(nombre)) Then
            HojaExiste = True
            Exit Function
        End If
    Next h
    HojaExiste = False
End Function


'==========================================================================
' ObtenerHojaSegura
' Devuelve la hoja o Nothing con mensaje de error.
'==========================================================================
Public Function ObtenerHojaSegura(ByVal nombre As String, _
                                   ByVal libro  As Workbook) As Worksheet
    Set ObtenerHojaSegura = Nothing
    If HojaExiste(nombre, libro) Then
        Set ObtenerHojaSegura = libro.Worksheets(nombre)
    Else
        MsgBox "ERROR: No se encuentra la pestaña '" & nombre & "'." & vbCrLf & _
               "Comprueba que existe y que el nombre es exacto.", _
               vbCritical, "Hoja No Encontrada"
    End If
End Function


'==========================================================================
' EliminarHojaSiExiste
'==========================================================================
Public Sub EliminarHojaSiExiste(ByVal nombre As String, _
                                 ByVal libro  As Workbook)
    If HojaExiste(nombre, libro) Then
        Application.DisplayAlerts = False
        libro.Worksheets(nombre).Delete
        Application.DisplayAlerts = True
    End If
End Sub


'==========================================================================
' ActivarModoRendimiento / RestaurarModoNormal
' NOTA: NO llames a RestaurarModoNormal antes de mover pestañas.
'       Hazlo DESPUÉS para evitar parpadeo.
'==========================================================================
Public Sub ActivarModoRendimiento()
    With Application
        .ScreenUpdating = False
        .Calculation    = xlCalculationManual
        .DisplayAlerts  = False
        .EnableEvents   = False
    End With
End Sub

Public Sub RestaurarModoNormal()
    With Application
        .ScreenUpdating = True
        .Calculation    = xlCalculationAutomatic
        .DisplayAlerts  = True
        .EnableEvents   = True
    End With
End Sub


'==========================================================================
' EnsureDirectoryExists
'==========================================================================
Public Sub EnsureDirectoryExists(ByVal ruta As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(ruta) Then
        On Error Resume Next
        fso.CreateFolder ruta
        On Error GoTo 0
        If Not fso.FolderExists(ruta) Then
            Set fso = Nothing
            MsgBox "ERROR CRÍTICO: No se pudo crear la carpeta:" & vbCrLf & ruta & vbCrLf & _
                   "Verifica permisos de escritura.", vbCritical, "Error de Carpeta"
            Exit Sub
        End If
    End If
    Set fso = Nothing
End Sub


'==========================================================================
' PrepararCarpetasBase
'==========================================================================
Public Sub PrepararCarpetasBase()
    EnsureDirectoryExists "C:\Clientes"
    EnsureDirectoryExists "C:\Clientes\VCA"
    EnsureDirectoryExists RUTA_BASE
End Sub


'==========================================================================
' ObtenerRutaVersionada
' Añade _v001, _v002... si el archivo ya existe.
'==========================================================================
Public Function ObtenerRutaVersionada(ByVal carpeta    As String, _
                                       ByVal nombreBase As String) As String
    Dim fso  As Object
    Dim base As String, ext As String, ruta As String
    Dim n    As Long

    Set fso  = CreateObject("Scripting.FileSystemObject")
    base     = fso.GetBaseName(nombreBase)
    ext      = fso.GetExtensionName(nombreBase)
    ruta     = fso.BuildPath(carpeta, nombreBase)
    n        = 0

    Do While fso.FileExists(ruta)
        n = n + 1
        If n > MAX_VERSIONES Then
            Set fso = Nothing
            MsgBox "ERROR: Límite de " & MAX_VERSIONES & " versiones alcanzado para '" & _
                   nombreBase & "'." & vbCrLf & "Limpia la carpeta de destino.", _
                   vbCritical, "Límite de Versiones"
            ObtenerRutaVersionada = ""
            Exit Function
        End If
        ruta = fso.BuildPath(carpeta, base & "_v" & Format(n, "000") & "." & ext)
    Loop

    ObtenerRutaVersionada = ruta
    Set fso = Nothing
End Function


'==========================================================================
' PedirCliente
'==========================================================================
Public Function PedirCliente(ByVal pac As String) As String
    Dim entrada As String, tres As String
    Do
        entrada = UCase(Trim(InputBox("Introduce el código de CLIENTE" & vbCrLf & _
                                      "Formato: 3 dígitos o " & pac & "xxx", "CLIENTE")))
        If entrada = vbNullString Then PedirCliente = "": Exit Function
        Select Case Len(entrada)
            Case 6
                tres = Left(entrada, 3)
                If tres <> pac Then
                    MsgBox "Los 3 primeros caracteres deben ser '" & pac & "'.", _
                           vbExclamation, "Código Inválido"
                Else
                    PedirCliente = Right(entrada, 3): Exit Function
                End If
            Case 3
                PedirCliente = entrada: Exit Function
            Case Else
                MsgBox "Introduce 3 dígitos o el código completo de 6 (" & pac & "xxx).", _
                       vbExclamation, "Longitud Inválida"
        End Select
    Loop
End Function


'==========================================================================
' PedirRelease
'==========================================================================
Public Function PedirRelease() As String
    Dim entrada As String
    Do
        entrada = Trim(InputBox("Introduce el número de RELEASE:", "RELEASE"))
        If entrada = vbNullString Then PedirRelease = "": Exit Function
        If IsNumeric(entrada) Then
            PedirRelease = entrada: Exit Function
        Else
            MsgBox "El Release debe ser numérico. Has introducido: '" & entrada & "'.", _
                   vbExclamation, "Valor Inválido"
        End If
    Loop
End Function


'==========================================================================
' AddOrAppendComment
' Añade o amplía comentario de validación. No duplica el mismo mensaje.
'==========================================================================
Public Sub AddOrAppendComment(ByVal celda  As Range, _
                               ByVal texto  As String)
    Dim msg As String
    msg = PREFIJO_VAL & " " & texto
    If celda.Comment Is Nothing Then
        celda.AddComment msg
        celda.Comment.Shape.TextFrame.AutoSize = True
    Else
        If InStr(1, celda.Comment.Text, msg, vbTextCompare) = 0 Then
            celda.Comment.Text celda.Comment.Text & vbCrLf & msg
            celda.Comment.Shape.TextFrame.AutoSize = True
        End If
    End If
End Sub


'==========================================================================
' EscribirLineaVCA  *** ÚNICA DEFINICIÓN EN TODO EL PROYECTO ***
'==========================================================================
Public Sub EscribirLineaVCA(ByVal wsT      As Worksheet, _
                              ByVal fila     As Long, _
                              ByVal tipo     As String, _
                              ByVal cliente  As String, _
                              ByVal pac      As String, _
                              ByVal release  As String, _
                              ByVal contador As Long, _
                              ByVal enlace   As String, _
                              ByVal valDebe  As String, _
                              ByVal valHaber As String)
    wsT.Cells(fila, "A") = tipo
    wsT.Cells(fila, "B") = cliente
    wsT.Cells(fila, "C") = pac & cliente
    wsT.Cells(fila, "D") = release
    wsT.Cells(fila, "E") = "V"
    wsT.Cells(fila, "F") = "VCA"
    wsT.Cells(fila, "G") = contador
    wsT.Cells(fila, "H") = "1"
    wsT.Cells(fila, "I")  = enlace
    wsT.Cells(fila, "J")  = "01"
    wsT.Cells(fila, "K")  = "99"
    wsT.Cells(fila, "M")  = "999"
    wsT.Cells(fila, "Q")  = "999"
    wsT.Cells(fila, "S")  = "999"
    wsT.Cells(fila, "U")  = "9"
    wsT.Cells(fila, "AA") = "9"
    If valDebe <> "" Then wsT.Cells(fila, "AB") = valDebe
    If valHaber <> "" Then wsT.Cells(fila, "AG") = valHaber
End Sub


'==========================================================================
' CrearCabecerasLINEASVCA
'==========================================================================
Public Sub CrearCabecerasLINEASVCA(ByVal wsT As Worksheet)
    wsT.Range("A1:AK1").Value = Array( _
        "Tipo", "Cliente", "Pac", "Release", _
        "Id", "Cod.Tabla", "Lineas", "Tip Lin", _
        "COD.ENL", "EM.DE", "EM.HA.", "CTR.DE", _
        "CTR.HA", "T.E.D", "T.E.H", "CAT.DE", _
        "CAT.HA", "T.C.D", "T.C.H", "D.I.D", _
        "D.I.H", "T.R.D", "T.R.H", "CENT.COST.DESDE", _
        "CENT.COST.HASTA", "AR.LI.D", "AR.LI.HA", "NUM.CUENTA", _
        "VALOR.ESPEC.", "NAT.", "CO.OP", "RESERVADO", "CONTR.NUM.CTA", _
        "CONTR.VAL.ESP.", "CON.NAT", "CON.CO.OP", "RESERVADO")
    Dim col As Variant
    For Each col In Array("A", "I", "J", "K", "M", "Q", "S", "U", "AA")
        wsT.Columns(CStr(col)).NumberFormat = "@"
    Next col
End Sub


'==========================================================================
' AplicarFormatoTablaLINEASVCA
'==========================================================================
Public Sub AplicarFormatoTablaLINEASVCA(ByVal wsT As Worksheet)
    Dim rng As Range
    Dim lo  As ListObject
    Set rng = wsT.Range("A1").CurrentRegion
    Set lo  = wsT.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.TableStyle = "TableStyleMedium2"
    wsT.Columns("A:I").AutoFit
    wsT.Columns("J:AA").ColumnWidth  = 1
    wsT.Columns("AC:AF").ColumnWidth = 1
End Sub


'==========================================================================
' GuardarHojaComoXLS  *** ÚNICA DEFINICIÓN EN TODO EL PROYECTO ***
'==========================================================================
Public Sub GuardarHojaComoXLS(ByVal ws              As Worksheet, _
                                ByVal rutaArchivo     As String, _
                                ByVal lineasGeneradas As Long)
    If rutaArchivo = "" Then Exit Sub

    Dim nuevoLibro As Workbook
    Dim numErr     As Long
    Dim descErr    As String

    ws.Copy
    Set nuevoLibro = ActiveWorkbook

    Application.DisplayAlerts = False
    On Error Resume Next
    nuevoLibro.SaveAs Filename:=rutaArchivo, FileFormat:=xlExcel8
    numErr  = Err.Number
    descErr = Err.Description
    On Error GoTo 0
    Application.DisplayAlerts = True

    If numErr <> 0 Then
        nuevoLibro.Close SaveChanges:=False
        MsgBox "ERROR al guardar:" & vbCrLf & rutaArchivo & vbCrLf & descErr, _
               vbCritical, "Error al Guardar"
        Exit Sub
    End If

    nuevoLibro.Close SaveChanges:=False

    MsgBox "¡Proceso completado!" & vbCrLf & vbCrLf & _
           "Líneas generadas : " & lineasGeneradas & vbCrLf & _
           "Archivo         : " & vbCrLf & rutaArchivo, _
           vbInformation, "Proceso Completado"

    If MsgBox("¿Deseas abrir el archivo generado?", _
              vbYesNo + vbQuestion, "Abrir Archivo") = vbYes Then
        CreateObject("WScript.Shell").Run Chr(34) & rutaArchivo & Chr(34), 1, False
    End If
End Sub


'==========================================================================
' ConstruirDiccionarioMayoria  (usado solo por VCA_POR)
'==========================================================================
Public Function ConstruirDiccionarioMayoria(ByVal ws          As Worksheet, _
                                              ByVal celdaInicio As Range) As Object
    ' Scripting.Dictionary requiere Windows — no funciona en Mac
    #If Mac Then
        MsgBox "Esta función no está disponible en Office para Mac." & vbCrLf & _
               "Se requiere Windows para usar Scripting.Dictionary.", _
               vbCritical, "Sistema No Compatible"
        Set ConstruirDiccionarioMayoria = Nothing
        Exit Function
    #End If

    Dim dictConteo  As Object
    Dim dictResult  As Object
    Dim tempDict    As Object
    Dim colBase     As Long
    Dim ultimaFila  As Long
    Dim fila        As Long
    Dim clave       As Variant
    Dim combo       As Variant
    Dim comboActual As String
    Dim maxN        As Long

    Set dictConteo = CreateObject("Scripting.Dictionary")
    Set dictResult = CreateObject("Scripting.Dictionary")
    colBase    = celdaInicio.Column
    ultimaFila = ws.Cells(ws.Rows.Count, colBase).End(xlUp).Row

    For fila = celdaInicio.Row To ultimaFila
        clave       = Trim(ws.Cells(fila, colBase).Value)
        comboActual = Trim(ws.Cells(fila, colBase + 1).Value) & "|" & _
                      Trim(ws.Cells(fila, colBase + 2).Value)
        If Trim(CStr(clave)) <> "" And Trim(comboActual) <> "|" Then
            If Not dictConteo.Exists(clave) Then
                Set tempDict = CreateObject("Scripting.Dictionary")
                dictConteo.Add clave, tempDict
            End If
            If dictConteo(clave).Exists(comboActual) Then
                dictConteo(clave)(comboActual) = dictConteo(clave)(comboActual) + 1
            Else
                dictConteo(clave).Add comboActual, 1
            End If
        End If
    Next fila

    For Each clave In dictConteo.Keys
        maxN = 0
        For Each combo In dictConteo(clave).Keys
            If dictConteo(clave)(combo) > maxN Then
                maxN = dictConteo(clave)(combo)
                dictResult(clave) = combo
            End If
        Next combo
    Next clave

    Set ConstruirDiccionarioMayoria = dictResult
End Function
