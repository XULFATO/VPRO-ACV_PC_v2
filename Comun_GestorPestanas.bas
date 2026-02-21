Attribute VB_Name = "Comun_GestorPestanas"
Option Explicit

'==========================================================================
' MÓDULO : Comun_GestorPestanas
' FUNCIÓN: Gestión centralizada de pestañas para todos los procesos.
'          - Detecta ejecuciones anteriores y pregunta si borrarlas
'          - Nunca toca HOME ni hojas protegidas
'          - Versiona hojas antiguas como 01_OLD_... si el usuario dice NO
'          - Reordena pestañas al finalizar un proceso
' DEPENDE: Comun_Constantes
'==========================================================================


'==========================================================================
' GestorPestanas_VerificarAntesDeEjecutar
' Llama a esto al inicio de cada proceso.
' Detecta si ya existen hojas del proceso y pregunta al usuario.
'
' Parámetros:
'   nombreProceso → texto para el mensaje ("VCA España", "VCA Portugal"...)
'   hojasDelProceso → array con los nombres de hojas que usa ese proceso
'
' Devuelve:
'   True  → puede continuar (borró o el usuario eligió continuar)
'   False → usuario canceló, abortar proceso
'==========================================================================
Public Function GestorPestanas_VerificarAntesDeEjecutar( _
                    ByVal nombreProceso   As String, _
                    ByVal hojasDelProceso As Variant) As Boolean

    Dim hojasExistentes() As String
    Dim nExistentes       As Long
    Dim nombre            As Variant
    Dim respuesta         As VbMsgBoxResult

    ' Detectar cuáles de las hojas del proceso ya existen
    nExistentes = 0
    ReDim hojasExistentes(0)

    For Each nombre In hojasDelProceso
        If HojaExiste(CStr(nombre), ThisWorkbook) Then
            If Not EsHojaProtegida(CStr(nombre)) Then
                ReDim Preserve hojasExistentes(nExistentes)
                hojasExistentes(nExistentes) = CStr(nombre)
                nExistentes = nExistentes + 1
            End If
        End If
    Next nombre

    ' Si no hay hojas anteriores → continuar directamente
    If nExistentes = 0 Then
        GestorPestanas_VerificarAntesDeEjecutar = True
        Exit Function
    End If

    ' Preguntar al usuario
    respuesta = MsgBox("Ya existe una ejecución de " & nombreProceso & "." & vbCrLf & vbCrLf & _
                       "¿Deseas borrarla antes de continuar?", _
                       vbYesNoCancel + vbQuestion, "Ejecución Anterior Detectada")

    Select Case respuesta
        Case vbYes
            ' Borrar todas las hojas del proceso (excepto protegidas)
            GestorPestanas_BorrarHojasProceso hojasDelProceso
            GestorPestanas_VerificarAntesDeEjecutar = True

        Case vbNo
            ' Versionar las existentes como OLD y continuar
            GestorPestanas_VersionarHojas hojasDelProceso
            GestorPestanas_VerificarAntesDeEjecutar = True

        Case vbCancel
            GestorPestanas_VerificarAntesDeEjecutar = False
    End Select
End Function


'==========================================================================
' GestorPestanas_BorrarHojasProceso
' Borra las hojas del proceso. NUNCA toca hojas protegidas.
'==========================================================================
Public Sub GestorPestanas_BorrarHojasProceso(ByVal hojasDelProceso As Variant)
    Dim nombre As Variant

    Application.DisplayAlerts = False
    For Each nombre In hojasDelProceso
        If HojaExiste(CStr(nombre), ThisWorkbook) Then
            If Not EsHojaProtegida(CStr(nombre)) Then
                ThisWorkbook.Worksheets(CStr(nombre)).Delete
            End If
        End If
    Next nombre

    ' Borrar también las OLD de este proceso
    Dim i       As Integer
    Dim wsNombre As String
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        wsNombre = ThisWorkbook.Worksheets(i).Name
        If Left(wsNombre, 4) Like "##_O" Then   ' patrón 01_OLD_...
            If Not EsHojaProtegida(wsNombre) Then
                ThisWorkbook.Worksheets(i).Delete
            End If
        End If
    Next i

    Application.DisplayAlerts = True
End Sub


'==========================================================================
' GestorPestanas_VersionarHojas
' Renombra las hojas existentes como 01_OLD_nombre, 02_OLD_nombre...
' y las mueve al final. NUNCA toca hojas protegidas.
'==========================================================================
Public Sub GestorPestanas_VersionarHojas(ByVal hojasDelProceso As Variant)
    Dim nombre      As Variant
    Dim nombreNuevo As String
    Dim contador    As Integer

    For Each nombre In hojasDelProceso
        If HojaExiste(CStr(nombre), ThisWorkbook) Then
            If Not EsHojaProtegida(CStr(nombre)) Then
                ' Buscar número de versión disponible
                For contador = 1 To 99
                    nombreNuevo = Left(Format(contador, "00") & "_OLD_" & CStr(nombre), 31)
                    If Not HojaExiste(nombreNuevo, ThisWorkbook) Then
                        Application.DisplayAlerts = False
                        ThisWorkbook.Worksheets(CStr(nombre)).Name = nombreNuevo
                        Application.DisplayAlerts = True
                        ThisWorkbook.Worksheets(nombreNuevo).Move _
                            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
                        Exit For
                    End If
                Next contador
            End If
        End If
    Next nombre
End Sub


'==========================================================================
' GestorPestanas_ReordenarVCA
' Reordena las pestañas al finalizar un proceso VCA.
' HOME siempre en posición 1.
' La pantalla debe estar OCULTA cuando se llama a esto.
'
' Parámetros:
'   hojaPrincipal → VCA_ESP o VCA_POR (irá en posición 2)
'   hojaDatos     → Contabilidad_Cuentas o Analisis Conceitos (pos 3)
'   colorLineas   → color de pestaña para LINEASVCA
'   hojaAlFinal   → la otra hoja de país que va al final (VCA_POR o VCA_ESP)
'==========================================================================
Public Sub GestorPestanas_ReordenarVCA(ByVal hojaPrincipal As String, _
                                        ByVal hojaDatos     As String, _
                                        ByVal colorLineas   As Long, _
                                        ByVal hojaAlFinal   As String)
    ' Pantalla debe estar oculta — la gestiona el proceso llamante
    With ThisWorkbook

        ' Mover LINEASVCA antes de todo (luego se reubicará)
        If HojaExiste(HOJA_LINEAS, ThisWorkbook) Then
            .Worksheets(HOJA_LINEAS).Move Before:=.Worksheets(1)
            .Worksheets(HOJA_LINEAS).Tab.Color = colorLineas
        End If

        ' Mover hoja de datos
        If HojaExiste(hojaDatos, ThisWorkbook) Then
            .Worksheets(hojaDatos).Move Before:=.Worksheets(1)
        End If

        ' Mover hoja principal del proceso
        If HojaExiste(hojaPrincipal, ThisWorkbook) Then
            .Worksheets(hojaPrincipal).Move Before:=.Worksheets(1)
        End If

        ' HOME siempre primera
        If HojaExiste(HOJA_HOME, ThisWorkbook) Then
            .Worksheets(HOJA_HOME).Move Before:=.Worksheets(1)
        End If

        ' La otra hoja de país va al final (antes de las OLD)
        If HojaExiste(hojaAlFinal, ThisWorkbook) Then
            ' Buscar primera OLD para insertar antes
            Dim i        As Integer
            Dim posOLD   As Integer
            posOLD = .Worksheets.Count
            For i = 1 To .Worksheets.Count
                If Left(.Worksheets(i).Name, 6) Like "##_OLD" Then
                    posOLD = i - 1
                    Exit For
                End If
            Next i
            If posOLD >= 1 And posOLD < .Worksheets.Count Then
                .Worksheets(hojaAlFinal).Move Before:=.Worksheets(posOLD + 1)
            Else
                .Worksheets(hojaAlFinal).Move After:=.Worksheets(.Worksheets.Count)
            End If
        End If

    End With
End Sub
