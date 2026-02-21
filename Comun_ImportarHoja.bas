Attribute VB_Name = "Comun_ImportarHoja"
Option Explicit

'==========================================================================
' MÓDULO : Comun_ImportarHoja
' FUNCIÓN: Importa una hoja de datos desde otro libro abierto.
'          La hoja existente se versiona como OLD antes de importar.
'          Detecta el país según la hoja activa (VCA_ESP o VCA_POR).
' DEPENDE: Comun_Constantes, Comun_GestorPestanas
'==========================================================================

Public Sub ImportarHojaDesdeOtroLibro()

    Dim wbActual       As Workbook
    Dim wb             As Workbook
    Dim libros         As Collection
    Dim nombres        As Collection
    Dim hojaOrigen     As Worksheet
    Dim nombreHojaDat  As String
    Dim seleccion      As Variant
    Dim prompt         As String
    Dim i              As Integer

    Set wbActual = ThisWorkbook
    Set libros   = New Collection
    Set nombres  = New Collection

    ' Determinar qué hoja importar según hoja activa
    Select Case wbActual.ActiveSheet.Name
        Case HOJA_ESP
            nombreHojaDat = HOJA_DATOS_ESP
        Case HOJA_POR
            nombreHojaDat = HOJA_DATOS_POR
        Case Else
            MsgBox "Ejecuta este proceso desde '" & HOJA_ESP & "' o '" & HOJA_POR & "'.", _
                   vbExclamation, "Hoja Incorrecta"
            Exit Sub
    End Select

    ' Recopilar libros abiertos (excluir el actual)
    For Each wb In Application.Workbooks
        If wb.Name <> wbActual.Name Then
            libros.Add wb
            nombres.Add wb.Name
        End If
    Next wb

    If libros.Count = 0 Then
        MsgBox "No hay otros libros Excel abiertos." & vbCrLf & _
               "Abre el Blueprint con '" & nombreHojaDat & "' y vuelve a intentarlo.", _
               vbExclamation, "Sin Libros Disponibles"
        Exit Sub
    End If

    ' Mostrar lista de libros disponibles
    prompt = "Selecciona el número del libro del que importar '" & nombreHojaDat & "':" & _
             vbCrLf & vbCrLf
    For i = 1 To nombres.Count
        prompt = prompt & i & ".  " & nombres(i) & vbCrLf
    Next i

    seleccion = InputBox(prompt, "Seleccionar Libro")
    If Not IsNumeric(seleccion) Then Exit Sub
    If CLng(seleccion) < 1 Or CLng(seleccion) > nombres.Count Then
        MsgBox "Número fuera de rango.", vbExclamation, "Cancelado"
        Exit Sub
    End If

    Set wb = libros(CInt(seleccion))

    ' Verificar que la hoja existe en el libro elegido
    If Not HojaExiste(nombreHojaDat, wb) Then
        MsgBox "El libro '" & wb.Name & "' no contiene la hoja '" & nombreHojaDat & "'.", _
               vbExclamation, "Hoja No Encontrada"
        Exit Sub
    End If
    Set hojaOrigen = wb.Sheets(nombreHojaDat)

    ' Versionar la hoja existente como OLD si la hay
    If HojaExiste(nombreHojaDat, wbActual) Then
        Dim hojasArr(0) As String
        hojasArr(0) = nombreHojaDat
        GestorPestanas_VersionarHojas hojasArr
    End If

    ' Copiar hoja al libro actual
    hojaOrigen.Copy After:=wbActual.Sheets(wbActual.Sheets.Count)

    MsgBox "La hoja '" & nombreHojaDat & "' se copió correctamente desde '" & wb.Name & "'.", _
           vbInformation, "Importación Completada"
End Sub
