Attribute VB_Name = "Comun_Limpieza"
Option Explicit

'==========================================================================
' MÓDULO : Comun_Limpieza
' FUNCIÓN: Limpieza del libro. Borra hojas generadas por procesos.
'          NUNCA borra HOME ni hojas protegidas (definidas en Comun_Constantes).
' DEPENDE: Comun_Constantes, Comun_GestorPestanas
'==========================================================================


'==========================================================================
' LimpiarHojasProceso
' Llamada desde el menú o desde un proceso concreto.
' Borra todas las hojas NO protegidas del libro.
' Pregunta confirmación antes de actuar.
'==========================================================================
Public Sub LimpiarHojasProceso()
    Dim i        As Integer
    Dim ws       As Worksheet
    Dim lista    As String
    Dim nBorrar  As Integer

    ' Construir lista de hojas que se borrarían
    lista   = ""
    nBorrar = 0
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        If Not EsHojaProtegida(ws.Name) Then
            lista   = lista & "  · " & ws.Name & vbCrLf
            nBorrar = nBorrar + 1
        End If
    Next i

    If nBorrar = 0 Then
        MsgBox "No hay hojas de proceso que limpiar." & vbCrLf & _
               "El libro solo contiene hojas protegidas.", _
               vbInformation, "Nada que Limpiar"
        Exit Sub
    End If

    ' Confirmar con el usuario
    If MsgBox("Se borrarán las siguientes hojas:" & vbCrLf & vbCrLf & lista & vbCrLf & _
              "Las hojas protegidas (HOME, VCA_ESP, VCA_POR) NO se tocarán." & vbCrLf & vbCrLf & _
              "¿Continuar?", vbYesNo + vbExclamation, "Confirmar Limpieza") = vbNo Then
        Exit Sub
    End If

    ' Borrar
    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Not EsHojaProtegida(ws.Name) Then
            ws.Delete
        End If
    Next i
    Application.DisplayAlerts  = True
    Application.ScreenUpdating = True

    MsgBox "Limpieza completada." & vbCrLf & _
           "Se han eliminado " & nBorrar & " hoja(s).", _
           vbInformation, "Limpieza Completada"
End Sub
