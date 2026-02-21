Attribute VB_Name = "Comun_ExportarFoto"
Option Explicit

'==========================================================================
' MÓDULO : Comun_ExportarFoto
' FUNCIÓN: Exporta todo el código VBA del proyecto a la hoja FOTO_VBA
'          en bloques OCR-friendly para fotografiar y reconstruir con IA.
' DEPENDE: Comun_Constantes
'==========================================================================

Private Const TAM_BLOQUE As Long   = 120
Private Const SEP_MOD    As String = "[[MOD:"
Private Const SEP_NL     As String = "~NL~"
Private Const SEP_FIN    As String = "]]"

Public Sub ExportarVBA_A_FOTO()

    Dim vbc          As Object
    Dim codTotal     As String
    Dim linea        As String
    Dim i            As Long, n As Long, r As Long
    Dim ws           As Worksheet
    Dim totalBloques As Long
    Dim bloquesPorCol As Long
    Dim colBase      As Long

    codTotal = ""

    ' Verificar acceso al VBProject
    ' Requiere: Opciones → Centro de Confianza → Configuración de macros →
    '           "Confiar en el acceso al modelo de objetos de proyectos de VBA"
    On Error Resume Next
    Dim testAcceso As Object
    Set testAcceso = ThisWorkbook.VBProject
    On Error GoTo 0
    If testAcceso Is Nothing Then
        MsgBox "No se puede acceder al código VBA del proyecto." & vbCrLf & vbCrLf & _
               "Para activarlo:" & vbCrLf & _
               "1. Archivo → Opciones → Centro de Confianza" & vbCrLf & _
               "2. Configuración del Centro de Confianza" & vbCrLf & _
               "3. Configuración de macros → activar:" & vbCrLf & _
               "   'Confiar en el acceso al modelo de objetos de proyectos de VBA'" & vbCrLf & _
               "4. Acepta y vuelve a ejecutar.", _
               vbCritical, "Acceso Bloqueado"
        Exit Sub
    End If

    ' Extraer código de todos los módulos
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        codTotal = codTotal & SEP_MOD & vbc.Name & SEP_FIN
        On Error Resume Next
        With vbc.CodeModule
            If .CountOfLines > 0 Then
                For i = 1 To .CountOfLines
                    linea = Trim$(.Lines(i, 1))
                    If linea <> "" And Left$(linea, 1) <> "'" Then
                        Do While InStr(linea, "  ") > 0
                            linea = Replace(linea, "  ", " ")
                        Loop
                        codTotal = codTotal & linea & SEP_NL
                    End If
                Next i
            End If
        End With
        On Error GoTo 0
    Next vbc

    If Len(codTotal) = 0 Then
        MsgBox "No se extrajo código." & vbCrLf & _
               "Verifica acceso al VBProject en Centro de Confianza.", _
               vbCritical, "Sin Código"
        Exit Sub
    End If

    ' Crear / recrear hoja FOTO_VBA
    EliminarHojaSiExiste HOJA_FOTO, ThisWorkbook
    Set ws   = ThisWorkbook.Sheets.Add
    ws.Name  = HOJA_FOTO

    ' Volcar en dos columnas
    totalBloques  = Application.WorksheetFunction.RoundUp(Len(codTotal) / TAM_BLOQUE, 0)
    bloquesPorCol = Application.WorksheetFunction.RoundUp(totalBloques / 2, 0)
    n = 1

    For i = 1 To Len(codTotal) Step TAM_BLOQUE
        colBase = IIf(n <= bloquesPorCol, 0, 3)
        r       = ((n - 1) Mod bloquesPorCol) + 2
        ws.Cells(r, colBase + 1).Value = n
        ws.Cells(r, colBase + 2).Value = "'" & Mid$(codTotal, i, TAM_BLOQUE)
        n = n + 1
    Next i

    ' Formato
    With ws
        .Cells.Font.Name         = "Consolas"
        .Cells.Font.Size         = 11
        .Cells.VerticalAlignment = xlTop
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 99
        .Columns("C").ColumnWidth = 3
        .Columns("D").ColumnWidth = 6
        .Columns("E").ColumnWidth = 98
        .Rows(1).Font.Bold        = True
        .Cells(1, 1).Value = "ID"
        .Cells(1, 2).Value = "DATA"
        .Cells(1, 4).Value = "ID"
        .Cells(1, 5).Value = "DATA"
        .Range("A2").Select
    End With

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom             = 100
    ActiveWindow.FreezePanes      = True
    Application.DisplayFullScreen = True

    MsgBox "¡Listo!" & vbCrLf & vbCrLf & _
           "Bloques: " & (n - 1) & vbCrLf & _
           "Fotografía las columnas A–B y D–E.", _
           vbInformation, "Exportación Completada"
End Sub
