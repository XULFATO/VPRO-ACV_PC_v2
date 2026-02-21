Attribute VB_Name = "Comun_Buscador"
Option Explicit

'==========================================================================
' MÓDULO : Comun_Buscador
' FUNCIÓN: Búsqueda genérica de celdas por texto en cualquier hoja.
'          Usado por VCA_ESP, VCA_POR y futuros procesos.
'          TODAS las búsquedas usan UCase → no distingue may/min.
' DEPENDE: Comun_Constantes
'==========================================================================


'==========================================================================
' BuscarCeldaPorTexto
' Busca texto en toda la hoja (o rango limitado) y devuelve la celda.
' Devuelve Nothing si no encuentra.
'==========================================================================
Public Function BuscarCeldaPorTexto(ByVal ws           As Worksheet, _
                                     ByVal texto         As String, _
                                     Optional ByVal busqueda    As XlLookAt = xlPart, _
                                     Optional ByVal rangoLimite As Range = Nothing) As Range
    Dim rngBuscar As Range
    Dim celda     As Range
    Dim textoBusc As String

    textoBusc = UCase(Trim(texto))

    If rangoLimite Is Nothing Then
        Set rngBuscar = ws.UsedRange
    Else
        Set rngBuscar = rangoLimite
    End If

    For Each celda In rngBuscar
        If busqueda = xlWhole Then
            If UCase(Trim(celda.Text)) = textoBusc Then
                Set BuscarCeldaPorTexto = celda
                Exit Function
            End If
        Else
            If InStr(1, UCase(celda.Text), textoBusc) > 0 Then
                Set BuscarCeldaPorTexto = celda
                Exit Function
            End If
        End If
    Next celda

    Set BuscarCeldaPorTexto = Nothing
End Function


'==========================================================================
' BuscarEnFila
' Busca texto en una fila concreta y devuelve la celda.
' Útil para localizar columnas dinámicas cuando ya sabemos la fila.
'==========================================================================
Public Function BuscarEnFila(ByVal ws       As Worksheet, _
                               ByVal fila     As Long, _
                               ByVal texto    As String, _
                               Optional ByVal busqueda As XlLookAt = xlPart) As Range
    Dim ultimaCol  As Long
    Dim col        As Long
    Dim textoBusc  As String
    Dim valorCelda As String

    textoBusc = UCase(Trim(texto))
    ultimaCol = ws.Cells(fila, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To ultimaCol
        valorCelda = UCase(Trim(ws.Cells(fila, col).Text))
        If busqueda = xlWhole Then
            If valorCelda = textoBusc Then
                Set BuscarEnFila = ws.Cells(fila, col)
                Exit Function
            End If
        Else
            If InStr(1, valorCelda, textoBusc) > 0 Then
                Set BuscarEnFila = ws.Cells(fila, col)
                Exit Function
            End If
        End If
    Next col

    Set BuscarEnFila = Nothing
End Function


'==========================================================================
' ObtenerColumnaReal
' Si la celda es combinada devuelve la columna izquierda del área.
' Si no es combinada devuelve su propia columna.
'==========================================================================
Public Function ObtenerColumnaReal(ByVal celda As Range) As Long
    If celda.MergeCells Then
        ObtenerColumnaReal = celda.MergeArea.Column
    Else
        ObtenerColumnaReal = celda.Column
    End If
End Function


'==========================================================================
' ObtenerFilaReal
' Si la celda es combinada devuelve la fila superior del área.
' Si no es combinada devuelve su propia fila.
'==========================================================================
Public Function ObtenerFilaReal(ByVal celda As Range) As Long
    If celda.MergeCells Then
        ObtenerFilaReal = celda.MergeArea.Row
    Else
        ObtenerFilaReal = celda.Row
    End If
End Function


'==========================================================================
' EsCeldaCombinada
'==========================================================================
Public Function EsCeldaCombinada(ByVal celda As Range) As Boolean
    EsCeldaCombinada = celda.MergeCells
End Function


'==========================================================================
' EsColorAzul
' Detecta cualquier tono azul sin depender de un RGB exacto.
' Usado por VCA_POR para detectar subcabeceras y por BuscarPrimeraFilaDatos.
'==========================================================================
Public Function EsColorAzul(ByVal colorCelda As Long) As Boolean
    Dim r As Long, g As Long, b As Long

    ' xlNone (-4142) y blanco no son azul
    If colorCelda = -4142 Or colorCelda = RGB(255, 255, 255) Then
        EsColorAzul = False
        Exit Function
    End If

    ' Protección contra valores negativos inesperados
    ' (pueden aparecer con colores temáticos no resueltos)
    If colorCelda < 0 Then
        EsColorAzul = False
        Exit Function
    End If

    ' VBA almacena colores en orden BGR
    r = colorCelda Mod 256
    g = (colorCelda \ 256) Mod 256
    b = (colorCelda \ 65536) Mod 256

    EsColorAzul = (b > r + 30)
End Function


'==========================================================================
' EsColorAzulCelda
' Versión segura que recibe la celda directamente.
' Comprueba ColorIndex antes de leer Color, evitando problemas con
' colores temáticos o formatos condicionales.
'==========================================================================
Public Function EsColorAzulCelda(ByVal celda As Range) As Boolean
    Dim colorCelda As Long

    ' Si no tiene color de fondo → no es azul
    If celda.Interior.ColorIndex = xlNone Then
        EsColorAzulCelda = False
        Exit Function
    End If

    ' Leer el color RGB resuelto
    colorCelda = celda.Interior.Color
    EsColorAzulCelda = EsColorAzul(colorCelda)
End Function


'==========================================================================
' BuscarPrimeraFilaDatos
' Desde una fila de cabecera, baja hasta la primera fila de datos:
'   - Salta 1 fila vacía si la hay
'   - Salta 1 fila azul (subcabecera) si la hay
' saltarAzul=False para hojas sin subcabecera de color (como ESP).
'==========================================================================
Public Function BuscarPrimeraFilaDatos(ByVal ws        As Worksheet, _
                                        ByVal filaCab   As Long, _
                                        ByVal col       As Long, _
                                        Optional ByVal saltarAzul As Boolean = True) As Long
    Dim fila As Long
    fila = filaCab + 1

    If Trim(ws.Cells(fila, col).Value) = "" Then
        fila = fila + 1
    End If

    If saltarAzul Then
        If EsColorAzulCelda(ws.Cells(fila, col)) Then
            fila = fila + 1
        End If
    End If

    BuscarPrimeraFilaDatos = fila
End Function
