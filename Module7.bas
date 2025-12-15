Attribute VB_Name = "Module7"
Option Explicit

' ============================================================================================
' MÓDULO DE AUTOMATIZACIÓN DE AUSENTISMO Y ASISTENCIAS (AGO–DIC 2025)
' Autor: Yael Velázquez Artolozaga – MOE2 (Bosch Querétaro)
'
' Resumen:
'   - Lee bloques por línea/mes desde la hoja fuente (WS_DATOS_DEFAULT, p.ej. "VSM2").
'   - Construye/actualiza la tabla "ConteoTbl" en hojas mensuales "MES_AÑO".
'   - Calcula agregados y 4 columnas de fórmulas (% y totales).
'   - Crea/actualiza 4 gráficos (Ausencias, Control, Operativo+Target, Composición).
'   - Respeta posición/formato si los gráficos ya existen.
'   - Estabilidad: backup, soft-lock, wrappers de rendimiento, logging.
' ============================================================================================

'==================== CONFIGURACIÓN ====================
Private Const WS_DATOS_DEFAULT As String = "VSM2"     ' Hoja fuente con los códigos por día
Private Const TBL_NAME As String = "ConteoTbl"        ' Nombre fijo de la tabla mensual
Private Const DO_FIX_LAYOUT As Boolean = False        ' True: shapes libres (no se mueven con filtros)
Private Const PLANT_NAME As String = "VS2"            ' Para título en gráfico Operativo

' Estabilidad / seguridad
Private Const ENABLE_BACKUP As Boolean = True         ' Copia del libro antes de correr
Private Const CLEAN_BELOW As Boolean = False          ' Limpieza debajo de tabla (opcional)
Private Const LOG_WARNINGS As Boolean = True          ' Registrar avisos en MASTER_LOG
Private Const REQUIRE_SOFT_LOCK As Boolean = True     ' Candado suave para evitar corrida concurrente

' Año y meses a procesar
Private Const TARGET_YEAR As Long = 2025
Private Const MONTH_START As Long = 8                 ' 8 = AGOSTO
Private Const MONTH_END As Long = 12                  ' 12 = DICIEMBRE

'==================== TARGET (línea de meta en gráfico Operativo) ====================
Private Const TARGET_ENABLED As Boolean = True
Private Const TARGET_VALUE As Double = 0.93           ' 93%
Private Const TARGET_SERIES_NAME As String = "Target 93%"
Private Const TARGET_HELPER_COL_OFFSET As Long = 2    ' Separación a la derecha de la tabla p/rango helper

' ------------------------------------------------------------------------------------
' HeadersArr: Encabezados fijos de la tabla "ConteoTbl"
' ------------------------------------------------------------------------------------
Private Function HeadersArr() As Variant
    HeadersArr = Array( _
        "Line", "0", "6", "7", "2", "4", "5", "8", "9", "10", _
        "Justificadas", "Injustificadas", "Bajas", _
        "Bajo control", "Fuera de control", _
        "Mes", "Año", _
        "Asistencias (1)", "Asist+Justif", "Total días", _
        "%Asistencia", "%Injustificadas", "%Justificadas" _
    )
End Function

'==================== UTILIDADES BÁSICAS ====================

' Mes en mayúsculas (ES)
Private Function MesNombreUpper_es(ByVal n&) As String
    Select Case n
        Case 1:  MesNombreUpper_es = "ENERO"
        Case 2:  MesNombreUpper_es = "FEBRERO"
        Case 3:  MesNombreUpper_es = "MARZO"
        Case 4:  MesNombreUpper_es = "ABRIL"
        Case 5:  MesNombreUpper_es = "MAYO"
        Case 6:  MesNombreUpper_es = "JUNIO"
        Case 7:  MesNombreUpper_es = "JULIO"
        Case 8:  MesNombreUpper_es = "AGOSTO"
        Case 9:  MesNombreUpper_es = "SEPTIEMBRE"
        Case 10: MesNombreUpper_es = "OCTUBRE"
        Case 11: MesNombreUpper_es = "NOVIEMBRE"
        Case 12: MesNombreUpper_es = "DICIEMBRE"
    End Select
End Function

' ¿Existe la hoja?
Private Function SheetExists(ByVal nm As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(nm) Is Nothing
    On Error GoTo 0
End Function

' Devuelve/crea hoja por nombre
Private Function EnsureSheet(ByVal name$) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.name = name
    End If
End Function

' Logging de advertencias en MASTER_LOG (timestamp + mensaje)
Private Sub LogWarn(msg$)
    If Not LOG_WARNINGS Then Exit Sub
    Dim ws As Worksheet, r&
    Set ws = EnsureSheet("MASTER_LOG")
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "Fecha/Hora": ws.Cells(1, 2).Value = "Mensaje"
    End If
    r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = msg
End Sub

'==================== LÍNEAS A PROCESAR ====================

' Set de líneas consolidado para Q4 (ajustar si se agregan líneas)
Private Function LineasQ4() As Variant
    LineasQ4 = Array("SE28", "SU10", "SU24", "SCU34", "SCU33", "SCU48", "KGT22", "SCU35")
End Function

'==================== MAPA DE RANGOS (AGO–DIC 2025) ====================

' Normaliza un rango "A10:A1" -> "A1:A10" (si no puede parsear, devuelve s)
Private Function NormalizeRangeStr(ByVal s As String) As String
    On Error GoTo Fallback
    Dim parts() As String, a$, b$, ac$, ar&, bc$, br&, i&, p&
    parts = Split(Replace(s, " ", ""), ":")
    If UBound(parts) <> 1 Then NormalizeRangeStr = s: Exit Function

    a = UCase$(parts(0)): b = UCase$(parts(1))

    p = 0
    For i = 1 To Len(a)
        If Mid$(a, i, 1) Like "#" Then p = i: Exit For
    Next i
    ac = left$(a, p - 1): ar = CLng(Mid$(a, p))

    p = 0
    For i = 1 To Len(b)
        If Mid$(b, i, 1) Like "#" Then p = i: Exit For
    Next i
    bc = left$(b, p - 1): br = CLng(Mid$(b, p))

    If ar <= br Then
        NormalizeRangeStr = ac & ar & ":" & bc & br
    Else
        NormalizeRangeStr = ac & br & ":" & bc & ar
    End If
    Exit Function
Fallback:
    NormalizeRangeStr = s
End Function

' Mapea (MES|LÍNEA) ? rango en WS_DATOS_DEFAULT (actualiza si cambia layout de la fuente)
Private Function RangoPorMesLinea(ByVal MES$, ByVal linea$) As String
    Dim k$: k = UCase$(MES) & "|" & UCase$(linea)

    Select Case True
        ' ---- AGOSTO ----
        Case k = "AGOSTO|SE28": RangoPorMesLinea = "FZ6:HI35"
        Case k = "AGOSTO|SU10": RangoPorMesLinea = "FZ37:HI89"
        Case k = "AGOSTO|SU24": RangoPorMesLinea = "FZ91:HI138"
        Case k = "AGOSTO|SCU34": RangoPorMesLinea = "FZ140:HI160"
        Case k = "AGOSTO|SCU33": RangoPorMesLinea = "FZ162:HI191"
        Case k = "AGOSTO|SCU48": RangoPorMesLinea = "FZ193:HI219"
        Case k = "AGOSTO|KGT22": RangoPorMesLinea = "FZ221:HI255"
        Case k = "AGOSTO|SCU35": RangoPorMesLinea = "FZ257:HI267"

        ' ---- SEPTIEMBRE ----
        Case k = "SEPTIEMBRE|SE28": RangoPorMesLinea = "HK6:IR35"
        Case k = "SEPTIEMBRE|SU10": RangoPorMesLinea = "HK37:IR89"
        Case k = "SEPTIEMBRE|SU24": RangoPorMesLinea = "HK91:IR138"
        Case k = "SEPTIEMBRE|SCU34": RangoPorMesLinea = "HK140:IR160"
        Case k = "SEPTIEMBRE|SCU33": RangoPorMesLinea = "HK162:IR191"
        Case k = "SEPTIEMBRE|SCU48": RangoPorMesLinea = "HK193:IR219"
        Case k = "SEPTIEMBRE|KGT22": RangoPorMesLinea = "HK221:IR255"
        Case k = "SEPTIEMBRE|SCU35": RangoPorMesLinea = "HK257:IR267"

        ' ---- OCTUBRE ----
        Case k = "OCTUBRE|SE28": RangoPorMesLinea = "IS6:KA35"
        Case k = "OCTUBRE|SU10": RangoPorMesLinea = "IS37:KA89"
        Case k = "OCTUBRE|SU24": RangoPorMesLinea = "IS91:KA138"
        Case k = "OCTUBRE|SCU34": RangoPorMesLinea = "IS140:KA160"
        Case k = "OCTUBRE|SCU33": RangoPorMesLinea = "IS162:KA191"
        Case k = "OCTUBRE|SCU48": RangoPorMesLinea = "IS193:KA219"
        Case k = "OCTUBRE|KGT22": RangoPorMesLinea = "IS221:KA255"
        Case k = "OCTUBRE|SCU35": RangoPorMesLinea = "IS257:KA267"

        ' ---- NOVIEMBRE ----
        Case k = "NOVIEMBRE|SE28": RangoPorMesLinea = "KB6:LI35"
        Case k = "NOVIEMBRE|SU10": RangoPorMesLinea = "KB37:LI89"
        Case k = "NOVIEMBRE|SU24": RangoPorMesLinea = "KB91:LI138"
        Case k = "NOVIEMBRE|SCU34": RangoPorMesLinea = "KB140:LI160"
        Case k = "NOVIEMBRE|SCU33": RangoPorMesLinea = "KB162:LI191"
        Case k = "NOVIEMBRE|SCU48": RangoPorMesLinea = "KB193:LI219"
        Case k = "NOVIEMBRE|KGT22": RangoPorMesLinea = "KB221:LI255"
        Case k = "NOVIEMBRE|SCU35": RangoPorMesLinea = "KB257:LI267"

        ' ---- DICIEMBRE ----
        Case k = "DICIEMBRE|SE28": RangoPorMesLinea = "LK6:MO35"
        Case k = "DICIEMBRE|SU10": RangoPorMesLinea = "LK37:MO89"
        Case k = "DICIEMBRE|SU24": RangoPorMesLinea = "LK91:MO138"
        Case k = "DICIEMBRE|SCU34": RangoPorMesLinea = "LK140:MO160"
        Case k = "DICIEMBRE|SCU33": RangoPorMesLinea = "LK162:MO191"
        Case k = "DICIEMBRE|SCU48": RangoPorMesLinea = "LK193:MO219"
        Case k = "DICIEMBRE|KGT22": RangoPorMesLinea = "LK221:MO255"
        Case k = "DICIEMBRE|SCU35": RangoPorMesLinea = "LK257:MO267"
    End Select
End Function

'==================== CONTEOS DESDE RANGO ====================

' Recorre el rango y suma ocurrencias por código (0,1,2,4,5,6,7,8,9,10).
' Calcula agregados operativos: Justificadas, Injustificadas, Bajas, Bajo/Fuera control,
' Asistencias (1) y Asist+Justif. *El código 10 NO impacta agregados*.
Private Function ContarCodigos(rng As Range) As Variant
    Dim r As Range, v
    Dim c0&, c1&, c6&, c7&, c2&, c4&, c5&, c8&, c9&, c10&
    Dim justif&, injustif&, bajas&, bajoCtrl&, fueraCtrl&, asist&, asistMasJustif&

    For Each r In rng.Cells
        v = Trim$(CStr(r.Value))
        If Len(v) > 0 And IsNumeric(v) Then
            Select Case CLng(v)
                Case 0:  c0 = c0 + 1
                Case 1:  c1 = c1 + 1
                Case 6:  c6 = c6 + 1
                Case 7:  c7 = c7 + 1
                Case 2:  c2 = c2 + 1
                Case 4:  c4 = c4 + 1
                Case 5:  c5 = c5 + 1
                Case 8:  c8 = c8 + 1
                Case 9:  c9 = c9 + 1
                Case 10: c10 = c10 + 1   ' CAMBIO DE TURNO (solo informativo)
            End Select
        End If
    Next r

    ' === El 10 NO suma en agregados ===
    justif = c2 + c4 + c5 + c8 + c9            ' (antes incluía + c10)
    injustif = c0 + c6
    bajas = c7
    bajoCtrl = c2 + c4 + c5                    ' (antes incluía + c10)
    fueraCtrl = c0 + c6 + c8 + c9
    asist = c1
    asistMasJustif = c1 + justif               ' (no suma c10)

    ' Notar que devolvemos c10 (para la columna "10") pero sin impacto en %/agregados.
    ContarCodigos = Array(c0, c6, c7, c2, c4, c5, c8, c9, c10, _
                          justif, injustif, bajas, bajoCtrl, fueraCtrl, asist, asistMasJustif)
End Function

'==================== TABLA: CREAR/ACTUALIZAR ====================

' Crea/asegura la tabla "ConteoTbl" con encabezados correctos y limpia cuerpo si existe
Private Sub EnsureMonthTable(ws As Worksheet)
    Dim headers, nCols&, lo As ListObject, lastRow&, rng As Range
    headers = HeadersArr()
    nCols = UBound(headers) + 1

    ws.Cells(1, 1).Resize(1, nCols).Value = headers

    On Error Resume Next
    Set lo = ws.ListObjects(TBL_NAME)
    On Error GoTo 0

    If lo Is Nothing Then
        ' Primera vez: crea con 1 fila dummy
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(2, nCols))
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.name = TBL_NAME
        lo.TableStyle = "TableStyleMedium2"
    Else
        ' Mantener ancho correcto y no encoger tabla
        lastRow = lo.Range.Row + lo.Range.rows.Count - 1
        If lastRow < 2 Then lastRow = 2
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, nCols))

        If lo.HeaderRowRange.Row <> 1 Or lo.HeaderRowRange.Column <> 1 _
           Or lo.Range.Columns.Count <> nCols Then
            lo.Resize rng
        End If

        If Not lo.DataBodyRange Is Nothing Then
            lo.DataBodyRange.ClearContents
        End If
    End If
End Sub

' Ajusta # de filas de la tabla al requerido por el # de líneas
Private Sub EnsureTableRowCount(ws As Worksheet, ByVal rowsNeeded&)
    Dim lo As ListObject, cur&
    Set lo = ws.ListObjects(TBL_NAME)
    cur = lo.ListRows.Count
    Do While cur < rowsNeeded: lo.ListRows.Add: cur = cur + 1: Loop
    Do While cur > rowsNeeded And cur > 0: lo.ListRows(cur).Delete: cur = cur - 1: Loop
End Sub

' Limpieza opcional de residuos por debajo de la tabla
Private Sub CleanupBelowTable(ws As Worksheet)
    If Not CLEAN_BELOW Then Exit Sub
    Dim lo As ListObject, lastTableRow&, lastUsedRow&
    On Error Resume Next
    Set lo = ws.ListObjects(TBL_NAME)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    lastTableRow = lo.Range.Row + lo.Range.rows.Count - 1
    lastUsedRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row

    If lastUsedRow > lastTableRow Then
        ws.Range(ws.Cells(lastTableRow + 1, 1), _
                 ws.Cells(lastUsedRow, lo.Range.Column + lo.Range.Columns.Count - 1)).Clear
    End If
End Sub

'==================== BÚSQUEDA DE COLUMNAS ====================

Private Function FindColExact(ws As Worksheet, key$) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then FindColExact = f.Column
End Function

Private Function FindCol(ws As Worksheet, key$) As Long
    Dim f As Range
    Set f = ws.rows(1).Find(What:=key, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    If Not f Is Nothing Then FindCol = f.Column
End Function

' Escribe "val" en la fila "rowNum" bajo el encabezado "exactHeader"
' Si no existe, intenta coincidencia parcial con "partialFallback"
Private Sub PutByHeader(ws As Worksheet, ByVal rowNum As Long, ByVal exactHeader As String, ByVal val, Optional ByVal partialFallback As String = "")
    Dim c As Long
    c = FindColExact(ws, exactHeader)
    If c = 0 And Len(partialFallback) > 0 Then c = FindCol(ws, partialFallback)
    If c > 0 Then ws.Cells(rowNum, c).Value = val
End Sub

' ¿Existe la columna en el ListObject?
Private Function ColumnExists(lo As ListObject, ByVal headerName As String) As Boolean
    On Error Resume Next
    ColumnExists = Not lo.ListColumns(headerName) Is Nothing
    On Error GoTo 0
End Function

'==================== FÓRMULAS (4 columnas finales) ====================

' Reaplica fórmulas de: Total días, %Asistencia, %Injustificadas, %Justificadas
' (Total días = Asistencias(1) + Justificadas + Injustificadas; el 10 no está en Justificadas)
Private Sub ReapplyFourFormulas(ws As Worksheet, lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim hAsis$, hJust$, hInj$
    hAsis = IIf(ColumnExists(lo, "Asistencias (1)"), "Asistencias (1)", IIf(ColumnExists(lo, "Asistencias"), "Asistencias", ""))
    hJust = IIf(ColumnExists(lo, "Justificadas"), "Justificadas", "")
    hInj = IIf(ColumnExists(lo, "Injustificadas"), "Injustificadas", "")
    If hAsis = "" Or hJust = "" Or hInj = "" Then Exit Sub

    With lo.ListColumns("Total días").DataBodyRange
        .Formula = "=IFERROR([@[" & hAsis & "]]+[@[" & hJust & "]]+[@[" & hInj & "]],0)"
    End With
    With lo.ListColumns("%Asistencia").DataBodyRange
        .Formula = "=IFERROR([@[" & hAsis & "]]/[@[Total días]],0)"
        .NumberFormat = "0.00%"
    End With
    With lo.ListColumns("%Injustificadas").DataBodyRange
        .Formula = "=IFERROR([@[" & hInj & "]]/[@[Total días]],0)"
        .NumberFormat = "0.00%"
    End With
    With lo.ListColumns("%Justificadas").DataBodyRange
        .Formula = "=IFERROR([@[" & hJust & "]]/[@[Total días]],0)"
        .NumberFormat = "0.00%"
    End With
End Sub

'==================== GRÁFICOS (NO mover/NO re-formatear si ya existen) ====================

' Crea/obtiene gráfico "chName". Solo posiciono/tamaño la 1ª vez. PlotVisibleOnly=True.
Private Function EnsureChart(ws As Worksheet, ByVal chName As String, _
    ByVal left As Double, ByVal top As Double, _
    ByVal width As Double, ByVal height As Double, _
    Optional ByRef wasCreated As Boolean = False) As ChartObject

    Dim co As ChartObject
    On Error Resume Next
    Set co = ws.ChartObjects(chName)
    On Error GoTo 0

    If co Is Nothing Then
        wasCreated = True
        Set co = ws.ChartObjects.Add(left, top, width, height)
        co.name = chName
    Else
        wasCreated = False
    End If

    co.Chart.PlotVisibleOnly = True
    Set EnsureChart = co
End Function

' Busca una serie por su nombre visible (resuelve si Name es fórmula)
Private Function FindSeriesByName(co As ChartObject, ByVal seriesName As String) As Series
    Dim sc As Series, nm As String
    For Each sc In co.Chart.SeriesCollection
        nm = sc.name
        If left$(nm, 1) = "=" Then On Error Resume Next: nm = CStr(Evaluate(nm)): On Error GoTo 0
        If StrComp(Trim$(nm), seriesName, vbTextCompare) = 0 Then
            Set FindSeriesByName = sc
            Exit Function
        End If
    Next sc
End Function

' Asegura/actualiza una serie vinculada a la columna "colName" de la tabla
Private Sub UpdateOrAddSeries(co As ChartObject, lo As ListObject, ByVal seriesTitle As String, ByVal colName As String)
    Dim sc As Series
    Set sc = FindSeriesByName(co, seriesTitle)
    If sc Is Nothing Then
        Set sc = co.Chart.SeriesCollection.NewSeries
        sc.name = seriesTitle
    End If
    sc.Values = lo.ListColumns(colName).DataBodyRange
    If ColumnExists(lo, "Line") Then sc.XValues = lo.ListColumns("Line").DataBodyRange
End Sub

' === Rangos auxiliares para la línea Target (93%) y para "ceros" de CAMBIO DE TURNO ===

' Helper constante con TARGET_VALUE (mismo # de filas que la tabla)
Private Function EnsureTargetHelperRange(ws As Worksheet, lo As ListObject) As Range
    Dim firstRow&, lastRow&, tgtCol&, r As Range
    If lo.DataBodyRange Is Nothing Then Exit Function
    firstRow = lo.DataBodyRange.Row
    lastRow = lo.DataBodyRange.rows(lo.DataBodyRange.rows.Count).Row
    tgtCol = lo.Range.Column + lo.Range.Columns.Count + TARGET_HELPER_COL_OFFSET
    Set r = ws.Range(ws.Cells(firstRow, tgtCol), ws.Cells(lastRow, tgtCol))
    r.Value = TARGET_VALUE
    r.NumberFormat = "0.00%"
    Set EnsureTargetHelperRange = r
End Function

' Helper de CEROS (mismo # de filas) para alimentar la serie "CAMBIO DE TURNO" sin peso
Private Function EnsureZeroHelperRange(ws As Worksheet, lo As ListObject) As Range
    Dim firstRow&, lastRow&, tgtCol&, r As Range
    If lo.DataBodyRange Is Nothing Then Exit Function
    firstRow = lo.DataBodyRange.Row
    lastRow = lo.DataBodyRange.rows(lo.DataBodyRange.rows.Count).Row
    tgtCol = lo.Range.Column + lo.Range.Columns.Count + TARGET_HELPER_COL_OFFSET + 1 ' otra columna helper
    Set r = ws.Range(ws.Cells(firstRow, tgtCol), ws.Cells(lastRow, tgtCol))
    r.Value = 0
    Set EnsureZeroHelperRange = r
End Function

' Asegura la serie de Target (línea) en el gráfico Operativo
Private Sub EnsureTargetSeries_Operativo(ws As Worksheet, lo As ListObject, co As ChartObject)
    If Not TARGET_ENABLED Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim rngTarget As Range, sc As Series
    Set rngTarget = EnsureTargetHelperRange(ws, lo)
    If rngTarget Is Nothing Then Exit Sub

    Set sc = FindSeriesByName(co, TARGET_SERIES_NAME)
    If sc Is Nothing Then
        Set sc = co.Chart.SeriesCollection.NewSeries
        sc.name = TARGET_SERIES_NAME
    End If
    sc.Values = rngTarget
    If ColumnExists(lo, "Line") Then sc.XValues = lo.ListColumns("Line").DataBodyRange
    On Error Resume Next
    sc.ChartType = xlLine
    sc.AxisGroup = xlPrimary
    On Error GoTo 0
End Sub

' 1) Gráfico: Ausencias (Justificadas / Injustificadas / Bajas)
Private Sub EnsureChart_Ausencias(ws As Worksheet, lo As ListObject, _
    ByVal L As Double, ByVal T As Double, ByVal W As Double, ByVal H As Double)

    Dim co As ChartObject, created As Boolean
    Set co = EnsureChart(ws, "chAusencias", L, T, W, H, created)
    If created Then
        co.Chart.ChartType = xlColumnClustered
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = "Ausencias por Línea (Justificadas vs Injustificadas vs Bajas)"
        co.Chart.Legend.Position = xlLegendPositionBottom
        On Error Resume Next: co.Chart.ApplyDataLabels: On Error GoTo 0
    End If

    UpdateOrAddSeries co, lo, "Justificadas", "Justificadas"
    UpdateOrAddSeries co, lo, "Injustificadas", "Injustificadas"
    UpdateOrAddSeries co, lo, "Bajas", "Bajas"
End Sub

' 2) Gráfico: Control (Bajo control vs Fuera de control)
Private Sub EnsureChart_Control(ws As Worksheet, lo As ListObject, _
    ByVal L As Double, ByVal T As Double, ByVal W As Double, ByVal H As Double)

    Dim co As ChartObject, created As Boolean
    Set co = EnsureChart(ws, "chControl", L, T, W, H, created)
    If created Then
        co.Chart.ChartType = xlColumnClustered
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = "Bajo Control vs Fuera de Control por Línea"
        co.Chart.Legend.Position = xlLegendPositionBottom
        On Error Resume Next: co.Chart.ApplyDataLabels: On Error GoTo 0
    End If

    UpdateOrAddSeries co, lo, "Bajo control", "Bajo control"
    UpdateOrAddSeries co, lo, "Fuera de control", "Fuera de control"
End Sub

' 3) Gráfico: Operativo (stacked 100%) + línea Target 93%
Private Sub EnsureChart_Operativo(ws As Worksheet, lo As ListObject, _
    ByVal L As Double, ByVal T As Double, ByVal W As Double, ByVal H As Double)

    Dim co As ChartObject, created As Boolean, MES$, ANIO$
    ParseMonthYearFromSheet ws.name, MES, ANIO

    Set co = EnsureChart(ws, "chOperativo", L, T, W, H, created)
    If created Then
        co.Chart.ChartType = xlColumnStacked100
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = "Asistencia y Ausentismo Operativo – " & PLANT_NAME & " | " & MES & " " & ANIO
        co.Chart.Legend.Position = xlLegendPositionBottom
        On Error Resume Next
        co.Chart.Axes(xlValue).TickLabels.NumberFormat = "0%"
        co.Chart.ApplyDataLabels
        On Error GoTo 0
    End If

    UpdateOrAddSeries co, lo, "%Asistencia", "%Asistencia"
    UpdateOrAddSeries co, lo, "%Injustificadas", "%Injustificadas"
    UpdateOrAddSeries co, lo, "%Justificadas", "%Justificadas"

    EnsureTargetSeries_Operativo ws, lo, co   ' Línea 93%
End Sub

' 4) Gráfico: Composición por Código (incluye serie "CAMBIO DE TURNO" con valores 0)
Private Sub EnsureChart_ComposicionCodigos(ws As Worksheet, lo As ListObject, _
    ByVal L As Double, ByVal T As Double, ByVal W As Double, ByVal H As Double)

    Dim co As ChartObject, created As Boolean
    Dim names, cols, i&, colName$, serieTitle$

    ' Renombrado del último: "CAMBIO DE TURNO" (antes "TIEMPO POR TIEMPO")
    names = Array("FALTA", "SUSPENSIÓN", "BAJA", "PERMISO SIN GOCE", "MÉDICO", "INCAPACIDAD", "VACACIONES", "PERMISO CON GOCE", "CAMBIO DE TURNO")
    cols = Array("0", "6", "7", "5", "8", "9", "2", "4", "10")

    Set co = EnsureChart(ws, "chComposicion", L, T, W, H, created)
    If created Then
        co.Chart.ChartType = xlColumnClustered
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = "Composición de Ausencias por Código"
        co.Chart.Legend.Position = xlLegendPositionBottom
        On Error Resume Next: co.Chart.ApplyDataLabels: On Error GoTo 0
    End If

    ' Si existía la serie con nombre viejo, eliminarla
    On Error Resume Next
    Dim old As Series
    Set old = FindSeriesByName(co, "TIEMPO POR TIEMPO")
    If Not old Is Nothing Then old.Delete
    On Error GoTo 0

    For i = LBound(cols) To UBound(cols)
        colName = CStr(cols(i))
        serieTitle = CStr(names(i))

        If colName = "10" Then
            ' === Código 10 SIN PESO: alimentar con CEROS (solo para mostrar nombre en la leyenda) ===
            Dim sc As Series, zr As Range
            Set sc = FindSeriesByName(co, "CAMBIO DE TURNO")
            If sc Is Nothing Then
                Set sc = co.Chart.SeriesCollection.NewSeries
                sc.name = "CAMBIO DE TURNO"
            End If
            Set zr = EnsureZeroHelperRange(ws, lo)
            If Not zr Is Nothing Then sc.Values = zr
            If ColumnExists(lo, "Line") Then sc.XValues = lo.ListColumns("Line").DataBodyRange
        Else
            UpdateOrAddSeries co, lo, serieTitle, colName
        End If
    Next i
End Sub

' Parsea nombre de hoja "MES_AÑO" ? MES, ANIO
Private Sub ParseMonthYearFromSheet(ByVal sheetName As String, ByRef MES$, ByRef ANIO$)
    Dim p&: p = InStr(1, sheetName, "_")
    If p > 0 Then
        MES = Replace(left$(sheetName, p - 1), "_", "")
        ANIO = Mid$(sheetName, p + 1)
    Else
        MES = sheetName: ANIO = ""
    End If
End Sub

' Orquesta la creación/actualización de los 4 gráficos
Private Sub EnsureAllCharts(ws As Worksheet, lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim X As Double, y As Double, W As Double, H As Double, GAP As Double
    X = lo.Range.left + lo.Range.width + 20
    y = lo.Range.top
    W = 560: H = 280: GAP = 20

    EnsureChart_Ausencias ws, lo, X, y, W, H
    EnsureChart_Control ws, lo, X + W + GAP, y, W, H
    EnsureChart_Operativo ws, lo, X, y + H + GAP, W, H
    EnsureChart_ComposicionCodigos ws, lo, X + W + GAP, y + H + GAP, W, H
End Sub

'==================== WRAPPERS DE RENDIMIENTO/SEGURIDAD ====================

' Acelerar (apaga pantalla/eventos/cálculo)
Private Sub SafeCalc_Enter()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' Restaurar estado normal
Private Sub SafeCalc_Exit()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' Copia de seguridad del libro (SaveCopyAs con timestamp)
Private Sub BackupWorkbookIfNeeded()
    If Not ENABLE_BACKUP Then Exit Sub
    On Error Resume Next
    If Len(ThisWorkbook.Path) > 0 Then
        ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\Backup_" & Format(Now, "yyyymmdd_HHMMSS") & ".xlsm"
    End If
    On Error GoTo 0
End Sub

'==================== SOFT-LOCK (candado suave) ====================

' Asegura hoja oculta "_CTRL" para token de bloqueo
Private Function EnsureCtrlSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("_CTRL"): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.name = "_CTRL"
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1").Value = "LOCK_TOKEN"
        ws.Range("B1").Value = "LOCK_TIME"
    End If
    Set EnsureCtrlSheet = ws
End Function

' Token pseudoaleatorio
Private Function NewToken() As String
    Randomize
    NewToken = "T" & Hex$(CLng((Rnd * 2 ^ 31))) & Hex$(CLng((Rnd * 2 ^ 31)))
End Function

' Intenta adquirir candado (expira 20 min). True si se obtiene.
Private Function SoftLock_Acquire(ByRef myToken As String) As Boolean
    If Not REQUIRE_SOFT_LOCK Then SoftLock_Acquire = True: Exit Function
    Dim ws As Worksheet, tok$, ts As Double
    Set ws = EnsureCtrlSheet()
    tok = CStr(ws.Range("A2").Value)
    ts = val(ws.Range("B2").Value)
    If Len(tok) > 0 And Now - ts < TimeSerial(0, 20, 0) Then
        SoftLock_Acquire = False
        Exit Function
    End If
    myToken = NewToken()
    ws.Range("A2").Value = myToken
    ws.Range("B2").Value = Now
    SoftLock_Acquire = True
End Function

' Libera candado si el token coincide
Private Sub SoftLock_Release(ByVal myToken As String)
    If Not REQUIRE_SOFT_LOCK Then Exit Sub
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = EnsureCtrlSheet()
    If CStr(ws.Range("A2").Value) = myToken Then
        ws.Range("A2").ClearContents
        ws.Range("B2").ClearContents
    End If
    On Error GoTo 0
End Sub

'==================== FILL PER MONTH ====================

' Range seguro por dirección (devuelve Nothing si inválida)
Private Function GetRangeSafe(ws As Worksheet, ByVal addr$) As Range
    On Error Resume Next
    Set GetRangeSafe = ws.Range(addr)
    On Error GoTo 0
End Function

' Núcleo de llenado mensual (tabla, conteos, fórmulas y gráficos)
Private Sub FillMonth(ByVal MES$, ByVal ANIO&)
    Dim ws As Worksheet, lo As ListObject, wsDatos As Worksheet
    Dim lines, i&, fila&, addr$, res As Variant, rng As Range

    If Not SheetExists(WS_DATOS_DEFAULT) Then
        LogWarn "Hoja fuente '" & WS_DATOS_DEFAULT & "' no existe. Omitido " & MES & " " & ANIO
        Exit Sub
    End If

    Set ws = EnsureSheet(UCase$(MES) & "_" & CStr(ANIO))
    Set wsDatos = ThisWorkbook.Worksheets(WS_DATOS_DEFAULT)

    ' 1) Estructura base de la tabla
    EnsureMonthTable ws

    ' 2) Aforo de filas según líneas
    lines = LineasQ4()
    EnsureTableRowCount ws, (UBound(lines) - LBound(lines) + 1)
    Set lo = ws.ListObjects(TBL_NAME)

    ' 3) Llenado por línea desde el rango mapeado (validando direcciones)
    fila = lo.DataBodyRange.Row
    For i = LBound(lines) To UBound(lines)

        addr = NormalizeRangeStr(RangoPorMesLinea(MES, CStr(lines(i))))
        If Len(addr) = 0 Then LogWarn "Sin rango para " & MES & "|" & CStr(lines(i)): GoTo SiguienteLinea

        Set rng = GetRangeSafe(wsDatos, addr)
        If rng Is Nothing Then LogWarn "Rango inválido '" & addr & "' en " & WS_DATOS_DEFAULT & " para " & MES & "|" & CStr(lines(i)): GoTo SiguienteLinea

        res = ContarCodigos(rng)
        If IsArray(res) Then
            ' === Escribir fila completa (códigos, agregados y metadatos) ===
            PutByHeader ws, fila, "Line", CStr(lines(i))
            PutByHeader ws, fila, "0", res(0)
            PutByHeader ws, fila, "6", res(1)
            PutByHeader ws, fila, "7", res(2)
            PutByHeader ws, fila, "2", res(3)
            PutByHeader ws, fila, "4", res(4)
            PutByHeader ws, fila, "5", res(5)
            PutByHeader ws, fila, "8", res(6)
            PutByHeader ws, fila, "9", res(7)
            PutByHeader ws, fila, "10", res(8)
            PutByHeader ws, fila, "Justificadas", res(9)
            PutByHeader ws, fila, "Injustificadas", res(10)
            PutByHeader ws, fila, "Bajas", res(11)
            PutByHeader ws, fila, "Bajo control", res(12)
            PutByHeader ws, fila, "Fuera de control", res(13)
            PutByHeader ws, fila, "Mes", MES
            PutByHeader ws, fila, "Año", ANIO
            PutByHeader ws, fila, "Asistencias (1)", res(14), "Asistencias"
            PutByHeader ws, fila, "Asist+Justif", res(15), "Asist"
        End If

SiguienteLinea:
        fila = fila + 1
    Next i

    ' 4) Fórmulas de totales/porcentajes (las 4 últimas columnas)
    ReapplyFourFormulas ws, lo

    ' 5) (Opcional) Shapes libres para no moverse con filtros
    If DO_FIX_LAYOUT Then
        Dim shp As Shape
        For Each shp In ws.Shapes
            On Error Resume Next: shp.Placement = xlFreeFloating: On Error GoTo 0
        Next shp
    End If

    ' 6) Gráficos vinculados a la tabla
    EnsureAllCharts ws, lo

    ' 7) Limpieza opcional
    CleanupBelowTable ws
End Sub

'==================== MACRO MAESTRO (GLOBAL) ====================

' Orquesta todo el proceso de AGO–DIC 2025
Public Sub MASTER()
    Dim oldCalc As XlCalculation, oldScr As Boolean, oldEvt As Boolean
    Dim token$, gotLock As Boolean

    If Not SheetExists(WS_DATOS_DEFAULT) Then
        MsgBox "La hoja fuente '" & WS_DATOS_DEFAULT & "' no existe. Cancelo.", vbExclamation, "MASTER"
        Exit Sub
    End If

    If ENABLE_BACKUP Then BackupWorkbookIfNeeded

    If REQUIRE_SOFT_LOCK Then
        gotLock = SoftLock_Acquire(token)
        If Not gotLock Then
            MsgBox "Otro usuario podría estar editando (candado activo). Intenta más tarde.", vbExclamation, "MASTER"
            Exit Sub
        End If
    End If

    ' Guardar estado UI/cálculo para restaurar al final
    oldCalc = Application.Calculation
    oldScr = Application.ScreenUpdating
    oldEvt = Application.EnableEvents

    On Error GoTo FALLA
    SafeCalc_Enter

    Dim y&, m&, MES$
    y = TARGET_YEAR
    For m = MONTH_START To MONTH_END
        MES = MesNombreUpper_es(m)
        FillMonth MES, y
    Next m

    SafeCalc_Exit
    Application.CalculateFull
    ThisWorkbook.Save
    MsgBox "Actualizado: tablas, fórmulas y gráficos", vbInformation, "MASTER"
    GoTo SALIR

FALLA:
    SafeCalc_Exit
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error en " & MES & " " & y & ": " & Err.Description, vbExclamation, "MASTER"

SALIR:
    If REQUIRE_SOFT_LOCK Then SoftLock_Release token
    Application.EnableEvents = oldEvt
    Application.ScreenUpdating = oldScr
    Application.Calculation = oldCalc
End Sub


