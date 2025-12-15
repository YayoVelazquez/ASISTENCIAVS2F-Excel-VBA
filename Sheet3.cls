Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' ================== PARÁMETROS ==================
Private Const VAC_CODE As Long = 2                 ' código de vacaciones
Private Const WINDOW_DAYS As Long = 15             ' bloquear HOY+1 .. HOY+15 (15 días exactos)
Private Const START_ENFORCE As Date = #9/1/2025#   ' empezar a aplicar
Private Const TARGET_YEAR As Long = 2025

' Layout semanal 7 días + 1 celda en blanco
Private Const WEEK_DAYS As Long = 7
Private Const WEEK_STRIDE As Long = 8              ' 7 + 1 hueco

' ============== MAPEOS (LOCAL) ==============
Private Function LineasQ4_Local() As Variant
    LineasQ4_Local = Array("SE28", "SU10", "SU24", "SCU34", "SCU33", "SCU48", "KGT22", "SCU35")
End Function

Private Function RangoPorMesLinea_Local(ByVal MES$, ByVal linea$) As String
    Dim k$: k = UCase$(MES) & "|" & UCase$(linea)
    Select Case True
        ' ---- SEPTIEMBRE ----
        Case k = "SEPTIEMBRE|SE28": RangoPorMesLinea_Local = "HK6:IR35"
        Case k = "SEPTIEMBRE|SU10": RangoPorMesLinea_Local = "HK37:IR89"
        Case k = "SEPTIEMBRE|SU24": RangoPorMesLinea_Local = "HK91:IR138"
        Case k = "SEPTIEMBRE|SCU34": RangoPorMesLinea_Local = "HK140:IR160"
        Case k = "SEPTIEMBRE|SCU33": RangoPorMesLinea_Local = "HK162:IR191"
        Case k = "SEPTIEMBRE|SCU48": RangoPorMesLinea_Local = "HK193:IR219"
        Case k = "SEPTIEMBRE|KGT22": RangoPorMesLinea_Local = "HK221:IR255"
        Case k = "SEPTIEMBRE|SCU35": RangoPorMesLinea_Local = "HK257:IR267"
        ' ---- OCTUBRE ----
        Case k = "OCTUBRE|SE28": RangoPorMesLinea_Local = "IS6:KA35"
        Case k = "OCTUBRE|SU10": RangoPorMesLinea_Local = "IS37:KA89"
        Case k = "OCTUBRE|SU24": RangoPorMesLinea_Local = "IS91:KA138"
        Case k = "OCTUBRE|SCU34": RangoPorMesLinea_Local = "IS140:KA160"
        Case k = "OCTUBRE|SCU33": RangoPorMesLinea_Local = "IS162:KA191"
        Case k = "OCTUBRE|SCU48": RangoPorMesLinea_Local = "IS193:KA219"
        Case k = "OCTUBRE|KGT22": RangoPorMesLinea_Local = "IS221:KA255"
        Case k = "OCTUBRE|SCU35": RangoPorMesLinea_Local = "IS257:KA267"
        ' ---- NOVIEMBRE ----
        Case k = "NOVIEMBRE|SE28": RangoPorMesLinea_Local = "KB6:LI35"
        Case k = "NOVIEMBRE|SU10": RangoPorMesLinea_Local = "KB37:LI89"
        Case k = "NOVIEMBRE|SU24": RangoPorMesLinea_Local = "KB91:LI138"
        Case k = "NOVIEMBRE|SCU34": RangoPorMesLinea_Local = "KB140:LI160"
        Case k = "NOVIEMBRE|SCU33": RangoPorMesLinea_Local = "KB162:LI191"
        Case k = "NOVIEMBRE|SCU48": RangoPorMesLinea_Local = "KB193:LI219"
        Case k = "NOVIEMBRE|KGT22": RangoPorMesLinea_Local = "KB221:LI255"
        Case k = "NOVIEMBRE|SCU35": RangoPorMesLinea_Local = "KB257:LI267"
        ' ---- DICIEMBRE ----
        Case k = "DICIEMBRE|SE28": RangoPorMesLinea_Local = "LK6:MO35"
        Case k = "DICIEMBRE|SU10": RangoPorMesLinea_Local = "LK37:MO89"
        Case k = "DICIEMBRE|SU24": RangoPorMesLinea_Local = "LK91:MO138"
        Case k = "DICIEMBRE|SCU34": RangoPorMesLinea_Local = "LK140:MO160"
        Case k = "DICIEMBRE|SCU33": RangoPorMesLinea_Local = "LK162:MO191"
        Case k = "DICIEMBRE|SCU48": RangoPorMesLinea_Local = "LK193:MO219"
        Case k = "DICIEMBRE|KGT22": RangoPorMesLinea_Local = "LK221:MO255"
        Case k = "DICIEMBRE|SCU35": RangoPorMesLinea_Local = "LK257:MO267"
    End Select
End Function

' ============== HELPERS DE FECHA/COLUMNA ==============
Private Function DaysInMonthLocal(ByVal y As Long, ByVal m As Long) As Long
    DaysInMonthLocal = Day(DateSerial(y, m + 1, 0))
End Function

Private Function DayFromCol_7plus1(ByVal firstCol As Long, ByVal col As Long, ByVal MES As Long) As Long
    Dim off As Long, pos As Long, weekIdx As Long, dayIx As Long, dimMes As Long
    off = col - firstCol: If off < 0 Then Exit Function
    pos = off Mod WEEK_STRIDE
    If pos >= WEEK_DAYS Then Exit Function
    weekIdx = off \ WEEK_STRIDE
    dayIx = weekIdx * WEEK_DAYS + pos + 1
    dimMes = DaysInMonthLocal(TARGET_YEAR, MES)
    If dayIx < 1 Or dayIx > dimMes Then Exit Function
    DayFromCol_7plus1 = dayIx
End Function

Private Function DateFromCell(ByVal c As Range) As Date
    Dim Ls, i As Long, m As Long, MES$, addr$, rng As Range, firstCol As Long, dIx As Long
    Ls = LineasQ4_Local()
    For m = 9 To 12
        Select Case m
            Case 9: MES$ = "SEPTIEMBRE"
            Case 10: MES$ = "OCTUBRE"
            Case 11: MES$ = "NOVIEMBRE"
            Case 12: MES$ = "DICIEMBRE"
        End Select
        For i = LBound(Ls) To UBound(Ls)
            addr$ = RangoPorMesLinea_Local(MES$, CStr(Ls(i)))
            If Len(addr$) > 0 Then
                On Error Resume Next: Set rng = Me.Range(addr$): On Error GoTo 0
                If Not rng Is Nothing Then
                    If Not Intersect(c, rng) Is Nothing Then
                        firstCol = rng.Column
                        dIx = DayFromCol_7plus1(firstCol, c.Column, m)
                        If dIx > 0 Then
                            DateFromCell = DateSerial(TARGET_YEAR, m, dIx)
                            Exit Function
                        End If
                    End If
                End If
                Set rng = Nothing
            End If
        Next i
    Next m
End Function

' ======== BLOQUEO PRINCIPAL ========
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Fin
    If Intersect(Target, Me.UsedRange) Is Nothing Then Exit Sub

    Dim freezeStart As Date, freezeEnd As Date
    freezeStart = Date + 1
    freezeEnd = Date + WINDOW_DAYS

    ' Pegado masivo
    If Target.Cells.CountLarge > 1 Then
        If AttemptHasBlocked2(Target, freezeStart, freezeEnd) Then
            Application.EnableEvents = False
            Application.Undo
            Application.EnableEvents = True
            MsgBox "VACACIONES (2) bloqueadas para fechas entre " & _
                   Format(freezeStart, "dd-mmm") & " y " & Format(freezeEnd, "dd-mmm") & ".", _
                   vbExclamation, "Bloqueo"
        End If
        Exit Sub
    End If

    ' Edición individual
    Dim cel As Range, v, d As Date
    For Each cel In Target.Cells
        v = cel.Value
        If Len(v) > 0 And CLng(val(v)) = VAC_CODE Then
            d = DateFromCell(cel)
            If d > 0 And d >= START_ENFORCE Then
                If d >= freezeStart And d <= freezeEnd Then
                    Application.EnableEvents = False
                    Application.Undo
                    Application.EnableEvents = True
                    MsgBox "Vacaciones (2) solo con 15 días de anticipación." & vbCrLf & _
                           "Fecha ingresada: " & Format(d, "dd-mmm-yyyy"), _
                           vbExclamation, "Bloqueado"
                End If
            End If
        End If
    Next cel

Fin:
    On Error Resume Next
    Application.EnableEvents = True
End Sub

Private Function AttemptHasBlocked2(ByVal rng As Range, ByVal freezeStart As Date, ByVal freezeEnd As Date) As Boolean
    Dim cel As Range, v, d As Date
    For Each cel In rng.Cells
        v = cel.Value
        If Len(v) > 0 And CLng(val(v)) = VAC_CODE Then
            d = DateFromCell(cel)
            If d > 0 And d >= START_ENFORCE Then
                If d >= freezeStart And d <= freezeEnd Then
                    AttemptHasBlocked2 = True
                    Exit Function
                End If
            End If
        End If
    Next cel
End Function


