Attribute VB_Name = "Module8"
Option Explicit

' ====== CONFIG ======
Private Const pwd As String = "YAELVELAZQUEZ"
' ====================

Public Sub LIMPIEZAPROTECCION()
    Dim ws As Worksheet
    Set ws = ActiveSheet   ' o fija por nombre: Set ws = ThisWorkbook.Worksheets("S28")

    ' Protección: usuario restringido, macro con permiso
    EnsureUIOnlyProtection ws, pwd

    ' Bloques (igual que tu Office Script)
    Dim ADDRS As Variant
    ADDRS = Array( _
        "FZ6:HI35", "FZ37:HI89", "FZ91:HI138", "FZ140:HI160", _
        "FZ162:HI191", "FZ193:HI219", "FZ221:HI255", "FZ257:HI267", _
        "HK6:IR35", "HK37:IR89", "HK91:IR138", "HK140:IR160", _
        "HK162:IR191", "HK193:IR219", "HK221:IR255", "HK257:IR267", _
        "IS6:KA35", "IS37:KA89", "IS91:KA138", "IS140:KA160", _
        "IS162:KA191", "IS193:KA219", "IS221:KA255", "IS257:KA267", _
        "KB6:LI35", "KB37:LI89", "KB91:LI138", "KB140:LI160", _
        "KB162:LI191", "KB193:LI219", "KB221:LI255", "KB257:LI267", _
        "LK6:MO35", "LK37:MO89", "LK91:MO138", "LK140:MO160", _
        "LK162:MO191", "LK193:MO219", "LK221:MO255", "LK257:MO267" _
    )

    ' Rendimiento
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean, prevEvents As Boolean
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents

    On Error GoTo FAILSAFE
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim i As Long
    For i = LBound(ADDRS) To UBound(ADDRS)
        ProcessRange_Ultimo7 ActiveSheet.Range(CStr(ADDRS(i)))
    Next i

FAILSAFE:
    ' Restaura siempre
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        MsgBox "Ocurrió un error (" & Err.Number & "): " & Err.Description, vbExclamation, "Macro"
    End If
End Sub

' === LÓGICA: último 7; borra TODO a la izquierda EXCEPTO los 7 anteriores ===
Private Sub ProcessRange_Ultimo7(ByVal rng As Range)
    If rng Is Nothing Then Exit Sub

    Dim vals As Variant
    vals = rng.Value2  ' 2D (1-based)

    Dim rows As Long, cols As Long
    rows = UBound(vals, 1): cols = UBound(vals, 2)

    Dim r As Long, c As Long
    Dim v As Variant, refCol As Long
    Dim pos() As Long, k As Long

    For r = 1 To rows
        ' 1) Posiciones de todos los "7"
        k = 0
        Erase pos
        For c = 1 To cols
            v = vals(r, c)
            If IsSeven(v) Then
                k = k + 1
                ReDim Preserve pos(1 To k)
                pos(k) = c
            End If
        Next c

        If k = 0 Then GoTo NextRow   ' fila sin 7 -> no se toca

        ' 2) Último 7 como referencia
        refCol = pos(k)

        ' 3) Limpiar a la izquierda del último 7, PERO conservar los 7 anteriores
        If refCol > 1 Then
            For c = 1 To refCol - 1
                v = vals(r, c)
                If Not IsEmpty(v) Then
                    If CStr(v) <> vbNullString Then
                        If Not IsSeven(v) Then
                            vals(r, c) = vbNullString
                            rng.Cells(r, c).Interior.Color = vbRed
                        End If
                        ' Si es 7, lo dejamos intacto
                    End If
                End If
            Next c
        End If

NextRow:
    Next r

    ' 4) Escribir cambios (permitido por UserInterfaceOnly)
    rng.Value2 = vals
End Sub

' === Helper: detecta 7 numérico o "7" como texto (con espacios) ===
Private Function IsSeven(ByVal v As Variant) As Boolean
    If IsNumeric(v) Then
        IsSeven = (v = 7)
    ElseIf VarType(v) = vbString Then
        IsSeven = (Trim$(CStr(v)) = "7")
    Else
        IsSeven = False
    End If
End Function

' === Protección: usuario bloqueado, macro con permiso ===
Private Sub EnsureUIOnlyProtection(ByVal ws As Worksheet, ByVal pwd As String)
    On Error Resume Next

    ' Quita protección actual (si hay)
    If ws.ProtectContents Then
        ws.Unprotect pwd
    End If

    ' Reaplica protección con UserInterfaceOnly:=True
    ws.Protect password:=pwd, UserInterfaceOnly:=True, _
               AllowFormattingCells:=True, _
               AllowUsingPivotTables:=True, _
               AllowFiltering:=True, _
               AllowSorting:=False, _
               AllowInsertingRows:=False, _
               AllowInsertingColumns:=False, _
               AllowDeletingRows:=False, _
               AllowDeletingColumns:=False, _
               AllowFormattingColumns:=False, _
               AllowFormattingRows:=False
    On Error GoTo 0
End Sub





