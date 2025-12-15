Attribute VB_Name = "Module1"
Public Sub FixChartPercentLabels()
    Dim ws As Worksheet, co As ChartObject, sc As Series
    Dim pt As Point, i As Long
    
    ' Recorre todas las hojas (meses AGO–DIC)
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set co = ws.ChartObjects("chOperativo")
        On Error GoTo 0
        
        If Not co Is Nothing Then
            ' Asegura que sea 100% Stacked
            co.Chart.ChartType = xlColumnStacked100
            co.Chart.Axes(xlValue).TickLabels.NumberFormat = "0%"
            
            ' Recorre series (%Asistencia, %Injustificadas, %Justificadas)
            For Each sc In co.Chart.SeriesCollection
                sc.ApplyDataLabels
                With sc.DataLabels
                    .ShowValue = True
                    .NumberFormat = "0.00%"   ' Fuerza formato %
                    .Position = xlLabelPositionInsideEnd
                End With
            Next sc
        End If
    Next ws
End Sub

