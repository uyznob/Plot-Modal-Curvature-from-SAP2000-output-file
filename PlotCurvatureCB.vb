Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit
Dim moDis()             As Single
Dim cell                As Variant
'Note: In SAP, joint 2 is the end joint
'when we divide one beam into m elements beam

Sub PlotCurvatureCB()
    Dim i, j, l As Integer
    Dim rng1 As Range
    Dim h As Double
    
    Call intoMm
    Call checkDuplicateName("Modal Displacement")
    Worksheets("Joint Displacements").Activate
    Call convertText2Number
    
    l = Application.WorksheetFunction.Max(Range(Range("A1"), Range("A1").End(xlDown)))
    ReDim moDis(1 To l, 1 To 5)
    
    'Take modal displacement to matrix moDis()
    Set rng1 = Range(Cells(4, 5), Cells(4, 5).End(xlDown))
    Call fillMatrixMoDis(rng1, l)
    
    'Display modal displacement
    Worksheets("Modal Displacement").Activate

    'Show modal displacement
    For i = 1 To l
        For j = 1 To 5
            Cells(i + 1, j + 1) = moDis(i, j)
        Next
        Cells(i + 1, 1) = i
    Next
    
    Call addTitle
    
    h = Range("G2").Value
    Call cmsTable(l, h)
    Call ncmdTable(l)
    Call addChart

End Sub
Public Sub checkDuplicateName(s As String)
    Dim ws As Worksheet
    Dim i As Integer
    For i = 1 To Worksheets.Count
       If Worksheets(i).Name = s Then
            Application.DisplayAlerts = False
            Sheets(i).Delete
            Application.DisplayAlerts = True
            Exit For
       End If
    Next i
    'Add "Modal Displacement" Sheet
    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = s
End Sub
Sub convertText2Number()
    'Convert the number in text format to number format
    Range(Range("A1"), Range("A1").End(xlDown)).NumberFormat = "0"
    Range(Range("A1"), Range("A1").End(xlDown)).Value = Range(Range("A1"), Range("A1").End(xlDown)).Value
End Sub
Sub addTitle()
    Range("A1").Value = "Modal displacement"
    Range("G1").Value = "h(mm)"
    Range("G2").Value = InputBox("Length of element (mm)")
    Range("H1").Value = "Curvature Modal Shape"
    Range("M1").Value = "Normalized Curvature Modal Shape"
    Range("A1:Q1").Interior.Color = rgbAquamarine
End Sub
Sub fillMatrixMoDis(rng As Range, l As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim container(1 To 5) As Single
    For Each cell In rng
        Select Case cell.Value
            Case 1
                i = cell.Offset(0, -4).Value
                moDis(i, 1) = cell.Offset(0, 3).Value
            Case 2
                i = cell.Offset(0, -4).Value
                moDis(i, 2) = cell.Offset(0, 3).Value
            Case 3
                i = cell.Offset(0, -4).Value
                moDis(i, 3) = cell.Offset(0, 3).Value
            Case 4
                i = cell.Offset(0, -4).Value
                moDis(i, 4) = cell.Offset(0, 3).Value
            Case 5
                i = cell.Offset(0, -4).Value
                moDis(i, 5) = cell.Offset(0, 3).Value
        End Select
    Next
    'Change the MoDis because Joint 2 is the end joint
    For i = 1 To 5
            container(i) = moDis(2, i)
    Next
    For i = 2 To l - 1
        For j = 1 To 5
            moDis(i, j) = moDis(i + 1, j)
        Next
    Next
    For i = 1 To 5
        moDis(l, i) = container(i)
    Next
End Sub
Sub intoMm()
    Dim disRange As Range
    Dim v As Single
    
    Sheets(1).Select
    If Range("F3").Value = "m" Then
        Range("F3:H3").Value = "mm"
        Set disRange = Range(Range("F4"), Range("F4").End(xlDown).Offset(0, 2))
        For Each cell In disRange
            cell.Value = cell.Value / 1000
        Next
    End If
End Sub
Sub addChart()
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim chartObj1 As ChartObject
    Dim chartObj2 As ChartObject
  
    'Set data range for the chart
    Set rng1 = Range(Range("A2"), Range("A2").End(xlDown))
    Set rng2 = Range(Range("B2"), Range("B2").End(xlDown).Offset(0, 4))
    Set rng3 = Range(Range("M2"), Range("M2").End(xlDown).Offset(0, 4))
        
    'Create a chart
    Set chartObj1 = ActiveSheet.ChartObjects.Add( _
                    Left:=Range("R2").Left, _
                    Width:=450, _
                    Top:=Range("R2").Top, _
                    Height:=250)
    Set chartObj2 = ActiveSheet.ChartObjects.Add( _
                    Left:=Range("R20").Left, _
                    Width:=450, _
                    Top:=Range("R20").Top, _
                    Height:=250)
     
    'Apply data to chart and determine type
    chartObj1.Chart.SetSourceData Source:=rng2
    chartObj1.Chart.ChartType = xlXYScatterLines
    chartObj2.Chart.SetSourceData Source:=rng3
    chartObj2.Chart.ChartType = xlXYScatterLines

    chartObj1.Chart.FullSeriesCollection(1).XValues = rng1
    chartObj2.Chart.FullSeriesCollection(1).XValues = rng1
    
    'Set charts name
    Set chartObj1 = ActiveSheet.ChartObjects("Chart 1")
    Set chartObj2 = ActiveSheet.ChartObjects("Chart 2")
    
    'Add Title
    chartObj1.Chart.HasTitle = True
    chartObj2.Chart.HasTitle = True
    chartObj1.Chart.ChartTitle.Text = "Modal Shape"
    chartObj2.Chart.ChartTitle.Text = "Curvature Modal Shape"
      
    'Format Font Type
    chartObj1.Chart.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Calibri"
    chartObj2.Chart.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Calibri"
  
    'Make Font Bold
    chartObj1.Chart.ChartArea.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    chartObj2.Chart.ChartArea.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    
    'Add name for legend
    chartObj1.Chart.FullSeriesCollection(1).Name = "Mode 1"
    chartObj1.Chart.FullSeriesCollection(2).Name = "Mode 2"
    chartObj1.Chart.FullSeriesCollection(3).Name = "Mode 3"
    chartObj1.Chart.FullSeriesCollection(4).Name = "Mode 4"
    chartObj1.Chart.FullSeriesCollection(5).Name = "Mode 5"
    
    chartObj2.Chart.FullSeriesCollection(1).Name = "Mode 1"
    chartObj2.Chart.FullSeriesCollection(2).Name = "Mode 2"
    chartObj2.Chart.FullSeriesCollection(3).Name = "Mode 3"
    chartObj2.Chart.FullSeriesCollection(4).Name = "Mode 4"
    chartObj2.Chart.FullSeriesCollection(5).Name = "Mode 5"
    
    ActiveWindow.Zoom = 70

End Sub
Sub cmsTable(l As Integer, h As Double)
'Populate curvature Modal Shape
    Dim i, j As Integer
    Dim x1 As Double
    Dim x2 As Double
    Dim x3 As Double
    
    For i = 2 To l - 1
        For j = 1 To 5
            x1 = CDbl(Cells(i + 2, j + 1).Value)
            x2 = CDbl(Cells(i + 1, j + 1).Value)
            x3 = CDbl(Cells(i, j + 1).Value)
            Cells(i + 1, j + 7).Value = cdm(x1, x2, x3, h)
        Next
    Next
    'Calculate curvature near the fixed restrain
    For i = 1 To 5
        Cells(2, i + 7) = 2 * Cells(3, i + 7) - Cells(4, i + 7)
        Cells(l + 1, i + 7) = 0
    Next

End Sub
Function cdm(x1 As Double, x2 As Double, x3 As Double, h As Double) As Double
'Center difference method
    cdm = (x1 - 2 * x2 + x3) / (h * h)
End Function
Sub ncmdTable(l As Integer)
'Normalize curvature modal shape table
    Dim i, j As Integer
    Dim temp As Double
  
    For i = 1 To l
        For j = 1 To 5
            temp = Cells(i + 1, j + 7).Value
            Cells(i + 1, j + 12).Value = temp / Cells(2, j + 7).Value
        Next
    Next
End Sub
