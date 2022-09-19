Private Sub Workbook_Open()
    open1_load1_close1_repete
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Range("A1:B8").Delete
    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects.Delete
    End If
 'If Me.Saved = False Then Me.Save
End Sub

Public Sub open1_load1_close1_repete()
'
' Open 1 CSV file, load it, close it and repete until list is empty or chart is full
'
Dim ochartObj As ChartObject
Dim oChart As Chart
Dim pathToCSVs As String
Dim n As Integer, NOKs As Byte
NOKs = 0
n = 1

' Delete old table of limits
Range("A1:B8").Delete

' Delete old charts
If ActiveSheet.ChartObjects.Count > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

' Create an empty plot
Set ochartObj = ActiveSheet.ChartObjects.Add(Top:=1, Left:=1, Width:=1220, Height:=550)
Set oChart = ochartObj.Chart
oChart.ChartType = xlXYScatterLinesNoMarkers
With oChart
 .HasTitle = True
 .ChartTitle.Text = "Kistler pressing"
End With

oChart.Axes(xlCategory).HasTitle = True
oChart.Axes(xlCategory).AxisTitle.Caption = "Stroke [mm]"

oChart.Axes(xlValue).HasTitle = True
oChart.Axes(xlValue).AxisTitle.Caption = "Force [kN]"

' Open CSV files to plot
pathToCSVs = ThisWorkbook.Path & "\CSV\"
CSVs = Dir(pathToCSVs & "*.csv")
Do While CSVs <> "" And n < 255 'Excel restriction to max 255 data series in one chart
    Workbooks.Open Filename:=pathToCSVs & CSVs, local:=True
    
    ' Load data to the plot
    If n = 1 Then
        ' Evaluation
        ThisWorkbook.Worksheets(1).Range("A1") = Workbooks(CSVs).Worksheets(1).Range("D99").Value
        ThisWorkbook.Worksheets(1).Range("B1") = Workbooks(CSVs).Worksheets(1).Range("D98").Value
        ThisWorkbook.Worksheets(1).Range("A2") = Workbooks(CSVs).Worksheets(1).Range("E99").Value
        ThisWorkbook.Worksheets(1).Range("B2") = Workbooks(CSVs).Worksheets(1).Range("D98").Value
        ThisWorkbook.Worksheets(1).Range("A3") = Workbooks(CSVs).Worksheets(1).Range("E99").Value
        ThisWorkbook.Worksheets(1).Range("B3") = Workbooks(CSVs).Worksheets(1).Range("E98").Value
        ThisWorkbook.Worksheets(1).Range("A4") = Workbooks(CSVs).Worksheets(1).Range("D99").Value
        ThisWorkbook.Worksheets(1).Range("B4") = Workbooks(CSVs).Worksheets(1).Range("E98").Value
        ThisWorkbook.Worksheets(1).Range("A5") = Workbooks(CSVs).Worksheets(1).Range("D99").Value
        ThisWorkbook.Worksheets(1).Range("B5") = Workbooks(CSVs).Worksheets(1).Range("D98").Value
        oChart.SeriesCollection.Add Source:=ThisWorkbook.Worksheets(1).Range("A1:A5")
        oChart.SeriesCollection(n).XValues = ThisWorkbook.Worksheets(1).Range("A1:A5")
        oChart.SeriesCollection(n).Values = ThisWorkbook.Worksheets(1).Range("B1:B5")
        oChart.SeriesCollection(n).Name = "Evaluation"
        With oChart.SeriesCollection(n)
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = 2
        End With
        
        ' Stop signal
        ThisWorkbook.Worksheets(1).Range("A7") = Workbooks(CSVs).Worksheets(1).Range("B120").Value
        ThisWorkbook.Worksheets(1).Range("B7") = Workbooks(CSVs).Worksheets(1).Range("C36").Value
        ThisWorkbook.Worksheets(1).Range("A8") = Workbooks(CSVs).Worksheets(1).Range("E24").Value
        ThisWorkbook.Worksheets(1).Range("B8") = Workbooks(CSVs).Worksheets(1).Range("C36").Value
        n = n + 1
        oChart.SeriesCollection.Add Source:=ThisWorkbook.Worksheets(1).Range("A7:A8")
        oChart.SeriesCollection(n).XValues = ThisWorkbook.Worksheets(1).Range("A7:A8")
        oChart.SeriesCollection(n).Values = ThisWorkbook.Worksheets(1).Range("B7:B8")
        oChart.SeriesCollection(n).Name = "Stop signal"
        With oChart.SeriesCollection(n)
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .Format.Line.Weight = 1
            .Format.Line.DashStyle = msoLineDash
        End With
        
        ' Axes and title settings
        oChart.ChartTitle.Text = Workbooks(CSVs).Worksheets(1).Range("D3").Value & " - " & _
            Workbooks(CSVs).Worksheets(1).Range("B4").Value & " - " & Workbooks(CSVs).Worksheets(1).Range("D5").Value
        oChart.Axes(xlCategory).MinimumScale = Workbooks(CSVs).Worksheets(1).Range("B24").Value 'osa X
        oChart.Axes(xlValue).MinimumScale = Workbooks(CSVs).Worksheets(1).Range("H24").Value 'osa Y
    End If
    
    ' Pressing curves
    n = n + 1
    oChart.SeriesCollection.Add Source:=Workbooks(CSVs).Worksheets(1).Range("A150:A1149")
    oChart.SeriesCollection(n).XValues = Workbooks(CSVs).Worksheets(1).Range("A150:A1149")
    oChart.SeriesCollection(n).Values = Workbooks(CSVs).Worksheets(1).Range("B150:B1149")
    oChart.SeriesCollection(n).Name = Workbooks(CSVs).Worksheets(1).Range("B10").Value & " " & _
        Workbooks(CSVs).Worksheets(1).Range("B7").Value & " " & Workbooks(CSVs).Worksheets(1).Range("B6").Text
    If Workbooks(CSVs).Worksheets(1).Range("B10").Value <> "OK" Then
        NOKs = NOKs + 1
    End If
    
    ' Close the CSV file w/o saving changes
    Workbooks(CSVs).Close SaveChanges:=False
    
    CSVs = Dir
Loop

MsgBox ("The chart shows " & n - 2 & " data series" & vbNewLine & _
    "which consists of " & NOKs & " NOK and " & n - 2 - NOKs & " OK curves")

End Sub
