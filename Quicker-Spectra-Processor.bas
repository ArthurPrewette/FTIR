Attribute VB_Name = "Module11"
Sub SpectraCorrectplot1()
Attribute SpectraCorrectplot1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SpectraCorrect-plot Macro
'
    Columns("A:B").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveSheet.Next.Activate
    Rows("1:1").Select
    ActiveSheet.Paste
    Range("C2").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("B1").Select
    Selection.End(xlDown).Select
    Range("C1870").Select
    Application.CutCopyMode = False
    Range("C1870").Select
    ActiveCell.FormulaR1C1 = "=Sheet1!R[-1]C-(Sheet1!R1869C-Sheet1!R1869C3)"
    Range("C1869").Select
    ActiveCell.FormulaR1C1 = _
        "=(Sheet1!R[-1]C-(Sheet1!R1869C-R1869C3))*(1-R1C1)+R1C1*RC"
    Rows("1870:1870").Select
    Selection.Delete Shift:=xlUp
    Range("C1869").Select
    ActiveCell.FormulaR1C1 = "=Sheet1!RC-(Sheet1!R1869C-Sheet1!R1869C3)"
    Range("C1868").Select
    ActiveCell.FormulaR1C1 = _
        "=(Sheet1!RC-(Sheet1!R1869C-R1869C3))*(1-R1C1)+R1C1*R[1]C"
    Range("C1868").Select
    Selection.AutoFill Destination:=Range("C2:C1868"), Type:=xlFillDefault
    Range("C2:C1868").Select
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:AH2"), Type:=xlFillDefault
    Range("C2:AH2").Select
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("D2:AH2").Select
    Selection.AutoFill Destination:=Range("D2:AH1869"), Type:=xlFillDefault
    Range("D2:AH1869").Select
    Range("C1869").Select
    Selection.AutoFill Destination:=Range("C1869:AH1869"), Type:=xlFillDefault
    Range("C1869:AH1869").Select
    Range("AH1869").Select
    ActiveCell.FormulaR1C1 = "=Sheet1!RC-(Sheet1!R1869C-Sheet1!R1869C3)"
    Range("AF1868").Select
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1829
    ActiveWindow.ScrollRow = 1815
    ActiveWindow.ScrollRow = 1813
    ActiveWindow.ScrollRow = 1811
    ActiveWindow.ScrollRow = 1809
    ActiveWindow.ScrollRow = 1807
    ActiveWindow.ScrollRow = 1804
    ActiveWindow.ScrollRow = 1807
    ActiveWindow.ScrollRow = 1809
    ActiveWindow.ScrollRow = 1811
    ActiveWindow.ScrollRow = 1813
    ActiveWindow.ScrollRow = 1815
    ActiveWindow.ScrollRow = 1822
    ActiveWindow.ScrollRow = 1824
    ActiveWindow.ScrollRow = 1822
    ActiveWindow.ScrollRow = 1824
    ActiveWindow.ScrollRow = 1826
    ActiveWindow.ScrollRow = 1829
    Range("J1866").Select
    ActiveCell.FormulaR1C1 = _
        "=(Sheet1!RC-(Sheet1!R1869C-R1869C3))*(1-R1C1)+R1C1*R[1]C"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "0.1"
        Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A1").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    Selection.Font.Size = 16
    Selection.Font.Size = 18
    Selection.Font.Size = 20
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Smoothing factor ^"
    Range("A2:A4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Columns("A:A").ColumnWidth = 11.29
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    Selection.Font.Size = 12
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "WaveNumber"
    Columns("B:B").EntireColumn.AutoFit
    
    Dim lColumn As Long
    Dim iCntr As Long
    lColumn = 60
    For iCntr = lColumn To 1 Step -1
        If Cells(1, iCntr) = Empty Then
            Columns(iCntr).Delete
        End If
    Next
    
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveWindow.ScrollRow = 1828
    ActiveWindow.ScrollRow = 1821
    ActiveWindow.ScrollRow = 1799
    ActiveWindow.ScrollRow = 1766
    ActiveWindow.ScrollRow = 913
    ActiveWindow.ScrollRow = 624
    ActiveWindow.ScrollRow = 514
    ActiveWindow.ScrollRow = 258
    ActiveWindow.ScrollRow = 221
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 1
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=ActiveSheet.Range("$B$1:$AH$1869")
    ActiveSheet.Shapes("Chart 1").IncrementLeft -111
    ActiveSheet.Shapes("Chart 1").IncrementTop -87.75
    ActiveChart.ApplyChartTemplate ( _
        "C:\Users\lukep\AppData\Roaming\Microsoft\Templates\Charts\ffftir.crtx")
    With ActiveSheet.Shapes("Chart 1").Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MaximumScale = 2200
    ActiveChart.Axes(xlCategory).MinimumScale = 1750
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = -0.01
    ActiveChart.Axes(xlValue).MaximumScale = 0.02
    ActiveChart.Axes(xlValue).MinimumScale = -0.05
    ActiveChart.Axes(xlValue).MaximumScale = 0.04
    ActiveChart.Axes(xlValue).MaximumScale = 0.1
    ActiveChart.Axes(xlValue).MajorUnitIsAuto = True
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 1").IncrementLeft 286.5
    ActiveSheet.Shapes("Chart 1").IncrementTop -6.75
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
    ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
    ActiveChart.Axes(xlValue).MinimumScale = -0.1
    ActiveChart.Axes(xlValue).MaximumScale = 0.2
    ActiveChart.Axes(xlCategory).Select
    Selection.MajorTickMark = xlInside
    Selection.MinorTickMark = xlInside
    ActiveChart.Axes(xlValue).Select
    Selection.MajorTickMark = xlInside
    Selection.MinorTickMark = xlInside
    ActiveChart.Axes(xlValue).CrossesAt = -0.1
    ActiveChart.PlotArea.Select
    ActiveChart.Legend.Select
    Selection.Height = 27.477
    Selection.Top = 229.397
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlValue).Select
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select

End Sub
