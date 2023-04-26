Sub SpectraCorrectplot1()
'
' SpectraCorrect-plot Macro
'
    Dim Template As String
    Template = "C:\Users\lukep\AppData\Roaming\Microsoft\Templates\Charts\ffftir.crtx" 'path to your chart template
    
    Dim endcol As Long
    Dim endrow As Long
    lastcol = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    Application.DisplayAlerts = False
    Columns("A:B").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Rows("1:1").Select
    Selection.Copy
    ActiveSheet.Next.Activate
    Rows("1:1").Select
    ActiveSheet.Paste
    
    Range("C1869").Select
    ActiveCell.FormulaR1C1 = "=Sheet1!RC-(Sheet1!R1869C-Sheet1!R1869C3)"
    Selection.AutoFill Destination:=Range("C1869:AH1869"), Type:=xlFillDefault
    
    Range("C1868").Select
    ActiveCell.FormulaR1C1 = "=(Sheet1!RC-(Sheet1!R1869C-R1869C3))*(1-R1C1)+R1C1*R[1]C"
    Selection.AutoFill Destination:=Range("C1868:AH1868"), Type:=xlFillDefault
    Range("C1868:AH1868").Select
    Selection.AutoFill Destination:=Range("C2:AH1868"), Type:=xlFillDefault
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "0.1"
    Selection.Interior.Color = 255
    Selection.Font.Bold = True
    Selection.Font.Size = 20
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Smoothing factor ^"
    Range("A2:A4").Select
    Selection.Merge
    Columns("A:A").ColumnWidth = 12
    Selection.Interior.Color = 65535
    Selection.Font.Size = 12
    Selection.WrapText = True
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
    Dim chartrange As Range
    Set chartrange = Selection
    
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=chartrange
    ActiveChart.ApplyChartTemplate ( _
        Template)
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
    ActiveChart.Axes(xlCategory).MajorUnit = 50
    Selection.MajorTickMark = xlInside
    Selection.MinorTickMark = xlInside
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = -0.1
    ActiveChart.Axes(xlValue).MaximumScale = 0.2
    ActiveChart.Axes(xlValue).MajorUnit = 0.05
    ActiveChart.Axes(xlValue).CrossesAt = -0.1
    Selection.MajorTickMark = xlInside
    Selection.MinorTickMark = xlInside
    ActiveChart.Legend.Delete


    
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Wavenumber (cm^-1)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Wavenumber (cm^-1)"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 18).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 18).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Size = 20
    End With

    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Absorbance (A.U.)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Absorbance (A.U.)"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 17).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 17).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Size = 20
    End With
    ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinorMajor)
    ActiveChart.ChartArea.Select


End Sub
