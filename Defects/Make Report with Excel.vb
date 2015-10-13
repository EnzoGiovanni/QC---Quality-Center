Option Explicit
Const SeptJours As Date = "06/01/1900"
Sub QC_PostProcessing()
    Dim Trash As Variant
    
    Trash = MoreFasterCode(True)
    
    'Make
    Trash = MakeTableAnoByWeek()
    Trash = CountingDefects(2, 4)
    Trash = CountingDefects(7, 5)
    Trash = Cumulate(4, 6)
    Trash = Cumulate(5, 7)
    Trash = Difference(6, 7, 8)
    
    'Formating Sheet
    FormatingSheet ("Defects")
    FindAllElt("-", ActiveWorkbook.Worksheets("Defects").Columns(7)).Cells.ClearContents
    FormatingSheet ("Linked")
    Trash = TitrateColumn()
    Trash = FormatingSheet("ByWeek")
    Trash = MakeGraphAnosByWeek()
    
    Trash = MoreFasterCode(False)
    
End Sub
'====================================================================================================
Function MakeGraphAnosByWeek()
'====================================================================================================
    Dim WkZone As Range
    With ThisWorkbook.Worksheets("ByWeek")
        Set WkZone = .Range(.Cells(1, 3), .Cells(.Cells(.Rows.Count, 1).End(xlUp).Row, 3))
        Set WkZone = WkZone.Application.Union(WkZone, .Range(.Cells(1, 8), .Cells(.Cells(.Rows.Count, 1).End(xlUp).Row, 8)))
    End With
    
    ThisWorkbook.Charts.Add().Name = "GrphByWeek"
    
    With ThisWorkbook.Charts("GrphByWeek")
        .Move after:=ThisWorkbook.Sheets.Item(ThisWorkbook.Sheets.Count)
        .Type = xlLine
        .SetSourceData Source:=WkZone, PlotBy:=xlColumns
    End With: Set WkZone = Nothing
    
    Dim Courbes As Series
    For Each Courbes In ThisWorkbook.Charts("GrphByWeek").SeriesCollection
        Courbes.Format.Line.Visible = msoFalse
        Courbes.Format.Line.Visible = msoTrue
        Courbes.MarkerStyle = xlMarkerStyleNone
    Next Courbes: Set Courbes = Nothing
    
    With ThisWorkbook.Charts("GrphByWeek")
        .SeriesCollection("Stock open defects").Format.Line.ForeColor.RGB = RGB(0, 0, 255)
        .SetElement (msoElementLegendTop)
        .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        .PageSetup.CenterHeader = "&D"
        
        'abscissa
        .Axes(xlCategory).HasMajorGridlines = True
        .Axes(xlCategory).MajorGridlines.Border.Color = RGB(217, 217, 217)
        .Axes(xlCategory).HasMinorGridlines = True
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Orientation = 55
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Year - Week Number"
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 8
        .Axes(xlCategory, xlPrimary).AxisTitle.Font.Bold = False
        
        'ordinates
        .Axes(xlValue).HasMinorGridlines = True
        .Axes(xlValue).MajorUnit = 5
        .Axes(xlValue).MinorUnit = 1
        .Axes(xlValue).MinorGridlines.Border.Color = RGB(217, 217, 217)
        .Axes(xlValue).TickLabels.Font.Size = 8
        
        .HasTitle = False
    End With
    
End Function
'====================================================================================================
Function TitrateColumn()
'====================================================================================================
    Dim ShBW As Worksheet: Set ShBW = ThisWorkbook.Worksheets("ByWeek")
    With ShBW
        .Cells(1, 4).Value = "number of opened defects"
        .Cells(1, 5).Value = "number of closed defects"
        .Cells(1, 6).Value = "Accumulated number of open defects"
        .Cells(1, 7).Value = "Accumulated number of close defects"
        .Cells(1, 8).Value = "Stock open defects"
    End With
    Set ShBW = Nothing
End Function
'====================================================================================================
Function Difference(ByRef ColSrc1 As Long, ByRef ColSrc2 As Long, ByRef ColDest As Long)
'====================================================================================================
    Dim ShBW As Worksheet: Set ShBW = ThisWorkbook.Worksheets("ByWeek")
    Dim WkLst As Range
    With ShBW
       Set WkLst = .Range(.Cells(2, ColDest), .Cells(.Cells(.Rows.Count, 1).End(xlUp).Row, ColDest))
    End With
    Dim Cellule As Range
    For Each Cellule In WkLst.Cells
        Cellule.Value = ShBW.Cells(Cellule.Row, ColSrc1).Value - ShBW.Cells(Cellule.Row, ColSrc2).Value
    Next Cellule: Set Cellule = Nothing
    Set WkLst = Nothing: Set ShBW = Nothing
End Function
'====================================================================================================
Function Cumulate(ByRef ColSrc As Long, ByRef ColDest As Long)
'====================================================================================================
    Dim ShBW As Worksheet: Set ShBW = ThisWorkbook.Worksheets("ByWeek")
    Dim WkLst As Range
    With ShBW
       Set WkLst = .Range(.Cells(2, ColSrc), .Cells(.Cells(.Rows.Count, ColSrc).End(xlUp).Row, ColSrc))
    End With
    Dim Cellule As Range
    For Each Cellule In WkLst.Cells
        If Cellule.Row = 2 Then
            ShBW.Cells(Cellule.Row, ColDest).Value = Cellule.Value
        Else
            ShBW.Cells(Cellule.Row, ColDest).Value = Cellule.Value + ShBW.Cells(Cellule.Row - 1, ColDest).Value
        End If
    Next Cellule: Set Cellule = Nothing
    Set WkLst = Nothing: Set ShBW = Nothing
End Function
'====================================================================================================
Function CountingDefects(ByRef Switch As Long, ByRef DestCol As Long)
'====================================================================================================
    Dim ShAnos As Worksheet: Set ShAnos = ThisWorkbook.Worksheets("Defects")
    Dim ShBW As Worksheet: Set ShBW = ThisWorkbook.Worksheets("ByWeek")
    Dim WkLst As Range
    With ShBW
       Set WkLst = .Range(.Cells(2, 3), .Cells(.Cells(.Rows.Count, 3).End(xlUp).Row, 3))
    End With
    Dim Week As Range
    For Each Week In WkLst
        ShBW.Cells(Week.Row, DestCol).Value = ThisWorkbook.Application.WorksheetFunction.CountIf(ShAnos.Columns(Switch), Week.Value)
    Next Week: Set Week = Nothing
    Set WkLst = Nothing: Set ShBW = Nothing: Set ShAnos = Nothing
End Function
'====================================================================================================
Function MakeTableAnoByWeek()
'====================================================================================================
    With ThisWorkbook.Worksheets
        .Add(after:=.Item(.Count), Type:=xlWorksheet).Name = "ByWeek"
    End With
    Dim ShBW As Worksheet: Set ShBW = ThisWorkbook.Worksheets("ByWeek")
    Dim ShLA As Worksheet: Set ShLA = ThisWorkbook.Worksheets("Defects")
    
    With ShBW
        .Cells(1, 1).Value = "Year"
        .Cells(1, 2).Value = "Week"
        .Cells(1, 3).Value = "Year-Week"
    End With
    
    Dim DateMin, DateMax As Date
    
    With ShLA
        DateMin = .Application.WorksheetFunction.Min(.Columns(6))
        DateMax = .Application.WorksheetFunction.Max(.Columns(6), .Columns(11))
    End With
    DateMin = DateMin - SeptJours
    DateMax = DateMax + 7 - DatePart("W", DateMax, vbMonday)
    Dim D As Date: D = DateMin
    Dim Ligne As Long: Ligne = 2
    Do
        With ShBW
            .Cells(Ligne, 1).Value = Year(D)
            .Cells(Ligne, 2).Value = WeekNumber(D)
            If .Cells(Ligne, 2).Value > 9 Then
                .Cells(Ligne, 3).Value = .Cells(Ligne, 1).Value & "-" & .Cells(Ligne, 2)
            Else
                .Cells(Ligne, 3).Value = .Cells(Ligne, 1).Value & "-" & "0" & .Cells(Ligne, 2)
            End If
        End With
    
        D = D + SeptJours
        Ligne = Ligne + 1
    Loop While D <= DateMax
    
    Set ShBW = Nothing: Set ShLA = Nothing
End Function
'====================================================================================================
Function FormatingSheet(ByRef ShName As String)
'====================================================================================================
    Dim Sh As Worksheet: Set Sh = ThisWorkbook.Worksheets(ShName)
    With Sh.Cells
        .Font.Name = "Arial"
        .Font.Size = 8
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Rows(1).AutoFilter
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
        .WrapText = True
    End With
    Sh.Select (True)
    With Sh
        .Rows(2).Select
        .Application.Windows(ThisWorkbook.Name).FreezePanes = True
    End With
End Function
'====================================================================================================
Function MoreFasterCode(ByRef Top As Boolean)
'====================================================================================================
    With Application
        If Top Then
            .DisplayAlerts = False: .ScreenUpdating = False: .EnableEvents = False: .Calculation = xlManual
        Else
            .Calculation = xlAutomatic: .EnableEvents = True: .ScreenUpdating = True: .DisplayAlerts = True
        End If
    End With
End Function
'====================================================================================================
Function FindAllElt(ByRef Element As Variant, ByRef Zone As Range) As Range
'====================================================================================================
    Dim Elt, Aire As Range: Dim FirstAdresse As String
    For Each Aire In Zone.Areas
        With Aire
            Set Elt = .Find(Element, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=True)
            If Not Elt Is Nothing Then
                FirstAdresse = Elt.Address()
                Do
                    If FindAllElt Is Nothing Then
                        Set FindAllElt = Elt
                    Else
                        Set FindAllElt = .Application.Union(FindAllElt, Elt)
                    End If
                    Set Elt = .FindNext(Elt)
                Loop While Elt.Address <> FirstAdresse And Not Elt Is Nothing
            End If
        End With
    Next Aire: Set Aire = Nothing: Set Elt = Nothing
End Function
'====================================================================================================
Function WeekNumber(ByRef D As Date)
'====================================================================================================
    'Calcul du n° de semaine selon la norme ISO, norme europe
    Dim date_jeudi, date_4_janvier, date_lundi_semaine_1 As Date
    Dim Nb_jours, numero As Integer
    date_jeudi = DateAdd("d", 4 - Weekday(D, vbMonday), D)
    date_4_janvier = DateSerial(Year(date_jeudi), 1, 4)
    date_lundi_semaine_1 = DateAdd("d", 1 - Weekday(date_4_janvier, vbMonday), date_4_janvier)
    Nb_jours = Abs(DateDiff("d", date_lundi_semaine_1, date_jeudi, vbMonday))
    WeekNumber = Int(Nb_jours / 7) + 1
End Function
