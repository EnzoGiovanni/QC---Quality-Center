Option Explicit
Const SeptJours As Date = "06/01/1900"
Sub QC_PostProcessing()
    Dim Trash As Variant
    
    Trash = MoreFasterCode (True)

    'Make
    Trash = MakeTableAnoByWeek()
    Trash = CountingDefects(2, 4)
    Trash = CountingDefects(7, 5)
    
    'Formating Sheet
    FormatingSheet ("Anomalies")
    FindAllElt("-", ActiveWorkbook.Worksheets("Anomalies").Columns(7)).Cells.ClearContents
    FormatingSheet ("Linked")
    Trash = FormatingSheet ("ByWeek")
    
    Trash = MoreFasterCode (False)
    
End Sub
'====================================================================================================
Function CountingDefects(Switch As Long, DestCol As Long)
'====================================================================================================
    Dim ShAnos As Worksheet: Set ShAnos = ThisWorkbook.Worksheets("anomalies")
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
        .Add(After:=.Item(.Count), Type:=xlWorksheet).Name = "ByWeek"
    End With
    Dim ShBW As Worksheet: Set ShBW = ThisWorkbook.Worksheets("ByWeek")
    Dim ShLA As Worksheet: Set ShLA = ThisWorkbook.Worksheets("anomalies")
    
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
Function FormatingSheet(ShName As String)
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
