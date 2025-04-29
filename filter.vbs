Private Function GetLanguageID() As Integer
    GetLanguageID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
End Function

' MUI for Ribbon Control: Labels
Sub GetLabel(ByVal control As IRibbonControl, ByRef Labeling As Variant)
    Dim languageID As Integer
    languageID = GetLanguageID()
    
    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")

    If languageID = 1058 Then 'Ukrainian
        labels.Add "groupSelFlt", "Фільтр за виділеним"
        labels.Add "FilterByText", "Текстом"
        labels.Add "FilterByBg", "Кольором клітинки"
        labels.Add "FilterByFt", "Кольором шрифту"
        
    ElseIf languageID = 1049 Then 'Russian
        labels.Add "groupSelFlt", "Фильтр по выделенному"
        labels.Add "FilterByText", "Тексту"
        labels.Add "FilterByBg", "Цвету ячейки"
        labels.Add "FilterByFt", "Цвету шрифта"
        
    Else 'English (default)
        labels.Add "groupSelFlt", "Filter by selection"
        labels.Add "FilterByText", "Text"
        labels.Add "FilterByBg", "Cell color"
        labels.Add "FilterByFt", "Font color"
    End If

    If labels.Exists(control.ID) Then
        Labeling = labels(control.ID)
    Else
        Labeling = "<Error>"
    End If
End Sub

' MUI for Ribbon Control: ScreenTips
Sub GetScreentip(ByVal control As IRibbonControl, ByRef screenTip As Variant)
    Dim languageID As Integer
    languageID = GetLanguageID()
    
    Dim tips As Object
    Set tips = CreateObject("Scripting.Dictionary")
    
    If languageID = 1058 Then 'Ukrainian
        tips.Add "FilterByText", "Фільтрувати за текстом у виділеній клітинці"
        tips.Add "FilterByBg", "Фільтрувати за кольором виділеної клітинки"
        tips.Add "FilterByFt", "Фільтрувати за кольором тексту виділеної клітинки"
        
    ElseIf languageID = 1049 Then 'Russian
        tips.Add "FilterByText", "Фильтровать по тексту в выделенной ячейке"
        tips.Add "FilterByBg", "Фильтровать по цвету выделенной ячейки"
        tips.Add "FilterByFt", "Фильтровать по цвету текста выделенной ячейки"
        
    Else 'English (default)
        tips.Add "FilterByText", "Filter selection text"
        tips.Add "FilterByBg", "Filter by selected cell color"
        tips.Add "FilterByFt", "Filter by selected text color"
    End If

    If tips.Exists(control.ID) Then
        screenTip = tips(control.ID)
    Else
        screenTip = "<Error>"
    End If
End Sub

' MUI for Ribbon Control: SuperTips
Sub GetSupertip(ByVal control As IRibbonControl, ByRef superTip As Variant)
    Dim languageID As Integer
    languageID = GetLanguageID()
    
    Dim superTips As Object
    Set superTips = CreateObject("Scripting.Dictionary")
    
    If languageID = 1058 Then 'Ukrainian
        superTips.Add "FilterByText", "Зауважте: фільтр саме за текстом, а не за значенням. Тобто 0,5 і 0,50 — це різні критерії."
        superTips.Add "FilterByBg", "Фільтрувати за кольором виділеної клітинки"
        superTips.Add "FilterByFt", "Фільтрувати за кольором тексту виділеної клітинки"
        
    ElseIf languageID = 1049 Then 'Russian
        superTips.Add "FilterByText", "Обратите внимание: фильтр именно по тексту, а не по значению. То есть 0,5 и 0,50 — это разные критерии."
        superTips.Add "FilterByBg", "Фильтровать по цвету выделенной ячейки"
        superTips.Add "FilterByFt", "Фильтровать по цвету текста выделенной ячейки"
        
    Else 'English (default)
        superTips.Add "FilterByText", "Notice: filter by text, not value. Text 0.5 and 0.50 is a different criteria."
        superTips.Add "FilterByBg", "Filter by selected cell color"
        superTips.Add "FilterByFt", "Filter by selected text color"
    End If
    
    If superTips.Exists(control.ID) Then
        superTip = superTips(control.ID)
    Else
        superTip = "<Error>"
    End If
End Sub

'Main filtering function
Public Sub FilterBySel(ByRef control As Office.IRibbonControl)
    Dim ws As Worksheet, fltTable As Range, flt As Filter
    Dim fltColumn As Integer, fltCriteria As String, fltOperator As Variant
    
    On Error GoTo ErrorHandler 
    
    Set ws = ActiveSheet
    Set fltTable = ActiveCell.CurrentRegion
    
    'Determining filter type
    Select Case control.ID
        Case "FilterByText"
            fltCriteria = ActiveCell.Text
            fltOperator = Null
            
        Case "FilterByBg"
            fltCriteria = ActiveCell.DisplayFormat.Interior.Color
            fltOperator = IIf(ActiveCell.DisplayFormat.Interior.ColorIndex = -4142, 12, 8) 'Auto / Black
            
        Case "FilterByFt"
            fltCriteria = ActiveCell.DisplayFormat.Font.Color
            fltOperator = IIf(ActiveCell.DisplayFormat.Font.ColorIndex = -4105 Or ActiveCell.DisplayFormat.Font.ColorIndex = 1, 13, 9) 'Auto / Black
    End Select
    
    'Enabling AutoFilter
    If Not ws.AutoFilterMode Then
        Selection.AutoFilter
    End If
    
    'Define filter field
    fltColumn = ActiveCell.Column - ws.AutoFilter.Range.Column + 1
    
    'Applying AutoFilter
    If Not Intersect(ActiveCell, ws.AutoFilter.Range) Is Nothing Then
        Set flt = ws.AutoFilter.Filters(fltColumn)
        If Not flt.On Then
            fltTable.AutoFilter Field:=fltColumn, Criteria1:=fltCriteria, Operator:=fltOperator
        Else
            If (flt.Operator = 0 And control.ID = "FilterByText") Or _
                ((flt.Operator = 8 Or flt.Operator = 12) And control.ID = "FilterByBg") Or _
                ((flt.Operator = 9 Or flt.Operator = 13) And control.ID = "FilterByFt") Then
                fltTable.AutoFilter Field:=fltColumn 'Filter reset
            Else
                fltTable.AutoFilter Field:=fltColumn, Criteria1:=fltCriteria, Operator:=fltOperator
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
