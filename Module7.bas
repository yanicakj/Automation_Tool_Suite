Attribute VB_Name = "Module7"
Option Explicit

Sub Run_Button5_Click()

    Call Search(ws7, "E5", "J5")

End Sub

Sub FilePicker_Button3_Click()

    Call FilePicker(ws7, "E5", "J5")

End Sub

Sub Clear_Button5_Click()

    Call ClearRange(ws7, 10, "D", "G", "D")

End Sub

Sub CompareForecastsAndActuals(targetSheet As Worksheet, ForecastTracker As Workbook, lastRow As Long, sizeByRow As Long)

    Dim rgFound     As Range
    Dim row         As Integer
    Dim currentWR   As String
    Dim monthTab    As Worksheet
    Dim lowerBound  As Integer
    Dim monthNum    As Integer
    Dim foundRow    As Integer
    Dim monthStart  As Integer
    Dim monthEnd    As Integer
    Dim localSpot   As Integer: localSpot = targetSheet.Cells(targetSheet.Rows.Count, "D").End(xlUp).row + 1
    Dim wSheet As Variant
    Dim comboBoxVal As String: comboBoxVal = targetSheet.OLEObjects("ComboBoxMWS7").Object.Text
    Dim currentYear As String
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim match As Object

    ' Checking to make sure the user inputted their information
    If comboBoxVal = "None" Then
        targetSheet.Range("E7").Activate
        MsgBox "Please select a month!"
        Exit Sub
    End If

    ' Setup
    For Each wSheet In ForecastTracker.Worksheets
        If InStr(wSheet.Name, Left(comboBoxVal, 3)) Then
            Set monthTab = wSheet
            monthNum = month("01 " & comboBoxVal & " 2018") + 2
            Exit For
        End If
    Next
    
    If monthTab Is Nothing Then
        MsgBox "Can't find tab on ForecastTracker for month : " & comboBoxVal
        Exit Sub
    End If
    
    ' getting year
    With regEx
        .Global = False
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "([0-9]{4})"
    End With
    
    If regEx.test(ForecastTracker.Name) Then
        Set match = regEx.Execute(ForecastTracker.Name)
        currentYear = match(0).Value
    End If
    
    With ForecastTracker.Worksheets(1)
        
        ' Find lower bound
        With .UsedRange
            ForecastTracker.Activate
            Set rgFound = .Cells.Find(What:=currentYear & " Project Forecast", After:=.Range("A1"), LookIn:=xlValues, _
                        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False)
            
            If rgFound Is Nothing Then
                MsgBox "Can't find cell on Forecast Tracker with value: '" & currentYear & " Project Forecast'." & _
                        "A cell with this value is needed on sheet 1 (used as a lower bound)!"
                Exit Sub
            Else
                lowerBound = rgFound.row - 3
                Set rgFound = Nothing
            End If
        End With
    
        ' Find month start bound - look into better indicator than "hours" & mergeCells
        On Error GoTo TabHandler
        With monthTab.UsedRange
            Set rgFound = .Cells.Find(What:="Hours", After:=.Range("A1"), LookIn:=xlValues, _
                        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False)
            
            If Not (rgFound Is Nothing) And rgFound.MergeCells Then
                monthStart = rgFound.row + 2
                Set rgFound = Nothing
            Else
                MsgBox "Can't find cell on Forecast Tracker with value: 'Hours'." & _
                        "A cell with this value is needed on sheet " & monthTab.Name & " (used as a start spot)!"
                Exit Sub
            End If
        End With
        
        ' coloring interior to nofill
        .Range(.Cells(3, monthNum), .Cells(lowerBound, monthNum)).Interior.ColorIndex = 0
    
        On Error GoTo 0
        ' Find lower month bound
        With monthTab
            monthEnd = .Cells(.Rows.Count, "A").End(xlUp).row
        End With
        
        ThisWorkbook.Activate
    
        ' Scanning WRs on first page for those on second
        For row = monthStart To monthEnd
            
            If Trim(Left(monthTab.Range("A" & row).Value, 5)) <> "HBCBS" Then
            
                monthTab.Range("G" & row).Interior.Color = RGB(255, 0, 0)
                UpdateTool targetSheet, _
                                monthTab.Range("A" & row).Value, _
                                "n/a", _
                                "n/a", _
                                "WR doesn't start with 'HBCBS' so macro skipped this one", _
                                localSpot, "red"
            
            Else
        
                With .UsedRange
                    Set rgFound = .Cells.Find(What:=Trim(CStr(monthTab.Range("A" & row).Value)), After:=.Range("A1"), LookIn:=xlValues, _
                                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                MatchCase:=False, SearchFormat:=False)
                    If rgFound Is Nothing Then
                        foundRow = -1
                    Else
                        foundRow = rgFound.row: Set rgFound = Nothing
                    End If
                    
                End With
        
                If foundRow = -1 Then ' WR not found on Forecast
                
                    ' if auto-insert is on
                    If targetSheet.OLEObjects("ComboBoxAI").Object.Text = "On" Then
                        
                        Call AutoInsertLineItem(ForecastTracker.Worksheets(1), monthTab, row, monthNum, lowerBound)
                        UpdateTool targetSheet, _
                                    monthTab.Range("A" & row).Value, _
                                    "Previously empty", _
                                    monthTab.Range("G" & row).Value, _
                                    "Auto-inserted line item, now " & monthTab.Range("G" & row).Value, _
                                    localSpot, "dark green"
                    
                    ' if auto-insert is off
                    Else
                    
                        monthTab.Range("G" & row).Interior.Color = RGB(255, 0, 0)
                        UpdateTool targetSheet, _
                                    monthTab.Range("A" & row).Value, _
                                    "n/a", _
                                    monthTab.Range("G" & row).Value, _
                                    "WR not found on Forecast Tab", _
                                    localSpot, "red"
                    
                    End If
                                    
                Else ' found WR on forecast
                
                    ' hours match
                    If Trim(CStr(.Cells(foundRow, monthNum).Value)) = Trim(CStr(monthTab.Range("G" & row).Value)) Then
                        .Cells(foundRow, monthNum).Interior.Color = RGB(0, 255, 0)
                        .Cells(foundRow, monthNum).Font.Color = RGB(0, 0, 0)
                        
                        monthTab.Range("G" & row).Interior.Color = RGB(0, 255, 0)
                        monthTab.Range("G" & row).Font.Color = RGB(0, 0, 0)
                        
                    ' hours empty on forecast
                    ElseIf Len(Trim(CStr(.Cells(foundRow, monthNum).Value))) = 0 Then
                        .Cells(foundRow, monthNum).Interior.Color = RGB(0, 190, 0)
                        .Cells(foundRow, monthNum).Font.Color = RGB(0, 0, 0)
                        .Cells(foundRow, monthNum).Value = monthTab.Range("G" & row).Value
                        
                        monthTab.Range("G" & row).Interior.Color = RGB(0, 190, 0)
                        monthTab.Range("G" & row).Font.Color = RGB(0, 0, 0)
                        
                        UpdateTool targetSheet, _
                                    monthTab.Range("A" & row).Value, _
                                    "Previously empty", _
                                    monthTab.Range("G" & row).Value, _
                                    "Auto-updated value on sheet to " & monthTab.Range("G" & row).Value, _
                                    localSpot, "dark green"
                        
                    ' if auto-update all is on
                    ElseIf targetSheet.OLEObjects("ComboBoxAU").Object.Text = "On" Then ' both have values & don't match
                    
                        .Cells(foundRow, monthNum).Interior.Color = RGB(0, 190, 0)
                        .Cells(foundRow, monthNum).Font.Color = RGB(0, 0, 0)
                        
                        monthTab.Range("G" & row).Interior.Color = RGB(0, 190, 0)
                        monthTab.Range("G" & row).Font.Color = RGB(0, 0, 0)
                        
                        UpdateTool targetSheet, _
                                    monthTab.Range("A" & row).Value, _
                                    "Previously " & .Cells(foundRow, monthNum).Value, _
                                    monthTab.Range("G" & row).Value, _
                                    "Auto-updated value on sheet to " & monthTab.Range("G" & row).Value, _
                                    localSpot, "dark green"
                                    
                        .Cells(foundRow, monthNum).Value = monthTab.Range("G" & row).Value
                    
                    Else ' both have values & don't match

                        If IsNumeric(monthTab.Range("G" & row).Value) And IsNumeric(.Cells(foundRow, monthNum).Value) Then
                                
                            ' auto-update value
                            If CInt(monthTab.Range("G" & row).Value) > CInt(.Cells(foundRow, monthNum).Value) Then
                                .Cells(foundRow, monthNum).Interior.Color = RGB(0, 190, 0)
                                .Cells(foundRow, monthNum).Font.Color = RGB(0, 0, 0)
                                
                                monthTab.Range("G" & row).Interior.Color = RGB(0, 190, 0)
                                monthTab.Range("G" & row).Font.Color = RGB(0, 0, 0)
                                
                                UpdateTool targetSheet, _
                                            monthTab.Range("A" & row).Value, _
                                            "Previously " & .Cells(foundRow, monthNum).Value, _
                                            monthTab.Range("G" & row).Value, _
                                            "Auto-updated value on sheet to " & monthTab.Range("G" & row).Value, _
                                            localSpot, "dark green"
                                            
                                .Cells(foundRow, monthNum).Value = monthTab.Range("G" & row).Value
                            
                            ' don't match - make yellow
                            Else
                                .Cells(foundRow, monthNum).Interior.Color = RGB(255, 255, 0)
                                .Cells(foundRow, monthNum).Font.Color = RGB(0, 0, 0)
                                
                                monthTab.Range("G" & row).Interior.Color = RGB(255, 255, 0)
                                monthTab.Range("G" & row).Font.Color = RGB(0, 0, 0)
                                
                                UpdateTool targetSheet, _
                                        monthTab.Range("A" & row).Value, _
                                        .Cells(foundRow, monthNum).Value, _
                                        monthTab.Range("G" & row).Value, _
                                        "Hours don't match. Forecasts are greater than actuals.", _
                                        localSpot, "yellow"
                            End If
                        
                        ' not numeric - can't compare - make yellow
                        Else
                        
                            .Cells(foundRow, monthNum).Interior.Color = RGB(255, 255, 0)
                            .Cells(foundRow, monthNum).Font.Color = RGB(0, 0, 0)
                            
                            monthTab.Range("G" & row).Interior.Color = RGB(255, 255, 0)
                            monthTab.Range("G" & row).Font.Color = RGB(0, 0, 0)
                            
                            UpdateTool targetSheet, _
                                        monthTab.Range("A" & row).Value, _
                                        .Cells(foundRow, monthNum).Value, _
                                        monthTab.Range("G" & row).Value, _
                                        "Hours don't match. One of the hours values is not numerical & can't be compared", _
                                        localSpot, "yellow"
                            
                        End If
                    End If
                End If
            End If
        Next
        
        ' Looping thru forecast tab to check missed line items
        For row = 3 To lowerBound
        
            .Cells(row, monthNum).Value = Replace(CStr(Trim(.Cells(row, monthNum).Value)), Chr(160), "")
            .Cells(row, monthNum).Value = Replace(CStr(Trim(.Cells(row, monthNum).Value)), " ", "")
            .Cells(row, monthNum).Value = Replace(CStr(Trim(.Cells(row, monthNum).Value)), vbCrLf, "")
            .Cells(row, monthNum).Value = Replace(CStr(Trim(.Cells(row, monthNum).Value)), Chr(10), "")
            .Cells(row, monthNum).Value = Replace(CStr(Trim(.Cells(row, monthNum).Value)), Chr(13), "")
        
            If .Cells(row, monthNum).Interior.Color <> RGB(255, 0, 0) And _
                .Cells(row, monthNum).Interior.Color <> RGB(255, 255, 0) And _
                .Cells(row, monthNum).Interior.Color <> RGB(0, 255, 0) And _
                .Cells(row, monthNum).Interior.Color <> RGB(0, 190, 0) And _
                Len(Trim(CStr(.Cells(row, monthNum).Value))) > 0 Then
                
                If Trim(CStr(.Cells(row, monthNum).Value)) = "0" Then
                
                    .Cells(row, monthNum).Interior.Color = RGB(0, 255, 0)
                    .Cells(row, monthNum).Font.Color = RGB(0, 0, 0)
                
                Else
                
                    .Cells(row, monthNum).Interior.Color = RGB(255, 0, 0)
                    .Cells(row, monthNum).Font.Color = RGB(0, 0, 0)
                    
                    UpdateTool targetSheet, _
                                .Cells(row, "A").Value, _
                                .Cells(row, monthNum).Value, _
                                "n/a", _
                                "WR not on Actuals tab", _
                                localSpot, "red"
                
                End If
            End If
        Next
        
    End With
    
    ' updating formulas
    Call UpdateFormulas(ForecastTracker.Worksheets(1), lowerBound + 1)

    ' Closing sequence for this routine
    Application.StatusBar = ""
    ThisWorkbook.Activate
    
    MsgBox "Done!"
    
    Exit Sub

TabHandler:
    Application.StatusBar = ""
    ThisWorkbook.Activate
    MsgBox "Please make sure the name of the tab is spelled correctly!"

End Sub

Sub alignRow(targetSheet As Worksheet, row As Integer)

    With targetSheet
    
        .Range("D" & row & ":F" & row).HorizontalAlignment = xlCenter
        .Range("G" & row).HorizontalAlignment = xlLeft
        .Range("D" & row & ":F" & row).VerticalAlignment = xlCenter
    
    End With

End Sub

Sub UpdateTool(targetSheet As Worksheet, dValue As String, eValue As String, fValue As String, gValue As String, ByRef localSpot, cellColor As String)

    With targetSheet
    
        Select Case cellColor
            Case "dark green"
                .Range("D" & localSpot).Interior.Color = RGB(0, 190, 0)
            Case "yellow"
                .Range("D" & localSpot).Interior.Color = RGB(255, 255, 0)
            Case "red"
                .Range("D" & localSpot).Interior.Color = RGB(255, 0, 0)
        End Select
    
        .Range("D" & localSpot).Value = dValue
        .Range("E" & localSpot).Value = eValue
        .Range("F" & localSpot).Value = fValue
        .Range("G" & localSpot).Value = gValue
        
        Call alignRow(targetSheet, CInt(localSpot))
        
        localSpot = localSpot + 1
        
    End With
    
End Sub

Public Sub AutoInsertLineItem(forecastSheet As Worksheet, monthTab As Worksheet, row As Integer, monthNum As Integer, ByRef lowerBound As Integer)

    With forecastSheet
    
        .Rows(lowerBound + 1).EntireRow.Insert
        lowerBound = lowerBound + 1
            
        .Range("A" & lowerBound).Value = monthTab.Range("A" & row).Value
        .Range("B" & lowerBound).Value = monthTab.Range("B" & row).Value
        .Cells(lowerBound, monthNum).Value = monthTab.Range("G" & row).Value
        .Range("O" & lowerBound).Value = "=SUM(C" & lowerBound & ":N" & lowerBound & ")"
        
        .Range("A" & lowerBound & ":B" & lowerBound).Interior.Color = RGB(217, 217, 217)
        .Range("C" & lowerBound & ":O" & lowerBound).Interior.ColorIndex = 0
        .Range("A" & lowerBound & ":O" & lowerBound).Borders.LineStyle = xlContinuous
        .Range("O" & lowerBound & ":O" & (lowerBound + 1)).Borders(xlEdgeRight).Weight = .Range("O" & (lowerBound - 3)).Borders(xlEdgeRight).Weight
        
        monthTab.Range("G" & row).Interior.Color = RGB(0, 190, 0)
        .Cells(lowerBound, monthNum).Interior.Color = RGB(0, 190, 0)
        
    End With

End Sub

Public Sub UpdateFormulas(forecastSheet As Worksheet, bottomRow As Integer)

    Dim colArray() As Variant: colArray = Array("C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O")
    Dim col As Variant

    With forecastSheet
        For Each col In colArray
            .Range(col & bottomRow).Value = "=SUM(" & col & "3" & ":" & col & (bottomRow - 1) & ")"
        Next
        .Rows(bottomRow).Calculate
    End With

End Sub




