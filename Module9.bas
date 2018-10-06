Attribute VB_Name = "Module9"
Option Explicit

Sub FilePicker_Button4_Click()

    Call FilePicker(ws9, "C5")
    
    If BookOpen(returnFileName(ws9.Range("C5").Value, "\")) Then
        ws9.Range("C5").Interior.Color = vbGreen
    End If

End Sub

Sub FilePicker_Button5_Click()

    Call FilePicker(ws9, "C7")

    If BookOpen(returnFileName(ws9.Range("C7").Value, "\")) Then
        ws9.Range("C7").Interior.Color = vbGreen
    End If

End Sub

Sub Run_Button6_Click()

    Call ScrubTrackers(ws9, "C5", "C7")

End Sub

Sub ScrubTrackers(targetSheet As Worksheet, docRng As String, docRng2 As String)

    Dim sizeByRow As Long
    Dim lastRow As Long
    Dim ForecastTracker As Workbook
    Dim GenpactPipeline As Workbook
    Dim trackerArray() As Variant: trackerArray = Array(ForecastTracker, GenpactPipeline)
    Dim docArray() As Variant: docArray = Array(docRng, docRng2)
    Dim docSpot As Variant
    Dim monthNumber As Integer
    Dim yearValue As String
    
    With targetSheet
    
        ' Checking to make sure the user inputted their information
        For docSpot = LBound(docArray) To UBound(docArray)
            If Len(CStr(Trim(.Range(docArray(docSpot)).Value))) = 0 Then
                .Range(docArray(docSpot)).Interior.ColorIndex = 0
                .Range(docArray(docSpot)).Activate
                MsgBox "Please select a Tracker!"
                Exit Sub
            Else
                Set trackerArray(docSpot) = AttachMasterTracker(targetSheet, CStr(Trim(.Range(docArray(docSpot)).Value)))
                If trackerArray(docSpot) Is Nothing Then
                    .Range(docArray(docSpot)).Interior.ColorIndex = 0
                    Exit Sub
                Else
                    .Range(docArray(docSpot)).Interior.Color = vbGreen
                End If
            End If
        Next
        
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        sizeByRow = lastRow - 9
        
        ' Resetting workbooks
        Set ForecastTracker = trackerArray(0)
        Set GenpactPipeline = trackerArray(1)
        
        ' Removing Filters
        RemoveFilters ForecastTracker
        RemoveFilters GenpactPipeline
        
        ' checking combobox for month
        If .OLEObjects("ComboBoxMonth").Object.Text = "Please select" Then
            .Range("F2").Activate
            MsgBox "Please select a month!"
            Exit Sub
        Else
            monthNumber = month("01 " & .OLEObjects("ComboBoxMonth").Object.Text & " 2018")
            yearValue = .Range("M2").Value
        End If
        
        If IsNumeric(yearValue) = False Then
            .Range("M2").Activate
            MsgBox "Please enter a valid value for the year!"
            Exit Sub
        ElseIf yearValue < 2016 Then
            .Range("M2").Activate
            MsgBox "Please enter a year greater than or equal to 2016!"
            Exit Sub
        End If
    
        ' Starting mission control
        ProgramDriver ForecastTracker, GenpactPipeline, .OLEObjects("ComboBoxMonth").Object.Text, monthNumber, yearValue
    
    End With
    
    ' Closing sequence for this routine
    Application.StatusBar = ""
    MsgBox "Done!"
    ThisWorkbook.Activate

End Sub

Sub ProgramDriver(ManiTracker As Workbook, BharathTracker As Workbook, startMonth As String, monthNum As Integer, yearValue As String)

    Dim row As Integer
    Dim col As Integer
    Dim lastBharathRow As Integer
    Dim lastToolRow As Integer
    Dim workRequest As String
    Dim lowerBound As Integer
    Dim rgFound As Range
    Dim foundRow As Integer
    Dim mLetter As String
    Dim bLetter As String
    Dim mStartCol As Integer
    Dim bStartCol As Integer
    Dim pipeCol As Integer
    Dim marked As Boolean: marked = False
    Dim pipeSheet As Worksheet: Set pipeSheet = BharathTracker.Worksheets(1)
    Dim maniSheet As Worksheet: Set maniSheet = ManiTracker.Worksheets(1)
    Dim nonZero As Boolean: nonZero = False
    Dim mVal As String
    Dim bVal As String
    Dim inc As Integer
    Dim mRng As Range
    Dim curMnth As String
    Dim aiString As String
    Dim rV As Integer, gV As Integer, bV As Integer
    
    Application.StatusBar = "Checking Work Requests"
    
    ' Finding last rows
    With pipeSheet
        lastBharathRow = .Cells(.Rows.Count, "B").End(xlUp).row
        lastToolRow = ws9.Cells(ws9.Rows.Count, "B").End(xlUp).row + 1
    End With
    
    ' Finding lower bound on Mani's Tracker - might be able to delete this
    With maniSheet.UsedRange
        Set rgFound = .Cells.Find(What:="Project Forecast $", After:=.Range("A1"), LookIn:=xlValues, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
        If rgFound Is Nothing Then
            MsgBox "Couldn't find cell on Forcast sheet with value 'Project Forecast $' to use as lower bound!"
            Exit Sub
        Else
            lowerBound = rgFound.row - 3
        End If
    End With
    
    ' setting vars
    inc = SetCellVariables(mLetter, bLetter, mStartCol, bStartCol, monthNum, yearValue)
    
    ' Removing color fills
    maniSheet.Range("C3:O" & lowerBound).Interior.ColorIndex = 0
    pipeSheet.Range(pipeSheet.Cells(2, bStartCol), pipeSheet.Cells(lastBharathRow, (bStartCol + 12))).Interior.ColorIndex = 0

    ' setup progress format
    progressBar ws9, "setup"

    With pipeSheet
    
        For row = 2 To lastBharathRow
        
            ' update status
            Application.StatusBar = CStr(Round((row - 2) / (lastBharathRow - 2), 2) * 100) & "% Complete"
            CheckProgress ws9, lastBharathRow, row
            
            If Left(Trim(CStr(.Range("B" & row).Value)), 5) = "HBCBS" Then
            
                workRequest = Left(Trim(CStr(.Range("B" & row).Value)), 13)
                
                ' pipeline has no values for year
                If WorksheetFunction.CountA(.Range(.Cells(row, bStartCol), .Cells(row, bStartCol + (12 - monthNum)))) = 0 Then
                
                    Set rgFound = findRange(workRequest, maniSheet)
                    ' WR is on mani's sheet
                    If Not (rgFound Is Nothing) Then
                        
                        foundRow = rgFound.row: Set rgFound = Nothing
                    
                        With maniSheet
                            ' WR has hours on mani's
                            If WorksheetFunction.CountA(.Range(mLetter & foundRow & ":N" & foundRow)) > 0 Then
                            
                                For col = mStartCol To 14
                                    ' if bharath is empty and Mani is not empty
                                    If Len(Trim(CStr(.Cells(foundRow, col).Value))) > 0 And Trim(CStr(.Cells(foundRow, col).Value)) <> "0" Then
                                        .Cells(foundRow, col).Interior.Color = RGB(255, 0, 0)
                                    End If
                                Next
                            
                            End If
                        End With
                    End If
                
                ' pipeline has values for this year
                Else
                
                    Set rgFound = findRange(workRequest, maniSheet)
                    
                    ' not on mani's forecast
                    If rgFound Is Nothing Then
                        
                        For col = bStartCol To bStartCol + 12 - monthNum
                            ' ensuring range has other numbers than 0's
                            If IsNumeric(.Cells(row, col).Value) Then
                                If .Cells(row, col).Value > 0 Then
                                    If nonZero = False Then nonZero = True
                                    .Cells(row, col).Interior.Color = RGB(255, 50, 50)
                                    ws9.Cells(lastToolRow, col - inc).Value = .Cells(row, col).Value
                                End If
                            End If
                        Next
                        
                        ' If atleast one value was not zero
                        If nonZero = True Then
                        
                            ' Add auto-insert line item here
                            If ws9.OLEObjects("ComboBoxAILI").Object.Text = "On" Then
                                
                                ' insert WR
                                Call AutoInsertLineItem2(maniSheet, pipeSheet, row, lowerBound)
                                
                                For col = bStartCol To bStartCol + 12 - monthNum
                                    ' ensuring range has other numbers than 0's
                                    If IsNumeric(.Cells(row, col).Value) Then
                                        If .Cells(row, col).Value > 0 Then
                                            .Cells(row, col).Interior.Color = RGB(255, 50, 50)
                                            maniSheet.Cells(lowerBound, col - inc - 1).Value = .Cells(row, col).Value
                                            maniSheet.Cells(lowerBound, col - inc - 1).Interior.Color = RGB(0, 190, 0)
                                        End If
                                    End If
                                Next
                                
                                aiString = "Auto-inserted onto your tracker"
                                rV = 110: gV = 255: bV = 110
                            
                            Else
                                aiString = "Missing from your tracker"
                                rV = 255: gV = 255: bV = 110
                            End If
                        
                            With ws9
                                .Range("B" & lastToolRow).Value = workRequest
                                .Range("C" & lastToolRow).Value = aiString
                                .Range("C" & lastToolRow).Interior.Color = RGB(rV, gV, bV)
                                lastToolRow = lastToolRow + 1
                                nonZero = False
                            End With
                        End If
                    
                    ' is on mani's forecast
                    Else
                    
                        foundRow = rgFound.row: Set rgFound = Nothing
                        pipeCol = bStartCol
                        
                        For col = mStartCol To 14
                            
                            mVal = Trim(CStr(maniSheet.Cells(foundRow, col).Value))
                            bVal = Trim(CStr(.Cells(row, pipeCol).Value))
                            curMnth = MonthName(col - 2, True)
                            
                            ' if equal
                            If mVal = bVal Then
                            
                                If Len(mVal) > 0 Then
                                    maniSheet.Cells(foundRow, col).Interior.Color = RGB(0, 255, 0)
                                    .Cells(row, pipeCol).Interior.Color = RGB(0, 255, 0)
                                End If
                                
                            ' not equal
                            Else
                            
                                maniSheet.Cells(foundRow, col).Interior.Color = RGB(255, 255, 0)
                                .Cells(row, pipeCol).Interior.Color = RGB(255, 255, 0)
                                
                                ' both have values
                                If Len(mVal) > 0 And Len(bVal) > 0 Then
                                
                                    If ws9.OLEObjects("ComboBox2FC").Object.Text = "All" Then
                                        
                                        maniSheet.Cells(foundRow, col).Value = bVal
                                        maniSheet.Cells(foundRow, col).Interior.Color = RGB(0, 190, 0)
                                        .Cells(row, pipeCol).Interior.Color = RGB(0, 190, 0)
                                        
                                        With ws9
                                            .Range("B" & lastToolRow).Value = workRequest
                                            .Range("C" & lastToolRow).Value = "Auto-updated " & curMnth & " from " & mVal & " to " & bVal
                                            lastToolRow = lastToolRow + 1
                                        End With
                                        
                                    ElseIf ws9.OLEObjects("ComboBox2FC").Object.Text = "Bharath's > Yours" Then
                                        
                                        If IsNumeric(bVal) And IsNumeric(mVal) Then
                                        
                                            If bVal > mVal Then
                                                maniSheet.Cells(foundRow, col).Value = bVal
                                                maniSheet.Cells(foundRow, col).Interior.Color = RGB(0, 190, 0)
                                                .Cells(row, pipeCol).Interior.Color = RGB(0, 190, 0)
                                                
                                                With ws9
                                                    .Range("B" & lastToolRow).Value = workRequest
                                                    .Range("C" & lastToolRow).Value = "Auto-updated " & curMnth & " from " & mVal & " to " & bVal
                                                    lastToolRow = lastToolRow + 1
                                                End With
                                                
                                            End If
                                            
                                        End If
                                        
                                    End If
                                    
                                ' only mani has value
                                ElseIf Len(mVal) > 0 Then
                                
                                    If mVal = "0" Then
                                        maniSheet.Cells(foundRow, col).Interior.ColorIndex = 0
                                        .Cells(row, pipeCol).Interior.ColorIndex = 0
                                    End If
                                
                                ' only pipeline has value
                                Else
                                
                                    If bVal = "0" Then
                                        maniSheet.Cells(foundRow, col).Interior.ColorIndex = 0
                                        .Cells(row, pipeCol).Interior.ColorIndex = 0
                                    Else
                                    
                                        If ws9.OLEObjects("ComboBox2FC").Object.Text = "All" Then
                                        
                                            maniSheet.Cells(foundRow, col).Value = bVal
                                            maniSheet.Cells(foundRow, col).Interior.Color = RGB(0, 190, 0)
                                            .Cells(row, pipeCol).Interior.Color = RGB(0, 190, 0)
                                            
                                            With ws9
                                                .Range("B" & lastToolRow).Value = workRequest
                                                .Range("C" & lastToolRow).Value = "Auto-updated " & curMnth & " from *blank* to " & bVal
                                                lastToolRow = lastToolRow + 1
                                            End With
                                            
                                        ElseIf ws9.OLEObjects("ComboBox2FC").Object.Text = "Bharath's > Yours" Then
                                        
                                            If IsNumeric(bVal) Then
                                            
                                                If bVal > 0 Then
                                                
                                                    maniSheet.Cells(foundRow, col).Value = bVal
                                                    maniSheet.Cells(foundRow, col).Interior.Color = RGB(0, 190, 0)
                                                    .Cells(row, pipeCol).Interior.Color = RGB(0, 190, 0)
                                                    
                                                    With ws9
                                                        .Range("B" & lastToolRow).Value = workRequest
                                                        .Range("C" & lastToolRow).Value = "Auto-updated " & curMnth & " from *blank* to " & bVal
                                                        lastToolRow = lastToolRow + 1
                                                    End With
                                                    
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        
                            pipeCol = pipeCol + 1
                        
                        Next
                    End If
                End If
            
            Else
                
                workRequest = Left(Trim(CStr(.Range("B" & row).Value)), 13)
                
                ' pipeline has no values for year
                If WorksheetFunction.CountA(.Range(.Cells(row, bStartCol), .Cells(row, bStartCol + (12 - monthNum)))) > 0 Then
                
                    For col = bStartCol To bStartCol + 12
                        ' ensuring range has other numbers than 0's
                        If IsNumeric(.Cells(row, col).Value) Then
                            If .Cells(row, col).Value > 0 Then
                                If nonZero = False Then nonZero = True
                                .Cells(row, col).Interior.Color = RGB(255, 50, 50)
                                ws9.Cells(lastToolRow, col - inc).Value = .Cells(row, col).Value
                            End If
                        End If
                    Next
                    
                    ' If atleast one value was not zero
                    If nonZero = True Then
                        With ws9
                            .Range("B" & lastToolRow).Value = workRequest & " from row " & row & " of pipeline"
                            .Range("B" & lastToolRow).WrapText = True
                            .Range("C" & lastToolRow).Value = "Not normal WR, no 'HBCBS'"
                            .Range("C" & lastToolRow).WrapText = True
                            .Range("C" & lastToolRow).Interior.Color = RGB(255, 200, 200)
                            lastToolRow = lastToolRow + 1
                            nonZero = False
                        End With
                    End If
                
                End If
                
            End If
        Next
    End With
    
    ' Checking for missings on Bharath's tracker
    With maniSheet
        For row = 3 To lowerBound
            For col = mStartCol To 14
            
                Set mRng = .Cells(row, col)
            
                ' if not colored and not empty
                If Len(Trim(CStr(mRng.Value))) > 0 And (mRng.Interior.Color = RGB(255, 255, 255) Or mRng.Interior.ColorIndex = 0) Then
                    
                    ' Note on Tool
                    mRng.Interior.Color = RGB(255, 50, 50)
                    ws9.Cells(lastToolRow, col + 1).Value = mRng.Value
                    If marked = False Then marked = True
                    
                End If
            Next
            
            If marked = True Then
                ws9.Range("B" & lastToolRow).Value = .Cells(row, 1).Value
                ws9.Range("C" & lastToolRow).Value = "Not on Bharath's Pipeline"
                ws9.Range("C" & lastToolRow).Interior.Color = RGB(200, 255, 0)
                lastToolRow = lastToolRow + 1
                marked = False
            End If
            
        Next
    End With
    
    ' update formulas - func from mod 7
    Call UpdateFormulas(maniSheet, lowerBound + 1)
    
    ' clean green progress bar
    progressBar ws9, "cleanup"
    
End Sub

Sub ClearForecastInfo()

    Dim lastRow As Integer
    
    With ws9

        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row + 1
        
        If lastRow > 9 Then
            .Range("A10:O" & lastRow).ClearContents
            .Range("A10:O" & lastRow).Interior.ColorIndex = 0
        End If
    
    End With
    
End Sub

Function SetCellVariables(ByRef mLetter As String, ByRef bLetter As String, ByRef mStartCol As Integer, ByRef bStartCol As Integer, monthNum As Integer, yearValue As String) As Integer

    Dim inc As Integer
    
    inc = CInt(yearValue) - 2016

    mStartCol = monthNum + 2
    bStartCol = (14 + monthNum) + (inc * 12)
    mLetter = Split(ws9.Cells(1, mStartCol).Address, "$")(1)
    bLetter = Split(ws9.Cells(1, bStartCol).Address, "$")(1)
    
    SetCellVariables = 11 + (12 * inc)

End Function

Function findRange(workRequest As String, lookupSheet As Worksheet) As Range

    With lookupSheet.UsedRange
        
        Set findRange = .Cells.Find(What:=workRequest, After:=.Range("A1"), LookIn:=xlValues, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
        
    End With

End Function

Public Sub AutoInsertLineItem2(maniSheet As Worksheet, pipeSheet As Worksheet, row As Integer, ByRef lowerBound As Integer)

    With maniSheet
    
        .Rows(lowerBound + 1).EntireRow.Insert
        lowerBound = lowerBound + 1
            
        .Range("A" & lowerBound).Value = pipeSheet.Range("B" & row).Value
        .Range("B" & lowerBound).Value = pipeSheet.Range("C" & row).Value
        '.Cells(lowerBound, monthNum).Value = pipeSheet.Range("G" & row).Value
        .Range("O" & lowerBound).Value = "=SUM(C" & lowerBound & ":N" & lowerBound & ")"
        
        .Range("A" & lowerBound & ":B" & lowerBound).Interior.Color = RGB(217, 217, 217)
        .Range("C" & lowerBound & ":O" & lowerBound).Interior.ColorIndex = 0
        .Range("A" & lowerBound & ":O" & lowerBound).Borders.LineStyle = xlContinuous
        .Range("O" & lowerBound & ":O" & (lowerBound + 1)).Borders(xlEdgeRight).Weight = .Range("O" & (lowerBound - 3)).Borders(xlEdgeRight).Weight
        
        'pipeSheet.Range("G" & row).Interior.Color = RGB(0, 190, 0)
        '.Cells(lowerBound, monthNum).Interior.Color = RGB(0, 190, 0)
        
    End With

End Sub


