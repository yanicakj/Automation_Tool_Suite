Attribute VB_Name = "Module3" 
Option Explicit

Sub FilePicker_Button1_Click()

    Call FilePicker(ws3, "E5", "J5")

End Sub

Sub Search(targetSheet As Worksheet, docRng As String, attachSpot As String)

    Dim sizeByRow As Long
    Dim lastRow As Long
    Dim MasterTracker As Workbook
    
    With targetSheet
        ' Checking to make sure the user inputted their information
        If Len(CStr(Trim(.Range(docRng).Value))) = 0 Then
            MsgBox "Please select a Document!"
            Exit Sub
        Else
            Set MasterTracker = AttachMasterTracker(targetSheet, CStr(Trim(.Range(docRng).Value)))
            If MasterTracker Is Nothing Then
                .Range(attachSpot).Value = "No File Attached"
                .Range(attachSpot).Font.Color = vbWhite
                .Range(attachSpot).Interior.Color = RGB(192, 0, 0)
                Exit Sub
            Else
                .Range(attachSpot).Value = "Tracker Attached"
                .Range(attachSpot).Interior.Color = vbGreen
                .Range(attachSpot).Font.Color = vbBlack
            End If
        End If
        
        lastRow = .Cells(.Rows.Count, "D").End(xlUp).row
        sizeByRow = lastRow - 9
        
        ' Clearing range
        If Not (targetSheet Is ws7) Then
            If lastRow > 9 Then
                .Range("E10:I" & lastRow).ClearContents
                .Range("E10:I" & lastRow).Interior.ColorIndex = 0
                
            End If
        End If
        
        ' Cleanup WRs
        CleanupWRs targetSheet, 10, "D"
        
        ' Removing Filters
        RemoveFilters MasterTracker
        
        ' Calling function per need
        If targetSheet Is ws3 Then
            Call MatchChecker(ws3, MasterTracker, lastRow, sizeByRow)
        ElseIf targetSheet Is ws5 Then
            Call ForecastChecker(ws5, MasterTracker, lastRow, sizeByRow)
        ElseIf targetSheet Is ws7 Then
            Call CompareForecastsAndActuals(ws7, MasterTracker, lastRow, sizeByRow)
        End If
    
    End With
    
    ' Closing sequence for this routine
    Application.StatusBar = ""
    ThisWorkbook.Activate

End Sub

Sub ClearAll_Button_Click()

    Call ClearRange(ws3, 10, "D", "I", "D")

End Sub

Sub ClearInfo_Button_Click()
    
    Call ClearRange(ws3, 10, "E", "I", "D")

End Sub

Sub Search_Button1_Click()

    Call Search(ws3, "E5", "J5")

End Sub

Sub MatchChecker(targetSheet As Worksheet, MasterTracker As Workbook, lastRow As Long, sizeByRow As Long)

    Dim currentWR As String
    Dim rgFound As Range
    Dim foundSheet As Worksheet
    Dim row As Long
    Dim wSheet As Variant

    With targetSheet
        ' Programmatically checking the master tracker for the user-inputted WRs
        For row = 10 To lastRow
            
            currentWR = Trim(CStr(.Range("D" & row).Value))
            Application.StatusBar = "Checking " & currentWR & ", " & CStr(Round((row - 10) / sizeByRow, 2) * 100) & "% Complete"
            While Left(currentWR, 5) <> "HBCBS" And (row - 10) < sizeByRow - 1
                row = row + 1
            Wend
            
            For Each wSheet In MasterTracker.Worksheets
                With wSheet.UsedRange
                    MasterTracker.Activate
                    Set rgFound = .Cells.Find(What:=currentWR, After:=.Range("A1"), LookIn:=xlValues, _
                        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                        MatchCase:=False, SearchFormat:=False)
                    ThisWorkbook.Activate
                End With
                
                If rgFound Is Nothing And Len(Trim(CStr(.Range("E" & row).Value))) = 0 Then
                    .Range("E" & row).Value = "Not on tracker"
                    .Range("E" & row).Interior.Color = RGB(255, 150, 150)
                    
                ElseIf Not (rgFound Is Nothing) And Len(Trim(CStr(.Range("F" & row).Value))) = 0 Then
                    Set foundSheet = MasterTracker.Worksheets(rgFound.Worksheet.Name)
                    .Range("E" & row).Value = "On Tracker"
                    .Range("F" & row).Value = rgFound.Worksheet.Name
                    .Range("E" & row).Interior.ColorIndex = 0
                    
                    If Len(Trim(CStr(foundSheet.Range("A" & rgFound.row).Value))) > 0 Then
                        .Range("G" & row).Value = foundSheet.Range("A" & rgFound.row).Value
                    Else
                        .Range("G" & row).Value = "No UATCOE Lead Found"
                    End If
                    
                    If Len(Trim(CStr(foundSheet.Range("E" & rgFound.row).Value))) > 0 Then
                        .Range("H" & row).Value = foundSheet.Range("E" & rgFound.row).Value
                    Else
                        .Range("H" & row).Value = "No UATCOE SME Found"
                    End If
                    .Range("I" & row).Value = foundSheet.Range("F" & rgFound.row).Value
                    
                ElseIf Not (rgFound Is Nothing) Then
                    Set foundSheet = MasterTracker.Worksheets(rgFound.Worksheet.Name)
                    .Range("F" & row).Value = .Range("F" & row).Value & vbCrLf & foundSheet.Name
                    .Cells(row, 6).Interior.Color = RGB(255, 255, 150)
                    
                    If Len(Trim(CStr(foundSheet.Range("A" & rgFound.row).Value))) > 0 Then
                        .Range("G" & row).Value = .Range("G" & row).Value & vbCrLf & foundSheet.Range("A" & rgFound.row).Value
                    Else
                        .Range("G" & row).Value = .Range("G" & row).Value & vbCrLf & "No UATCOE Lead Found"
                    End If
                    
                    If Len(Trim(CStr(foundSheet.Range("E" & rgFound.row).Value))) > 0 Then
                        .Range("H" & row).Value = .Range("H" & row).Value & vbCrLf & foundSheet.Range("E" & rgFound.row).Value
                    Else
                        .Range("H" & row).Value = .Range("H" & row).Value & vbCrLf & "No UATCOE SME Found"
                    End If
                    .Range("I" & row).Value = foundSheet.Range("F" & rgFound.row).Value
                    
                End If
            Next
            .Range("F" & row).HorizontalAlignment = xlLeft
            .Range("D" & row & ":E" & row).HorizontalAlignment = xlCenter
            .Range("G" & row & ":H" & row).HorizontalAlignment = xlCenter
            .Range("I" & row).HorizontalAlignment = xlLeft
            .Range("D" & row & ":I" & row).VerticalAlignment = xlCenter
        Next row
    End With

End Sub

