Attribute VB_Name = "Utils" 
Option Explicit

Public Sub CleanupWRs(targetSheet As Worksheet, startRow As Integer, col As String)

    Dim lastRow As Long
    Dim row As Long
    
    With targetSheet
        lastRow = .Cells(.Rows.Count, col).End(xlUp).row
        
        If lastRow <= startRow Then Exit Sub
        
        For row = startRow To lastRow
            .Range(col & row).Value = Replace(CStr(Trim(.Range(col & row).Value)), Chr(160), "")
            .Range(col & row).Value = Replace(CStr(Trim(.Range(col & row).Value)), " ", "")
            .Range(col & row).Value = Replace(CStr(Trim(.Range(col & row).Value)), vbCrLf, "")
            .Range(col & row).Value = Replace(CStr(Trim(.Range(col & row).Value)), Chr(10), "")
            .Range(col & row).Value = Replace(CStr(Trim(.Range(col & row).Value)), Chr(13), "")
            .Range(col & row).HorizontalAlignment = xlCenter
            .Range(col & row).VerticalAlignment = xlCenter
            .Range(col & row).Font.Name = "Calibri"
            .Range(col & row).Font.Size = 11
        Next
    End With

End Sub

Public Function BookOpen(strBookName As String) As Boolean

    Dim oBk As Workbook
    On Error Resume Next
    Set oBk = Workbooks(strBookName)
    On Error GoTo 0
    If oBk Is Nothing Then
        BookOpen = False
    Else
        BookOpen = True
    End If

End Function

Public Sub RemoveFilters(targetWorkbook As Workbook)

    Dim ws As Variant
    
    For Each ws In targetWorkbook.Worksheets
    
        ' Remove filters
        If ws.FilterMode Then ws.ShowAllData
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    Next

End Sub

Public Sub FilePicker(targetSheet As Worksheet, cellRange As String, Optional attachSpot As String)
 
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
 
        If .SelectedItems.Count > 0 Then
           targetSheet.Range(cellRange).Value = .SelectedItems(1)
        End If
        
    End With
    
    With targetSheet
    
        If Not (targetSheet Is ws9) Then
    
            If BookOpen(returnFileName(.Range(cellRange).Value, "\")) Then
                .Range(attachSpot).Value = "Tracker Attached"
                .Range(attachSpot).Interior.Color = vbGreen
                .Range(attachSpot).Font.Color = vbBlack
            Else
                .Range(attachSpot).Value = "No File Attached"
                .Range(attachSpot).Font.Color = vbWhite
                .Range(attachSpot).Interior.Color = RGB(192, 0, 0)
            End If
            
        End If
        
    End With

End Sub

Public Function AttachMasterTracker(targetSheet As Worksheet, cellValue As String) As Workbook

    Dim username    As String
    Dim nameOfDoc   As String
    Dim check       As Boolean
    Dim nameLen     As Integer
    Dim wbName      As String
    
    ' Grabbing user information
    With targetSheet

        wbName = returnFileName(cellValue, "\")
        
        ' First checking if open
        If BookOpen(wbName) Then
            Set AttachMasterTracker = Workbooks(wbName)
        Else
            On Error GoTo FileOpener1
            Application.DisplayAlerts = False
            If InStr(cellValue, "\") Then
                Workbooks.Open Filename:=CStr(cellValue), UpdateLinks:=3
            Else
                Workbooks.Open Filename:="C:\Users\" & Environ("username") & "\Desktop\" & wbName, UpdateLinks:=3
            End If
            On Error GoTo 0
            Application.DisplayAlerts = True
            ActiveWindow.Visible = True
            ThisWorkbook.Activate
            Set AttachMasterTracker = Workbooks(wbName)
        End If
    
    End With
    
    Exit Function
    
FileOpener1:
    MsgBox "Unable to open file at path " & cellValue
    Application.StatusBar = ""
    Application.ScreenUpdating = True
    
End Function

Public Function returnFileName(filePath As String, delim As String) As String

    Dim words() As String

    If InStr(filePath, delim) Then
        words = Split(filePath, delim)
        returnFileName = Trim(CStr(words(UBound(words) - LBound(words))))
    Else
        returnFileName = Trim(CStr(filePath))
    End If
    
End Function

Sub ClearRange(targetSheet As Worksheet, startRow As Integer, startCol As String, endCol As String, wrCol As String)

    Dim lastRow As Long

    With targetSheet
    
        lastRow = .Cells(.Rows.Count, wrCol).End(xlUp).row
        
        If lastRow < startRow Then Exit Sub
        
        .Range(startCol & startRow & ":" & endCol & lastRow).ClearContents
        .Range(startCol & startRow & ":" & endCol & lastRow).Interior.ColorIndex = 0
        .Range(startCol & startRow & ":" & endCol & lastRow).UnMerge
        .Range(startCol & startRow & ":" & endCol & lastRow).Font.Bold = False
        .Range(startCol & startRow & ":" & endCol & lastRow).Font.Size = 11
        .Range(startCol & startRow & ":" & endCol & lastRow).Font.Name = "Calibri"
    End With

End Sub

Public Sub progressBar(targetSheet As Worksheet, indicator As String)

    With targetSheet
    
        If indicator = "setup" Then
        
            .Range("A10").Value = "Done!"
            .Range("A20").Value = "Start"
            .Range("A10").Font.Color = vbWhite
            .Range("A10").Font.Bold = True
            .Range("A20").Font.Color = vbWhite
            .Range("A20").Font.Bold = True
            .Range("A10:A20").Interior.Color = vbBlack
            .Range("A10").HorizontalAlignment = xlCenter
            .Range("A20").HorizontalAlignment = xlCenter
            
        ElseIf indicator = "cleanup" Then
        
            .Range("A10:A20").Interior.ColorIndex = 0
            .Range("A10:A20").ClearContents
        
        End If
    
    End With
    
End Sub

Public Sub CheckProgress(targetSheet As Worksheet, lastLocalRow As Integer, row As Integer)

    ' Checking status
    With targetSheet
        If Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 10 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 20 Then
            MakeGreen targetSheet, "A19"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 20 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 30 Then
            MakeGreen targetSheet, "A18"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 30 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 40 Then
            MakeGreen targetSheet, "A17"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 40 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 50 Then
            MakeGreen targetSheet, "A16"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 50 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 60 Then
            MakeGreen targetSheet, "A15"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 60 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 70 Then
            MakeGreen targetSheet, "A14"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 70 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 80 Then
            MakeGreen targetSheet, "A13"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 80 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 90 Then
            MakeGreen targetSheet, "A12"
        ElseIf Round((row - 8) / (lastLocalRow - 7), 2) * 100 >= 90 And Round((row - 8) / (lastLocalRow - 7), 2) * 100 < 100 Then
            MakeGreen targetSheet, "A11"
        End If
    End With

End Sub

Sub MakeGreen(targetSheet, rangeString As String)

    With targetSheet
        ' removed screenupdating lines
        .Range(rangeString).Interior.Color = RGB(0, 200, 0)
    End With

End Sub

