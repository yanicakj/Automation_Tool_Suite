VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ws9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

    Dim lastRow As Long
    Dim cellInfo1 As String
    Dim cellInfo2 As String
    
    With ws9
    
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        cellInfo1 = Trim(CStr(.Range("C5").Value))
        cellInfo2 = Trim(CStr(.Range("C7").Value))
        
        If lastRow = 9 Then
            .Range("A1").Activate
            .Range("C5").Activate
        End If
    
        If BookOpen(returnFileName(cellInfo1, "\")) Then
            .Range("C5").Interior.Color = vbGreen
            .Range("C5").Font.Color = vbBlack
        Else
            .Range("C5").Font.Color = vbBlack
            .Range("C5").Interior.ColorIndex = 0
        End If
        
        If BookOpen(returnFileName(cellInfo2, "\")) Then
            .Range("C7").Interior.Color = vbGreen
            .Range("C7").Font.Color = vbBlack
        Else
            .Range("C7").Font.Color = vbBlack
            .Range("C7").Interior.ColorIndex = 0
        End If
    End With

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    Dim cellInfo1 As String
    Dim cellInfo2 As String
    
    With ws9
    
        cellInfo1 = Trim(CStr(.Range("C5").Value))
        cellInfo2 = Trim(CStr(.Range("C7").Value))
        
        If BookOpen(returnFileName(cellInfo1, "\")) Then
            .Range("C5").Interior.Color = vbGreen
            .Range("C5").Font.Color = vbBlack
        Else
            .Range("C5").Font.Color = vbBlack
            .Range("C5").Interior.ColorIndex = 0
        End If
        
        If BookOpen(returnFileName(cellInfo2, "\")) Then
            .Range("C7").Interior.Color = vbGreen
            .Range("C7").Font.Color = vbBlack
        Else
            .Range("C7").Font.Color = vbBlack
            .Range("C7").Interior.ColorIndex = 0
        End If

    End With

End Sub


