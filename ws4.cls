VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ws4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

    Dim lastRow As Long
    
    lastRow = ws4.Cells(ws4.Rows.Count, "C").End(xlUp).row
    
    If lastRow = 9 Then
        ws4.Range("A1").Activate
        ws4.Range("C7").Activate
    End If

End Sub

