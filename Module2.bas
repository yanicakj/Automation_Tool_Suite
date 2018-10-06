Attribute VB_Name = "Module2"
Option Explicit

Sub Run_Button2_Click()

    Call CQscrub(ws2, "A")

End Sub

Sub Looper2(counter As Integer, sizeByRow As Integer, ie As InternetExplorer, searchStringObj As Object, ObjElement As Object, jumpBack As Integer, targetSheet As Worksheet, wrCol As String)

    Dim identification As String
    Dim leadObj As Object
    Dim fundingObj As Object
    Dim closeObj As Object
    Dim leadCol As String: leadCol = "B"
    Dim fundCol As String: fundCol = "C"

    On Error Resume Next
    
    With targetSheet
    
        Application.StatusBar = "Checking " & CStr(.Range(wrCol & CStr(counter + 10)).Value) & ", " & CStr(Round(counter / sizeByRow, 2) * 100) & "% Complete"
        searchStringObj.Value = CStr(.Range(wrCol & CStr(counter + 10)).Value)
        ObjElement.Click
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        Application.Wait (Now + TimeValue("0:00:06"))
        
        ' Finding the IT Lead
        identification = "..." & CStr(9 + ((counter - jumpBack) * 23))
        Do While leadObj Is Nothing
            Set leadObj = ie.document.getElementById(identification)
        Loop
        .Range(leadCol & (counter + 10)).Value = leadObj.Value
        
        ' Finding Funding CC
        identification = "..." & CStr(16 + ((counter - jumpBack) * 23))
        Do While fundingObj Is Nothing
            Set fundingObj = ie.document.getElementById(identification)
        Loop
        .Range(fundCol & (counter + 10)).Value = fundingObj.Value
        
        ' Closing tab
        Set closeObj = ie.document.getElementsByClassName(". . .")(0)
        closeObj.Click
        Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
            DoEvents
        Loop
        Application.Wait (Now + TimeValue("0:00:01"))
    
    End With
    
    FormatRow targetSheet, counter + 10, "C"
    
    On Error GoTo 0

End Sub

' CQScrub sub in located in Module 1


