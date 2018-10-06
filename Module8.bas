Attribute VB_Name = "Module8"
Option Explicit

Sub Clear_Button6_Click()

    Call Clearer(ws8)

End Sub

Sub Run_Button7_Click()

    Call BackendScrub(ws8)

End Sub

Sub BackendScrub(targetSheet As Worksheet)

    Dim lastRow As Integer
    Dim sizeByRow As Integer
    Dim conn As Object
    Dim rs As Object
    Dim row As Integer
    Dim col As String
    Dim workRequest As String
    Dim selectString As String
    Dim systemRows As Collection: Set systemRows = New Collection
    Dim extra As Integer: extra = 2
    Dim cellInfo As String
    Dim localCounter As Integer: localCounter = 0
    Dim stateSpot As Integer
    Dim releaseDate As String
    Dim systemsMissed As Boolean: systemsMissed = False
    Dim extraVert As Integer: extraVert = 0
    Dim systemsInd As String
    Dim indicator As String
    Dim i As Integer
    Dim db2Username As String: db2Username = "..."
    Dim db2Password As String: db2Password = "..."
    Dim deliv As Variant
    Dim delivArray() As Variant: delivArray = Array("UAT CoE Test Plan", "UAT CoE Test Scenarios", "UAT CoE Test Cases", _
                                                "UAT CoE RTM(UAT)", "UAT CoE Test Result", "UAT CoE Sign Off")
    
    
    Application.ScreenUpdating = False
    
    ' Input validation
    With targetSheet
        
        If .OLEObjects("ComboBox2BA").Object.Text = "On" Then
            If Len(Trim(CStr(.Range("C6").Value))) = 0 Then
                .Range("C6").Activate
                MsgBox "Please enter your LAN password!"
                Exit Sub
            End If
        End If
        
        ' Finding last row with empties and calculating size of list of WRs
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        sizeByRow = lastRow - 9
        
        ' getting dropdown values
        systemsInd = .OLEObjects("ComboBox2BA").Object.Text
        indicator = .OLEObjects("ComboBox3BA").Object.Text
        
    End With
    
    ' Checking to make sure there are WRs in row
    If sizeByRow = 0 Then
        MsgBox "No Work Requests Entered"
        Exit Sub
    End If
    
    ' clean wrs
    Call CleanupWRs(targetSheet, 10, "A")

    ' updating status bar
    Application.StatusBar = "Running"
    
    ' set headers
    Call FormatColumnHeadings(indicator, systemsInd, targetSheet)
    
    ' set select string
    selectString = SelectStringSetup(indicator, systemsInd)
    
    ' set wr col
    col = SetColumn(indicator)
    
    ' Organize WRs and Headers
    If indicator = "Master Tracker" Then
        ShuffleWRs targetSheet, lastRow
    End If
    
    ' Initializing connection objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Connected to database
    On Error GoTo ConnectionHandler
    conn.ConnectionString = "Provider=IBMDADB2.1;UID=" & db2Username & ";PWD=" & db2Password & ";Data Source=...;ProviderType=OLEDB"
    conn.Open

    On Error GoTo 0
    With targetSheet
    
        Select Case indicator
            Case "Intake"
                Call IntakeProcess(targetSheet, selectString, conn, rs, systemRows, systemsMissed, lastRow)

            Case "Master Tracker"
                Call MasterTrackerProcess(targetSheet, selectString, conn, rs, systemRows, systemsMissed, lastRow)
                
            Case "Funding CC"
                Call FundingProcess(targetSheet, selectString, conn, rs, systemRows, systemsMissed, lastRow)
    
            Case "UAT Hours"
                Call UatHoursProcess(targetSheet, selectString, conn, rs, systemRows, systemsMissed, lastRow)
                
            Case "Planning Tab"
                Call PlanningProcess(targetSheet, selectString, conn, rs, systemRows, systemsMissed, lastRow)

        End Select
        
    End With
    
    ' close connection
    conn.Close
    
    ' function to remove bold
    Call UnBolder(targetSheet, col)
    
    ' Checking front end for systems if missed
    If systemsMissed = True Then
        FrontEndRunner2 sizeByRow, targetSheet
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    MsgBox "Done!"
    Exit Sub
    
' For connection string error (db2Username & db2Password)
ConnectionHandler:
    
    MsgBox "Error connecting to database"
    Application.StatusBar = ""
    Exit Sub

End Sub

Sub UnBolder(targetSheet As Worksheet, wrCol As String)

    Dim lastRow As Long
    Dim lastCol As Integer
    
    With targetSheet
        lastRow = .Cells(.Rows.Count, wrCol).End(xlUp).row
        lastCol = .Cells(9, .Columns.Count).End(xlToLeft).Column
        
        If lastRow <= 9 Then Exit Sub
        
        .Range(.Cells(10, 1), .Cells(lastRow, lastCol)).Font.Bold = False
    End With

End Sub

Sub FormatColumnHeadings(indicator As String, systemsInd As String, targetSheet As Worksheet)

    With targetSheet
        
        Select Case indicator
            Case "Intake"
                .Range("A9").Value = "Work Request"
                .Range("B9").Value = "IT Lead"
                .Range("C9").Value = "State"
                .Range("D9").Value = "Uat Inv."
                .Range("E9").Value = "Release"
                .Range("F9").Value = "Headline"
                
                .Columns("A").ColumnWidth = 22
                .Columns("B").ColumnWidth = 34
                .Columns("C").ColumnWidth = 28
                .Columns("D").ColumnWidth = 21
                .Columns("E").ColumnWidth = 28
                .Columns("F").ColumnWidth = 25
                
                If systemsInd = "On" Then
                    .Range("G9").Value = "Systems"
                    .Columns("G").ColumnWidth = 45
                End If
                
            Case "Master Tracker"
                .Range("A9").Value = "UAT-COE Lead"
                .Range("B9").Value = "Project Name"
                .Range("C9").Value = "Work Request"
                .Range("D9").Value = "Release"
                .Range("E9").Value = "UAT-COE SME for Sign Off"
                .Range("F9").Value = "CQ Status"
                .Range("G9").Value = "UAT Inv"
                .Range("H9").Value = "IT Lead"
                
                .Columns("A").ColumnWidth = 22
                .Columns("B").ColumnWidth = 42
                .Columns("C").ColumnWidth = 28
                .Columns("D").ColumnWidth = 22
                .Columns("E").ColumnWidth = 33
                .Columns("F").ColumnWidth = 19
                .Columns("G").ColumnWidth = 15
                .Columns("H").ColumnWidth = 27
                
            Case "Funding CC"
                .Range("A9").Value = "Work Request"
                .Range("B9").Value = "IT Lead"
                .Range("C9").Value = "Funding CC"
                
                .Columns("A").ColumnWidth = 22
                .Columns("B").ColumnWidth = 26
                .Columns("C").ColumnWidth = 24
                
            Case "UAT Hours"
                .Range("A9").Value = "Work Request"
                .Range("B9").Value = "UAT Hours"
                .Range("C9").Value = "SIT Hours"
                .Range("D9").Value = "Dev Hours"
                .Range("E9").Value = "Type"
                .Range("F9").Value = "Class"
                
                .Columns("A").ColumnWidth = 22
                .Columns("B").ColumnWidth = 26
                .Columns("C").ColumnWidth = 26
                .Columns("D").ColumnWidth = 26
                .Columns("E").ColumnWidth = 26
                .Columns("F").ColumnWidth = 26
                
            Case "Planning Tab"
                .Range("A9").Value = "Work Request"
                .Range("B9").Value = "Deliverables"
                .Range("C9").Value = "Initial Milestone Date"
                .Range("D9").Value = "Final Milestone Date"
                .Range("E9").Value = "Committed By"
                .Range("F9").Value = "Status"
                
                .Columns("A").ColumnWidth = 22
                .Columns("B").ColumnWidth = 26
                .Columns("C").ColumnWidth = 26
                .Columns("D").ColumnWidth = 26
                .Columns("E").ColumnWidth = 26
                .Columns("F").ColumnWidth = 26
            
        End Select
    End With

End Sub

Function SelectStringSetup(indicator As String, systemsInd As String) As String

    Select Case indicator
        Case "Intake"
            SelectStringSetup = " ..."
            If systemsInd = "On" Then
                SelectStringSetup = SelectStringSetup & ", ... "
            End If
        Case "Master Tracker"
            SelectStringSetup = " ..."
        Case "Funding CC"
            SelectStringSetup = " ..."
        Case "UAT Hours"
            SelectStringSetup = " ... "
        Case "Planning Tab"
            SelectStringSetup = " ..."
    End Select

End Function

Function SetColumn(indicator As String) As String

    SetColumn = "A"
    
    If indicator = "Master Tracker" Then
        SetColumn = "C"
    End If

End Function

Sub Clearer(targetSheet As Worksheet)

    Dim conn As Long
    Dim lastRow As Integer
    Dim lastRow2 As Integer
    Dim counter As Integer
    Dim row As Long
    
    Application.ScreenUpdating = False
    
    ' Clearing previous connections
    With ThisWorkbook
        For conn = .Connections.Count To 1 Step -1
            .Connections(conn).Delete
        Next conn
    End With

    ' Clearing previous content
    With targetSheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        lastRow2 = .Cells(.Rows.Count, "C").End(xlUp).row
    End With
    
    If lastRow2 > lastRow Then
        lastRow = lastRow2
    End If
    
    With targetSheet
    
        If Left(Trim(CStr(.Range("C" & 10).Value)), 5) = "HBCBS" Then
            .Range("A10:A" & lastRow).Value = .Range("C10:C" & lastRow).Value
        End If
        
        If Len(CStr(.Range("C9").Value)) > 0 Then
            
            .Range("A9:P" & lastRow).UnMerge
            .Range("A9:P" & lastRow).ClearContents
            .Range("A9:P" & lastRow).Interior.ColorIndex = 0
            .Range("A9").Value = "Enter WRs Below"
            .Columns("A").ColumnWidth = 29
            .Columns("B").ColumnWidth = 21
            .Columns("C").ColumnWidth = 29
            .Columns("D").ColumnWidth = 16
            .Columns("E").ColumnWidth = 24
            .Columns("F").ColumnWidth = 22
            .Columns("G").ColumnWidth = 8.11
            .Columns("H").ColumnWidth = 8.11
            
            For row = 10 To lastRow
                If Len(Trim(CStr(.Range("A" & row).Value))) = 0 Then
                    .Rows(row).Delete
                    row = row - 1
                    counter = counter + 1
                    If counter > lastRow Then
                        Exit For
                    End If
                End If
            Next
        End If
        
        ' removing borders
        .Range("A9:P" & lastRow).Borders.LineStyle = xlNone
        .Range("A9:H9").Borders.LineStyle = xlContinuous
        .Range("A9:H9").Font.Bold = True
        
    End With
    
    Application.ScreenUpdating = True
    
End Sub

Sub FrontEndRunner2(sizeByRow As Integer, targetSheet As Worksheet)

    Dim i As Integer: i = 0
    Dim ie As InternetExplorer
    Dim objCollection As Object
    Dim loginObj As Object
    Dim ObjElement As Object
    Dim searchStringObj As Object
    Dim counter As Integer: counter = 0
    Dim closeObj As Object
    Dim logoutObj As Object
    Dim pwObj As Object
    Dim checkedSoFar As Integer: checkedSoFar = 0
    Dim password
    
    Application.StatusBar = "Front End Grabbing Systems"
    
    On Error GoTo ErrHandler2
    
    Set ie = CreateObject("InternetExplorer.Application")
    ' Uncomment below for testing
'    ie.Height = 1000
'    ie.Width = 1000
    ie.Visible = False
    
    On Error Resume Next
    ie.navigate "http://clearquest/cqweb/"
    
    On Error GoTo 0
    Application.StatusBar = "Loading http//..."
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:07"))
    Application.StatusBar = "Please wait..."
    
    With targetSheet
    
        ' First checking if user is logged in
        password = .Range("C6").Value
        Set objCollection = ie.document.getElementsByTagName("..")
    
        Do While i < objCollection.Length
            If objCollection(i).Name = "passwordId" Then
                objCollection(i).Value = password
                Set loginObj = ie.document.getElementById("...")
                loginObj.Click
                Application.Wait (Now + TimeValue("0:00:05"))
                Exit Do
            End If
            i = i + 1
        Loop: i = 0
        
        ' First searching for "search" bar and saving as object
        Do While i < objCollection.Length
            If objCollection(i).Name = "..." Then
                Set searchStringObj = objCollection(i)
                Exit Do
            End If
            i = i + 1
        Loop
        Set ObjElement = ie.document.getElementById("...")
    
        ' Starting main loop of checking WRs
        For counter = 0 To sizeByRow
    
            If .Range("G" & (counter + 10)).Interior.Color = RGB(0, 125, 255) Then
            
                Call IntakeLooper2(counter, sizeByRow, ie, searchStringObj, ObjElement, checkedSoFar, targetSheet)
                checkedSoFar = checkedSoFar + 1
                
            End If
    
        Next
    
    End With
    
    Application.StatusBar = "Finishing Process..."
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Set closeObj = ie.document.getElementsByClassName(". . .")(0)
    closeObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Set logoutObj = ie.document.getElementById("...")
    logoutObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:03"))

    ' Closing down program
    Set pwObj = ie.document.getElementById("...")
    pwObj.Value = password
    Set loginObj = ie.document.getElementById("...")
    loginObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:06"))
    ie.Quit
    Set ie = Nothing
    Set ObjElement = Nothing
    Set objCollection = Nothing
    Application.StatusBar = ""
    Exit Sub

ErrHandler2:

    MsgBox "Please wait a couple seconds. IE is still closing from previous use. Please restart Wait, Open & Close Worksheet, or Force Shutdown Internet Explorer"

End Sub

Sub IntakeLooper2(counter As Integer, sizeByRow As Integer, ie As InternetExplorer, searchStringObj As Object, ObjElement As Object, checkedSoFar As Integer, targetSheet As Worksheet)

    Dim closeObj        As Object
    Dim systemsObj      As Object
    Dim systems         As String
    Dim systemsID       As String

    ' Updating the status bar programmatically
    Application.StatusBar = "Checking " & CStr(targetSheet.Range("A" & CStr(counter + 10)).Value) & ", " & CStr(Round(counter / sizeByRow, 2) * 100) & "% Complete"
    
    searchStringObj.Value = CStr(targetSheet.Range("A" & CStr(counter + 10)).Value)
    ObjElement.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:06"))

    ' Finding Systems Tested
    systemsID = "..." & CStr(1 + (checkedSoFar * 4))
    On Error Resume Next
    Set systemsObj = ie.document.getElementById(systemsID)
    systems = systemsObj.innerText
    targetSheet.Range("G" & (counter + 10)).Value = systems
    targetSheet.Range("G" & (counter + 10)).Activate
    targetSheet.Range("G" & (counter + 10)).Interior.ColorIndex = 0
    
    ' Closing tab
    Set closeObj = ie.document.getElementsByClassName(". . .")(0)
    closeObj.Click
    Do While ie.Busy Or Not ie.readyState = READYSTATE_COMPLETE
        DoEvents
    Loop
    Application.Wait (Now + TimeValue("0:00:01"))
    
End Sub

Sub ShuffleWRs(targetSheet As Worksheet, lastRow As Integer)

    Dim row As Integer

    With targetSheet
        For row = 10 To lastRow
            .Range("C" & row).NumberFormat = "@"
            .Range("C" & row).Value = .Range("A" & row).Value
        Next
        .Range("A10:A" & lastRow).ClearContents
        .Range("A10:A" & lastRow).Borders.LineStyle = xlNone
        .Range("A10:A" & lastRow).Interior.ColorIndex = 0
        .Range("A9").Borders.LineStyle = xlContinuous
        .Range("D10:D" & lastRow).NumberFormat = "m/d/yyyy"
    End With

End Sub

Sub IntakeProcess(targetSheet As Worksheet, selectString As String, conn As Object, rs As Object, ByRef systemRows As Collection, ByRef systemsMissed As Boolean, lastRow As Integer)

    Dim workRequest As String
    Dim row As Integer
    Dim i As Integer
    
    With targetSheet
    
        For row = 10 To lastRow
            workRequest = Trim(CStr(.Cells(row, "A").Value))
            rs.Open "Select " & selectString & " from ... as t1 inner join ... as t2 on t1... = t2... where t1.id = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                On Error GoTo Skipper3
                For i = 0 To rs.Fields.Count - 1
                    .Cells(row, i + 2).Value = rs.Fields(i).Value
                Next
                .Range("E" & row).NumberFormat = "m/d/yyyy"
                On Error GoTo 0
            Else
                .Range("B" & row).Value = "Empty Query"
                .Range("B" & row).Interior.Color = RGB(255, 0, 0)
            End If
            rs.Close
        Next
        
    End With
    
    Exit Sub
    
Skipper3:

    With targetSheet
        .Range("G" & row).Interior.Color = RGB(0, 125, 255)
        systemRows.Add .Range("A" & row).Value
    End With
    
    systemsMissed = True
    Resume Next

End Sub


Sub MasterTrackerProcess(targetSheet As Worksheet, selectString As String, conn As Object, rs As Object, ByRef systemRows As Collection, ByRef systemsMissed As Boolean, lastRow As Integer)

    Dim workRequest As String
    Dim row As Integer

    With targetSheet

    For row = 10 To lastRow
        workRequest = Trim(CStr(.Cells(row, "C").Value))
        rs.Open "Select " & selectString & " from ... as t1 inner join ... as t2 on t1... = t2... where t1... = '" & workRequest & "' with ur", conn
        If Not (rs.BOF Or rs.EOF) Then
            .Range("B" & row).Value = rs.Fields(0).Value
            .Range("D" & row).Value = rs.Fields(1).Value
            .Range("F" & row).Value = rs.Fields(2).Value
            .Range("G" & row).Value = rs.Fields(3).Value
            .Range("H" & row).Value = rs.Fields(4).Value
        Else
            .Range("B" & row).Value = "Empty Query"
            .Range("B" & row).Interior.Color = RGB(255, 0, 0)
        End If
        rs.Close
    Next
    
    End With

End Sub

Sub FundingProcess(targetSheet As Worksheet, selectString As String, conn As Object, rs As Object, ByRef systemRows As Collection, ByRef systemsMissed As Boolean, lastRow As Integer)

    Dim workRequest As String
    Dim row As Integer

    With targetSheet

        For row = 10 To lastRow
            workRequest = Trim(CStr(.Cells(row, "A").Value))
            rs.Open "Select " & selectString & " from ... where ... = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                On Error GoTo Skipper3
                For i = 0 To rs.Fields.Count - 1
                    .Cells(row, i + 2).NumberFormat = "0"
                    .Cells(row, i + 2).Value = rs.Fields(i).Value
                Next
                On Error GoTo 0
            Else
                .Range("B" & row).Value = "Empty Query"
                .Range("B" & row).Interior.Color = RGB(255, 0, 0)
            End If
            rs.Close
        Next
        
    End With
    
    Exit Sub
    
Skipper3:

    With targetSheet
        .Range("G" & row).Interior.Color = RGB(0, 125, 255)
        systemRows.Add .Range("A" & row).Value
    End With
    
    systemsMissed = True
    Resume Next
    
End Sub

Sub UatHoursProcess(targetSheet As Worksheet, selectString As String, conn As Object, rs As Object, ByRef systemRows As Collection, ByRef systemsMissed As Boolean, lastRow As Integer)

    Dim workRequest As String
    Dim row As Integer

    With targetSheet

        For row = 10 To lastRow
            workRequest = Trim(CStr(.Cells(row + extraVert, "A").Value))
            rs.Open "Select " & selectString & " from ... where ... = '" & workRequest & "' with ur", conn
            If Not (rs.BOF Or rs.EOF) Then
                Do While Not rs.EOF
                    On Error GoTo Skipper3
                    For i = 0 To rs.Fields.Count - 1
                        .Cells(row + extraVert, i + 2).NumberFormat = "0"
                        .Cells(row + extraVert, i + 2).Value = rs.Fields(i).Value
                    Next
                    rs.MoveNext
                    If Not (rs.BOF Or rs.EOF) Then ' added here
                        .Rows(row + extraVert + 1).Insert shift:=xlShiftDown
                        .Range("A" & CStr(row + extraVert) & ":A" & CStr(row + extraVert + 1)).Merge
                        .Range("A" & CStr(row + extraVert) & ":A" & CStr(row + extraVert + 1)).VerticalAlignment = xlCenter
                        extraVert = extraVert + 1
                    End If
                    On Error GoTo 0
                Loop
            Else
                .Range("B" & row + extraVert).Value = "Empty Query"
                .Range("B" & row + extraVert).Interior.Color = RGB(255, 0, 0)
                
            End If
            rs.Close
        Next

    End With
    
    Exit Sub
    
Skipper3:

    With targetSheet
        .Range("G" & row).Interior.Color = RGB(0, 125, 255)
        systemRows.Add .Range("A" & row).Value
    End With
    
    systemsMissed = True
    Resume Next

End Sub

Sub PlanningProcess(targetSheet As Worksheet, selectString As String, conn As Object, rs As Object, ByRef systemRows As Collection, ByRef systemsMissed As Boolean, lastRow As Integer)

    Dim workRequest As String
    Dim row As Integer

    With targetSheet
    
        For row = 10 To lastRow
            workRequest = Trim(CStr(.Cells(row, "A").Value))
            For Each deliv In delivArray
                rs.Open "Select " & selectString & " from ... where ... = '" & workRequest & "' and ... = '" & deliv & "' with ur", conn
                If Not (rs.BOF Or rs.EOF) Then
                    On Error GoTo Skipper3
                    For i = 0 To rs.Fields.Count - 1
                        If Len(Trim(CStr(.Cells(row, i + 2).Value))) = 0 Then
                            cellInfo = ""
                        Else
                            cellInfo = .Cells(row, i + 2).Value & " " & vbCrLf
                        End If
                        .Cells(row, i + 2).NumberFormat = "@"
                        .Cells(row, i + 2).Value = cellInfo & rs.Fields(i).Value
                    Next
                    On Error GoTo 0
                    
                End If
                rs.Close
            Next
        Next
    
    End With
    
    Exit Sub
    
Skipper3:

    With targetSheet
        .Range("G" & row).Interior.Color = RGB(0, 125, 255)
        systemRows.Add .Range("A" & row).Value
    End With
    
    systemsMissed = True
    Resume Next

End Sub
