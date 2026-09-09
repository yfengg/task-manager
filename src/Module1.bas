Attribute VB_Name = "Module1"
Sub rank()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim startRow As Long
    startRow = 6
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then lastRow = startRow - 1
    lastRow = lastRow + 1

    Dim taskName As String, Urgency As String, Diffculty As String, Importance As String
    taskName = InputBox("Please enter Task name")
    If taskName = "" Then Exit Sub

    

    Urgency = InputBox("Please enter the urgency score(1-5)")
    If Urgency = "" Then Exit Sub
    
    Diffculty = InputBox("Please enter the Diffculty score(1-5)")
    If Diffculty = "" Then Exit Sub
    
    Importance = InputBox("Please enter the Importance score(1-5)")
    If Importance = "" Then Exit Sub
    
    
    
    

    ws.Cells(lastRow, 1).Value = taskName
    ws.Cells(lastRow, 2).Value = Urgency
    ws.Cells(lastRow, 3).Value = Diffculty
    ws.Cells(lastRow, 4).Value = Importance
    
    MsgBox "record has been added!"
End Sub


Sub StudyTarget()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet2")
    
    Dim startRow As Long
    startRow = 6
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then lastRow = startRow - 1
    lastRow = lastRow + 1
    
    Dim taskName As String
    
    
    Dim week As String
    Dim TargetStudyHour As String
    
    week = InputBox("Please enter the Week (e.g., Week 1, Week 2)")
    If week = "" Then Exit Sub
    
    TargetStudyHour = InputBox("Please enter the Target Study Hours")
    If TargetStudyHour = "" Then Exit Sub
    

    ws.Cells(lastRow, 1).Value = week
    ws.Cells(lastRow, 2).Value = TargetStudyHour
    

    
    
    MsgBox "New study target has been added!", vbInformation, "Success"
End Sub

Sub Tracking()
'replaced with addtask
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet4")
    
    Dim startRow As Long
    startRow = 5
    

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then lastRow = startRow - 1
    lastRow = lastRow + 1
    
    Dim taskName As String
    taskName = InputBox("Please enter Task name")
    If taskName = "" Then Exit Sub
    
    Dim dueday As String, Estimatedhours As String
    dueday = InputBox("Please enter dueday")
    If dueday = "" Then Exit Sub
    
    Estimatedhours = InputBox("Please enter the Estimated hours")
    If Estimatedhours = "" Then Exit Sub
    
    Dim status As String
    status = InputBox("Please select your status. 1. Done ,2.In Progress,3. Not Started")
    Select Case status
        Case "1": status = "Done"
        Case "2": status = "In Progress"
        Case "3": status = "Not Started"
        Case Else: status = "Not Started"
    End Select
    

    
    ws.Cells(lastRow, 1).Value = taskName
    ws.Cells(lastRow, 2).Value = dueday
    ws.Cells(lastRow, 3).Value = Estimatedhours
    ws.Cells(lastRow, 4).Value = status
    
    MsgBox "Record has been added!"
End Sub

Sub AddTask()
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim taskName As String
    Dim dueDate As Variant
    Dim estHours As Variant
    Dim status As String
    Dim priority As String
    
    Set ws = ThisWorkbook.Sheets("Sheet2")
    
    taskName = InputBox("Enter task name:")
    If taskName = "" Then Exit Sub
    
    dueDate = InputBox("Enter due date (MM/DD/YYYY):")
    If Not IsDate(dueDate) Then
        MsgBox "Invalid date."
        Exit Sub
    End If
    
    estHours = InputBox("Estimated hours:")
    If estHours = "" Then estHours = 0
    
    status = InputBox("Status (e.g. In Progress / Not Started / Done):")
    priority = InputBox("Priority (High, Medium, Low):")
    
    nextRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row + 1
    If nextRow < 5 Then nextRow = 5
    
    ws.Cells(nextRow, "I").Value = taskName
    ws.Cells(nextRow, "J").Value = dueDate
    ws.Cells(nextRow, "K").Value = estHours
    ws.Cells(nextRow, "L").Value = status
    ws.Cells(nextRow, "M").Value = priority
    
    MsgBox "Task Added!"
End Sub
Sub Schedule()
    Dim wsSched As Worksheet
    Dim wsTasks As Worksheet
    Dim wsAct As Worksheet
    Dim i As Long
    Dim lastRowTasks As Long
    Dim lastRowAct As Long
    Dim col As Long
    Dim nextRow As Long
    Dim nameStr As String, dueDate As Variant, priority As String
    Dim hours As Variant
    
    Dim weekStart As Date, weekEnd As Date
    Dim strDate As String

    strDate = InputBox("Enter the Monday of the week you want to generate (YYYY-MM-DD):", "Select Week")
    If strDate = "" Then Exit Sub

    If Not IsDate(strDate) Then
        MsgBox "Invalid date entered. Please enter a valid date in YYYY-MM-DD format.", vbExclamation
        Exit Sub
    End If

    weekStart = CDate(strDate)
    weekEnd = weekStart + 6

    
    Set wsSched = ThisWorkbook.Sheets("Schedule")
    Set wsTasks = ThisWorkbook.Sheets("Sheet2")
    Set wsAct = ThisWorkbook.Sheets("ActivityLog")
    
    wsSched.Cells.Clear
    
    wsSched.Cells(1, 2).Value = "Monday"
    wsSched.Cells(1, 3).Value = "Tuesday"
    wsSched.Cells(1, 4).Value = "Wednesday"
    wsSched.Cells(1, 5).Value = "Thursday"
    wsSched.Cells(1, 6).Value = "Friday"
    wsSched.Cells(1, 7).Value = "Saturday"
    wsSched.Cells(1, 8).Value = "Sunday"
    
    lastRowTasks = wsTasks.Cells(wsTasks.Rows.Count, "I").End(xlUp).Row
    
    For i = 5 To lastRowTasks
        nameStr = wsTasks.Cells(i, "I").Value
        dueDate = wsTasks.Cells(i, "J").Value
        hours = wsTasks.Cells(i, "K").Value
        priority = wsTasks.Cells(i, "M").Value
        
        If nameStr = "" Or Not IsDate(dueDate) Then GoTo NextTask
        
        If dueDate < weekStart Or dueDate > weekEnd Then GoTo NextTask
        
        col = Weekday(dueDate, vbMonday) + 1
        nextRow = 2
        Do While wsSched.Cells(nextRow, col).Value <> ""
            nextRow = nextRow + 1
        Loop
        
        wsSched.Cells(nextRow, col).Value = nameStr & " (" & hours & "h)"
        
        Select Case LCase(priority)
            Case "high": wsSched.Cells(nextRow, col).Interior.Color = RGB(255, 150, 150)
            Case "medium": wsSched.Cells(nextRow, col).Interior.Color = RGB(255, 230, 150)
            Case "low": wsSched.Cells(nextRow, col).Interior.Color = RGB(200, 255, 200)
        End Select
        
NextTask:
    Next i
    lastRowAct = wsAct.Cells(wsAct.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRowAct
        nameStr = wsAct.Cells(i, "B").Value
        dueDate = wsAct.Cells(i, "A").Value
        hours = wsAct.Cells(i, "C").Value
        priority = "medium"
        
        If nameStr = "" Or Not IsDate(dueDate) Then GoTo NextAct
        
        If dueDate < weekStart Or dueDate > weekEnd Then GoTo NextAct
        
        col = Weekday(dueDate, vbMonday) + 1
    
        nextRow = 2
        Do While wsSched.Cells(nextRow, col).Value <> ""
            nextRow = nextRow + 1
        Loop
        
        wsSched.Cells(nextRow, col).Value = nameStr & " (" & hours & "h)"
        wsSched.Cells(nextRow, col).Interior.Color = RGB(200, 220, 255) ' Light blue for activities
        
NextAct:
    Next i
End Sub

Sub PrintSchedule()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Schedule")
    
    With ws.PageSetup
        .PrintArea = ws.UsedRange.Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$1:$1"
    End With
    
    ws.PrintOut
End Sub

Sub DailyActivityForm()
    frmDailyActivities.Show
End Sub

Sub GotoSchedule()
    Sheets("Schedule").Activate
End Sub

Sub CheckTimeManagement()
    Dim ws As Worksheet
    Dim dayCol As Long, r As Long, lastRow As Long
    Dim studyHours As Double, freeHours As Double, actHours As Double
    Dim status As String
    
    Set ws = ThisWorkbook.Sheets("Schedule")
    
    ws.Range("J1").Value = "Day"
    ws.Range("K1").Value = "Study Hours"
    ws.Range("L1").Value = "Free Hours"
    ws.Range("M1").Value = "Activity Hours"
    ws.Range("N1").Value = "Status"
    
    Dim daysArr As Variant
    daysArr = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    
    Dim i As Long
    For i = 0 To 6
        ws.Cells(i + 2, "J").Value = daysArr(i)
    Next i

    For dayCol = 2 To 8
        
        studyHours = 0
        freeHours = 0
        actHours = 0
        
        lastRow = ws.Cells(ws.Rows.Count, dayCol).End(xlUp).Row
        For r = 2 To lastRow
            If ws.Cells(r, dayCol).Value <> "" Then
                Dim txt As String, h As Double
                txt = ws.Cells(r, dayCol).Value
                
                h = Val(Mid(txt, InStr(txt, "(") + 1))
                
                If InStr(1, txt, "study", vbTextCompare) > 0 Then
                    studyHours = studyHours + h
                ElseIf InStr(1, txt, "free", vbTextCompare) > 0 Then
                    freeHours = freeHours + h
                Else
                    actHours = actHours + h
                End If
            End If
        Next r
  
        Dim outputRow As Long
        outputRow = dayCol
        
        ws.Cells(outputRow, "K").Value = studyHours
        ws.Cells(outputRow, "L").Value = freeHours
        ws.Cells(outputRow, "M").Value = actHours
        
        If studyHours >= 4 And freeHours >= 2 Then
            status = "Balanced"
        ElseIf studyHours > 6 Then
            status = "Overloaded"
        Else
            status = "Needs Adjustment"
        End If
        
        ws.Cells(outputRow, "N").Value = status
    Next dayCol
    
    MsgBox "Time Management Summary Generated!", vbInformation
End Sub

Sub DailyPlanner()


    Dim ws As Worksheet
    Dim lastRow As Long
    Dim selectedDate As Date
    Dim outputStartRow As Long, outputCol As Long
    Dim i As Long
    Dim cellDate As Variant
    Dim userInput As String
    Dim writeRow As Long
    
    
    Set ws = ThisWorkbook.Sheets("ActivityLog")
    
  
    userInput = InputBox("Enter the date to generate the plan (yyyy-mm-dd):", "Select Date")
    
    If userInput = "" Then Exit Sub
    If Not IsDate(userInput) Then
        MsgBox "Invalid date. Try again.", vbExclamation
        Exit Sub
    End If
    
    selectedDate = CDate(userInput)
    
 
    outputStartRow = 1
    outputCol = 7
    writeRow = outputStartRow + 1
    

    ws.Range(ws.Cells(outputStartRow, outputCol), ws.Cells(100, outputCol)).ClearContents
    
    
    ws.Cells(outputStartRow, outputCol).Value = "Activities on " & selectedDate
    
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To lastRow
        cellDate = ws.Cells(i, 1).Value
        
        If Trim(cellDate) <> "" And IsDate(cellDate) Then
            If CDate(cellDate) = selectedDate Then
            
                Dim act As String
                Dim hrs As Variant
                
                act = ws.Cells(i, 2).Value
                hrs = ws.Cells(i, 3).Value
                
                ws.Cells(writeRow, outputCol).Value = act & " Ñ " & hrs & " hrs"
                writeRow = writeRow + 1
            End If
        End If
    Next i
    
   
    If writeRow = outputStartRow + 1 Then
        ws.Cells(writeRow, outputCol).Value = "No activities on this date."
    End If

End Sub

