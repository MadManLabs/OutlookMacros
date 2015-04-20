VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfReportDate 
   Caption         =   "Reporting Date"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "usfReportDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfReportDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCreateTask_Click()
    createPayrollTask
End Sub

Private Sub UserForm_Initialize()
    cboxYearPrior.AddItem "2015"
    cboxYearPrior.AddItem "2016"
    cboxYearPrior.AddItem "2017"
    cboxYearPrior.AddItem "2018"
    
    cboxYear.AddItem "2015"
    cboxYear.AddItem "2016"
    cboxYear.AddItem "2017"
    cboxYear.AddItem "2018"
    
    cboxYear.Text = "2015"
    cboxYearPrior.Text = "2015"
End Sub

Private Sub txtbMonth_Change()
    If Len(txtbMonth.Text) = txtbMonth.MaxLength Then
        txtbDay.SetFocus
    End If
End Sub

Private Sub txtbDay_Change()
    If Len(txtbDay.Text) = txtbDay.MaxLength Then
        btnCreateTask.SetFocus
    End If
End Sub

Private Sub txtbMonthPrior_Change()
    If Len(txtbMonthPrior.Text) = txtbMonthPrior.MaxLength Then
        txtbDayPrior.SetFocus
    End If
End Sub

Private Sub txtbDayPrior_Change()
    If Len(txtbDayPrior.Text) = txtbDayPrior.MaxLength Then
        txtbMonth.SetFocus
    End If
End Sub

Private Sub createPayrollTask()
    Dim obj As Object
    Dim objOutlook As Outlook.Application
    Dim reportDate As String
    Dim objNewTask As Outlook.TaskItem
    Dim reminderTime As String
    Dim subjectText As String
    Dim bodyText As String

    ' Create the Outlook session.
    Set objOutlook = CreateObject("Outlook.Application")

    ' Create the task.
    Set objNewTask = objOutlook.CreateItem(olTaskItem)
    
    Dim startDate As Date, dueDate As Date, priorDate As Date
    Dim strSubject As String, strBody As String
    
    strSubject = "REPORT HOURS by "
    strBody = "For the time period of "
    reminderTime = "14:00:00"
    
    reportDate = txtbMonthPrior.Text & "/" & txtbDayPrior.Text & _
                    "/" & cboxYearPrior.Text
    priorDate = DateValue(reportDate)
    
    reportDate = txtbMonth.Text & "/" & txtbDay.Text & "/" & cboxYear.Text
    dueDate = DateValue(reportDate)
    
    Select Case Weekday(dueDate)
        Case vbMonday
            startDate = DateAdd("d", -3, dueDate)
        Case Else
            startDate = DateAdd("d", -1, dueDate)
    End Select
    
    strBody = strBody & MonthName(Month(priorDate), True) & " " & _
                Day(priorDate) & " to " & MonthName(Month(startDate), True) _
                & " " & Day(startDate) & ", I have worked a total of * hours"
    strSubject = strSubject & MonthName(Month(dueDate), True) & " " & _
                Day(dueDate)
    
    With objNewTask
        .Body = strBody
        .startDate = startDate
        .dueDate = dueDate
        .Subject = strSubject
        .ReminderSet = True
        .reminderTime = startDate + CDate(reminderTime)
        .Save
    End With
    
    Dim answer As Integer
    answer = MsgBox("Create next date?", vbYesNo)
    Select Case answer
        Case vbYes
            txtbMonthPrior.Text = txtbMonth.Text
            txtbDayPrior.Text = txtbDay.Text
            cboxYearPrior.Text = cboxYear.Text
            txtbMonth.Text = ""
            txtbDay.Text = ""
            txtbMonth.SetFocus
        Case vbNo
            Unload usfReportDate
    End Select
End Sub

Private Sub test()
    Dim reportDate As String
    reportDate = "01/02/1989"
    Dim exDate As Date
    exDate = DateValue(reportDate)
    MsgBox Year(exDate)
End Sub
