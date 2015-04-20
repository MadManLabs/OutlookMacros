Attribute VB_Name = "mdlEmailCreation"
Public Sub btnMissedMessage()
    MsgForm.Show
End Sub


Sub SendMessage(displayMsg As Boolean, Optional AttachmentPath)
    Dim objOutlook As Outlook.Application
    Dim objOutlookMsg As Outlook.MailItem
    Dim objOutlookRecip As Outlook.Recipient
    Dim objOutlookAttach As Outlook.Attachment
    Dim subjectText As String
    Dim bodyText As String

    ' Create the Outlook session.
    Set objOutlook = CreateObject("Outlook.Application")

    ' Create the message.
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

    With objOutlookMsg
        ' Add the To recipient(s) to the message.
        Set objOutlookRecip = .Recipients.Add(MsgForm.cboxFieldTo.Text)
        objOutlookRecip.Type = olTo

'        ' Add the CC recipient(s) to the message.
'        Set objOutlookRecip = .Recipients.Add()
'        objOutlookRecip.Type = olCC
'
'       ' Add the BCC recipient(s) to the message.
'        Set objOutlookRecip = .Recipients.Add()
'        objOutlookRecip.Type = olBCC

        'Create subject text.
        If MsgForm.optCalled.Value Or MsgForm.optReturned.Value Then
            If MsgForm.optVisited.Value Or _
            MsgForm.optSchedule.Value Then
                subjectText = "MISSED MESSAGE"
            Else
                subjectText = "MISSED CALL"
            End If
        ElseIf MsgForm.optVisited.Value Then
            If MsgForm.optCalled.Value Or MsgForm.optSchedule.Value _
            Or MsgForm.optReturned.Value Then
                subjectText = "MISSED MESSAGE"
            Else
                subjectText = "MISSED VISIT"
            End If
            
        ElseIf MsgForm.optSchedule.Value Then
            If MsgForm.optCalled.Value Or _
            MsgForm.optVisited.Value Or MsgForm.optReturned.Value Then
                subjectText = "MISSED MESSAGE"
            Else
                subjectText = "APPOINTMENT REQUEST"
            End If
        Else
            subjectText = "MISSED MESSAGE"
        End If
        
        If Not MsgForm.txtbFieldCaller.Text = "" Then
            subjectText = subjectText + ": " + MsgForm.txtbFieldCaller.Text
        End If
        
        If Not MsgForm.txtbFieldBusiness.Text = "" Then
            If Not MsgForm.txtbFieldCaller.Text = "" Then
                subjectText = subjectText + " /"
            End If
            subjectText = subjectText + " " + MsgForm.txtbFieldBusiness.Text
        End If
        
        If Not MsgForm.txtbFirstThree.Text = "" Then
            If (Not MsgForm.txtbFieldCaller.Text = "") Or _
            (Not MsgForm.txtbFieldBusiness.Text = "") Then
                subjectText = subjectText + " /"
            End If
            
            If Not MsgForm.txtbAreaCode.Text = "" Then
                subjectText = subjectText + " (" + MsgForm.txtbAreaCode.Text + ")"
            End If
            subjectText = subjectText + " " + MsgForm.txtbFirstThree.Text + _
            "-" + MsgForm.txtbLastFour.Text
        End If
        
        ' Set the Subject, Body, and Importance of the message.
       .Subject = subjectText
       
        If MsgForm.radUrgent.Value Or MsgForm.radEOD.Value Then
            .Importance = olImportanceHigh  'High importance
        End If
        
        bodyText = MsgForm.cboxFieldTo.Text + "," & vbCrLf & vbCrLf + _
        "While you were unavailable:" & vbCrLf
        If Not MsgForm.txtbFieldCaller.Text = "" Then
            bodyText = bodyText + MsgForm.txtbFieldCaller.Text
            If Not MsgForm.txtbFieldBusiness.Text = "" Then
                bodyText = bodyText + " with " + MsgForm.txtbFieldBusiness.Text
            End If
        Else
            bodyText = bodyText + MsgForm.txtbFieldBusiness.Text
        End If
        
        If MsgForm.optCalled.Value Then
            bodyText = bodyText + " called" + vbCrLf + vbCrLf
        ElseIf MsgForm.optVisited.Value Then
            bodyText = bodyText + " visited" + vbCrLf + vbCrLf
        ElseIf MsgForm.optReturned.Value Then
            bodyText = bodyText + " returned your call" + vbCrLf + vbCrLf
        ElseIf MsgForm.optSchedule.Value Then
            bodyText = bodyText + " wishes to meet with you" + vbCrLf + vbCrLf
        Else
            bodyText = bodyText + " did something that warrants a message" + vbCrLf + vbCrLf
        End If
        
        bodyText = bodyText + "Action requested:" + vbCrLf
        
        If MsgForm.optCallBack.Value Then
            bodyText = bodyText + "Please call back" + vbCrLf + vbCrLf
        ElseIf MsgForm.optWaitForCall.Value Then
            bodyText = bodyText + "None - will call back" + vbCrLf + vbCrLf
        ElseIf MsgForm.optEmailBack.Value Then
            bodyText = bodyText + "Please email back" + vbCrLf + vbCrLf
        ElseIf MsgForm.optCheckEmail.Value Then
            bodyText = bodyText + "None - will email you" + vbCrLf + vbCrLf
        Else
            bodyText = bodyText + "None" + vbCrLf + vbCrLf
        End If
        
        If Not MsgForm.txtbMessage.Text = "" Then
            bodyText = bodyText + "Message:" + vbCrLf + MsgForm.txtbMessage.Text + _
            vbCrLf + vbCrLf
        End If
        
        If MsgForm.radUrgent.Value Or MsgForm.radEOD.Value Or _
        MsgForm.radASAP.Value Then
            bodyText = bodyText + "Priority:" + vbCrLf
            
            If MsgForm.radUrgent.Value Then
                bodyText = bodyText + "URGENT/ASAP" + vbCrLf + vbCrLf
            ElseIf MsgForm.radEOD.Value Then
                bodyText = bodyText + "By end of day" + vbCrLf + vbCrLf
            ElseIf MsgForm.radASAP.Value Then
                bodyText = bodyText + "When you get the chance" + vbCrLf + vbCrLf
            End If
        End If
        
        bodyText = bodyText + "Thanks," + vbCrLf + MsgForm.cboxFieldFrom.Text
       
        .Body = bodyText


        ' Add attachments to the message.
        If Not IsMissing(AttachmentPath) Then
            Set objOutlookAttach = .Attachments.Add(AttachmentPath)
        End If
        
        ' Resolve each Recipient's name.
        For Each objOutlookRecip In .Recipients
            objOutlookRecip.Resolve
        Next

        Unload MsgForm

       ' Should we display the message before sending?
        If displayMsg Then
            .Display
        Else
            .Save
            .Send
        End If
        
    End With
    Set objOutlook = Nothing
End Sub
