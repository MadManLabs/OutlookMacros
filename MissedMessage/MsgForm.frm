VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MsgForm 
   Caption         =   "UserForm1"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   OleObjectBlob   =   "MsgForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MsgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    cboxFieldTo.AddItem "Bob Fletcher"
    cboxFieldTo.AddItem "Tom Darr"
    cboxFieldTo.AddItem "Chad Burnett"
    cboxFieldTo.AddItem "Deanne Puckett"
    cboxFieldTo.AddItem "Diane Lund"
    cboxFieldTo.AddItem "Kyle Evora"
    cboxFieldTo.AddItem "David Skaggs"
    
    cboxFieldFrom.AddItem "Bob Fletcher"
    cboxFieldFrom.AddItem "Tom Darr"
    cboxFieldFrom.AddItem "Chad Burnett"
    cboxFieldFrom.AddItem "Deanne Puckett"
    cboxFieldFrom.AddItem "Diane Lund"
    cboxFieldFrom.AddItem "Kyle Evora"
    cboxFieldFrom.AddItem "David Skaggs"
   
    Dim myNamespace As Outlook.NameSpace
    Set myNamespace = Application.GetNamespace("MAPI")
    cboxFieldFrom.Text = myNamespace.CurrentUser
    
    radASAP.Value = True
End Sub

Private Sub btnSend_Click()
    Dim displayMsg As Boolean
    displayMsg = False
    
    Dim continueSend As Boolean
    continueSend = True
    
    Dim MsgError As String
    MsgError = "The following error(s) were found:" + vbNewLine + vbNewLine
    
    If IsEmpty(cboxFieldTo) Then
        MsgError = Chr(183) + " Missing recipient."
        continueSend = False
    End If
    
    If IsEmpty(txtbFieldCaller) And IsEmpty(txtbFieldBusiness) Then
        MsgError = Chr(183) + " Need at least Caller or Business"
        continueSend = False
    End If
    
    If (Not txtbAreaCode.Text = "") Or (Not txtbFirstThree.Text = "") Or _
    (Not txtbLastFour.Text = "") Then
        If (Not txtbAreaCode.Text = "") Then
            If txtbAreaCode.TextLength <> 3 Then
                MsgError = Chr(183) + " Phone # Error: Area Code too short."
                continueSend = False
            End If
        End If
        
        If txtbFirstThree.TextLength <> 3 Then
            MsgError = Chr(183) + " Phone # Error: Missing digits."
            continueSend = False
        End If
        
        If txtbLastFour.TextLength <> 4 Then
            MsgError = Chr(183) + " Phone # Error: Missing digits."
            continueSend = False
        End If
    End If
    
    If Me.optDisplayEmail.Value Then
        displayMsg = True
    End If
    
    If continueSend Then
        mdlEmailCreation.SendMessage displayMsg
    End If
End Sub

Private Sub txtbAreaCode_Change()
    If Len(txtbAreaCode.Text) = txtbAreaCode.MaxLength Then
        txtbFirstThree.SetFocus
    End If
End Sub

Private Sub txtbFirstThree_Change()
    If Len(txtbFirstThree.Text) = txtbFirstThree.MaxLength Then
        txtbLastFour.SetFocus
    End If
End Sub
