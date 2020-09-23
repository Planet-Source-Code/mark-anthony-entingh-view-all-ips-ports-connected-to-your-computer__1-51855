VERSION 5.00
Begin VB.UserControl File 
   BackColor       =   &H00F0F0F0&
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   94
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   105
      Top             =   -90
   End
   Begin VB.TextBox TxtChange 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   315
      MaxLength       =   50
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox imgIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   30
      Width           =   255
      Begin VB.Image ImgErr 
         Height          =   240
         Left            =   0
         Picture         =   "File.ctx":0000
         Top             =   -15
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgMsg 
         Height          =   240
         Left            =   0
         Picture         =   "File.ctx":0374
         Top             =   -1
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00F0F0F0&
      Caption         =   "MyVPN_6.conn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00787878&
      Height          =   240
      Left            =   315
      TabIndex        =   0
      Top             =   30
      Width           =   1080
   End
End
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event DblClick()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event MessageOver(msgtype)

Public Indx As Integer, ArrayIndex As Integer, MenuIndex As Integer, TCPBufferIndex

Private Type typFile
    Icon As Integer
End Type

Private cFile As typFile, Focused As Single
Private msgtype As Integer, BlinkCount As Integer

Private Sub imgIcon_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub imgIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    UserControl_EnterFocus
End Sub

Private Sub Label_Change()
    UserControl.Width = (Label.Left + Label.Width) * Screen.TwipsPerPixelX
End Sub

Public Property Let Caption(txt)
    Label.Caption = txt
End Property

Public Property Get Caption()
    Caption = Label.Caption
End Property

Public Property Set Icon(fil)
    imgIcon.Picture = Nothing
    imgIcon.Picture = fil
    'cFile.Icon = fil
End Property

Public Property Let Icon(fil)
    fil = cFile.Icon
End Property

Private Sub Label_Click()
If Focused = 0 Then
        Focused = 2
    ElseIf Focused = 1 Then
        Focused = 2
    ElseIf Focused = 2 Then
        Focused = 1
        Pause 0.25
        If Focused = 1 Then
        
        Label.Visible = False
        TxtChange.Text = Label.Caption
        TxtChange.Width = UserControl.ScaleWidth - TxtChange.Left
        TxtChange.Visible = True
        TxtChange.SetFocus
        End If
    End If
End Sub

Private Sub Label_DblClick()
    Focused = 2
    RaiseEvent DblClick
End Sub

Private Sub Label_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    UserControl_EnterFocus
End Sub

Private Sub Timer1_Timer()
    Select Case msgtype
    Case 1
        If ImgMsg.Visible = True Then
            ImgMsg.Visible = False
        Else
            ImgMsg.Visible = True
        End If
    Case 2
        If ImgErr.Visible = True Then
            ImgErr.Visible = False
        Else
            ImgErr.Visible = True
        End If
    End Select
    
    BlinkCount = BlinkCount + 1
    If BlinkCount > 30 Then
        ImgMsg.Visible = False
        ImgErr.Visible = False
        BlinkCount = 0
        Timer1.Enabled = False
        RaiseEvent MessageOver(msgtype)
        msgtype = 0
    End If
End Sub

Private Sub TxtChange_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then ChangeText
End Sub

Private Sub TxtChange_LostFocus()
ChangeText
End Sub

Private Sub ChangeText()
    Label.Caption = TxtChange.Text
    TxtChange.Visible = False
    Label.Visible = True
    RaiseEvent Change
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    Label.BackColor = RGB(206, 206, 206)
    Label.ForeColor = vbBlack
End Sub

Public Sub EnterFocus()
    UserControl_EnterFocus
End Sub

Public Sub ExitFocus()
    UserControl_ExitFocus
    If TxtChange.Visible = True Then ChangeText
End Sub

Private Sub UserControl_ExitFocus()
    Focused = 0
    Label.BackColor = RGB(240, 240, 240)
    Label.ForeColor = RGB(120, 120, 120)
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    UserControl_EnterFocus
End Sub

Sub Pause(Interval)
'Pauses for a given time
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(Interval)
        DoEvents
    Loop
End Sub

Sub SetState(stat)
    Select Case stat
        Case 1 'message
            msgtype = 1
        Case 2 'warning
            msgtype = 2
    End Select
    BlinkCount = 0
    Timer1.Enabled = True
End Sub
