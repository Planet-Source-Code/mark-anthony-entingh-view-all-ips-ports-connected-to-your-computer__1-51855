VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form1 
   Caption         =   "DX - MyVPN6.conn"
   ClientHeight    =   6060
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicStatus 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   10155
      TabIndex        =   14
      Top             =   5775
      Width           =   10155
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   5760
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   10365
      Begin VB.Frame Frame6 
         Height          =   465
         Left            =   9750
         TabIndex        =   8
         Top             =   0
         Width           =   615
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   30
            Picture         =   "FrmMain.frx":0000
            ScaleHeight     =   315
            ScaleWidth      =   555
            TabIndex        =   13
            Top             =   120
            Width           =   555
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   10365
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   0
         TabIndex        =   3
         Top             =   -45
         Width           =   2760
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   2685
         TabIndex        =   2
         Top             =   -45
         Width           =   2760
      End
   End
   Begin VB.Frame Frame5 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   615
      Width           =   10365
      Begin VB.ComboBox Address 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Text            =   "DX://Home/Connections/MyVPN6.conn"
         Top             =   135
         Width           =   8730
      End
      Begin VB.Label btn_Go 
         Caption         =   "Go"
         Height          =   240
         Left            =   9975
         TabIndex        =   7
         Top             =   180
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   195
         Width           =   600
      End
   End
   Begin VB.Frame DXFrame 
      Height          =   4770
      Left            =   0
      TabIndex        =   11
      Top             =   990
      Width           =   10365
      Begin DefenceXplore.DX DX 
         Height          =   5865
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   10345
      End
   End
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   4635
      Left            =   0
      TabIndex        =   9
      Top             =   1125
      Visible         =   0   'False
      Width           =   10365
      ExtentX         =   18283
      ExtentY         =   8176
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isLoaded As Boolean

Private Sub Address_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call btn_Go_Click
End Sub

Private Sub DX_NavigateComplete(url As String)
Address.Text = url
End Sub

Private Sub Form_Load()
    Web1.Visible = False
    DXFrame.Visible = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    DXFrame.Width = Me.ScaleWidth
    DXFrame.Height = Me.ScaleHeight - DXFrame.Top - StatusBar1.Height
    Web1.Width = Me.ScaleWidth
    Web1.Height = Me.ScaleHeight - Web1.Top - StatusBar1.Height
    DX.Width = (DXFrame.Width - 4) * Screen.TwipsPerPixelX
    DX.Height = (DXFrame.Height - 10) * Screen.TwipsPerPixelY
    Address.Width = Me.Width - Address.Left - 875
    btn_Go.Left = Me.Width - 500
    Frame1.Width = Me.ScaleWidth
    Frame2.Width = Me.ScaleWidth
    Frame5.Width = Me.ScaleWidth
    Frame6.Left = (Frame1.Width * Screen.TwipsPerPixelX) - Frame6.Width
    PicStatus.Top = Me.ScaleHeight - PicStatus.Height - 1
    PicStatus.Width = Me.ScaleWidth - 20
End Sub

Private Sub btn_Go_Click()
tmp = LCase(Address.Text)
If Left(tmp, 5) = "dx://" Then
    DXFrame.Visible = True
    DX.Navigate tmp
Else
    'use the microsoft webbrowser component instead
    Web1.Navigate2 tmp
End If
End Sub

Private Sub Web1_NavigateComplete2(ByVal pDisp As Object, url As Variant)
    If isLoaded = False Then
        isLoaded = True
    Else
        Web1.Visible = True
        DXFrame.Visible = False
    End If
    
End Sub
