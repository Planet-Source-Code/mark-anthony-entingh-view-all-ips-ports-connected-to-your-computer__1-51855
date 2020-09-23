VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl DX 
   BackColor       =   &H00F0F0F0&
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   Begin VB.PictureBox PicHide 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   3930
      Left            =   0
      ScaleHeight     =   262
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   566
      TabIndex        =   4
      Top             =   0
      Width           =   8490
      Begin VB.PictureBox PicHideContainer 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   240
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   546
         TabIndex        =   5
         Top             =   1395
         Width           =   8190
         Begin VB.PictureBox PicMeterContainer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   3285
            ScaleHeight     =   6
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   98
            TabIndex        =   6
            Top             =   225
            Width           =   1500
            Begin VB.PictureBox PicMeter 
               BackColor       =   &H00B1B1B1&
               BorderStyle     =   0  'None
               Height          =   180
               Left            =   -15
               ScaleHeight     =   180
               ScaleWidth      =   15
               TabIndex        =   7
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.Label LblHide 
            BackStyle       =   0  'Transparent
            Caption         =   "please wait..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3585
            TabIndex        =   9
            Top             =   0
            Width           =   1020
         End
         Begin VB.Label lblMeter 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3285
            TabIndex        =   8
            Top             =   360
            Width           =   4095
         End
      End
   End
   Begin VB.Timer TmrActPorts 
      Interval        =   5000
      Left            =   3180
      Top             =   2325
   End
   Begin VB.HScrollBar HScroll1 
      CausesValidation=   0   'False
      Height          =   240
      Left            =   0
      Max             =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4065
      Width           =   8595
   End
   Begin MSComctlLib.ImageList Image1 
      Left            =   3015
      Top             =   1605
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":32E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":5A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":824A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":869C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":AE4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":D600
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":FDB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":155A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":1593E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":16408
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DX.ctx":167A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Conn 
      Index           =   0
      Left            =   3885
      Top             =   1770
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Container 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   0
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   514
      TabIndex        =   1
      Top             =   0
      Width           =   7710
      Begin DefenceXplore.File File1 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
      End
      Begin VB.PictureBox Logo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   5490
         Picture         =   "DX.ctx":16B3C
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   104
         TabIndex        =   3
         Top             =   1500
         Width           =   1560
      End
   End
   Begin VB.Menu Menu 
      Caption         =   ".conn"
      Index           =   0
      Begin VB.Menu mnu2 
         Caption         =   "&Send"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu4 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnu5 
         Caption         =   "&Default Name"
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu7 
         Caption         =   "&Connect"
      End
      Begin VB.Menu Mnu8 
         Caption         =   "Ban IP"
      End
      Begin VB.Menu Mnu9 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu10 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "folder"
      Index           =   1
      Begin VB.Menu Mnu11 
         Caption         =   "&Open"
      End
      Begin VB.Menu Mnu12 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu13 
         Caption         =   "&Properties"
      End
   End
End
Attribute VB_Name = "DX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event NavigateComplete(url As String)


Private Type typPortItem
    local_ip As String
    local_port As String
    remote_ip As String
    remote_port As String
    state As String
    type As String
    active As Single
    fileindex As Integer
End Type
Private PortItem() As typPortItem

Private Type typNode
    Label As String
    Parent As Integer
    Filetype As String
    MenuIndex As Integer
    ArrayIndex As Integer
    Features As String
    Icon As Integer
End Type
Private Node() As typNode


Public Title As String, url As String


Private NotDone As Boolean, Initialized As Boolean
Private cWidth As Integer 'container width
Private SelectedIndex

'file positioning variables
Private SpaceY, SpaceX, MaxWidth  'general positioning settings
Private NextY, NextX  'next file position


Private Sub Container_Resize()
If Container.Width > UserControl.ScaleWidth Then
    tmp = Container.Width - UserControl.ScaleWidth
    HScroll1.Max = tmp
    HScroll1.LargeChange = SpaceX
    If HScroll1.Enabled = False Then HScroll1.Enabled = True
Else
    If HScroll1.Enabled = True Then HScroll1.Enabled = False: Container.Left = 0
End If
End Sub

Private Sub File1_Change(Index As Integer)
    Refresh
End Sub

Private Sub File1_DblClick(Index As Integer)
    SelectedIndex = Index
    Call Menu_Click(File1(Index).MenuIndex)
End Sub

Private Sub File1_MessageOver(Index As Integer, msgtype As Variant)
    If msgtype = 2 Then
        File1(Index).Visible = False
        File1(Index).Tag = "0"
        Refresh
    End If
End Sub

Private Sub File1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectedIndex = Index
    If Button = 2 Then
        File1(Index).SetFocus
        PopupMenu Menu(File1(Index).MenuIndex)
    End If
End Sub

Private Sub HScroll1_Change()
    Container.Left = HScroll1.Value * -1
    Logo.Left = UserControl.ScaleWidth - Logo.Width - Container.Left: Logo.Top = UserControl.ScaleHeight - Logo.Height - HScroll1.Height
    DoEvents
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub Menu_Click(Index As Integer)
    Select Case Index
    Case 0
    Case 1
        'open folder
        Mnu11_Click
    End Select
End Sub

Private Sub Mnu11_Click()
    'open
    LoadFolderNode File1(SelectedIndex).Indx
End Sub

Private Sub TmrActPorts_Timer()
    'Timer for refreshing the active ports items
    GetAllActiveLocalPorts (False)
End Sub

Private Sub UserControl_Initialize()
    ReDim Node(0)
    ReDim PortItem(0)
    
    NotDone = False
    Initialized = False
    'Initialized = True
    TmrActPorts.Enabled = False
    
    SpaceY = 3
    SpaceX = 6
    MaxWidth = 250
    NextY = SpaceY
    NextX = SpaceX
    
    File1(0).Top = -100
End Sub

Private Sub UserControl_Paint()
    If Initialized = False Then
        LoadDXDrive
        Navigate "dx://home"
        Initialized = True
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Logo.Left = UserControl.ScaleWidth - Logo.Width - Container.Left: Logo.Top = UserControl.ScaleHeight - Logo.Height - HScroll1.Height
    Container.Height = UserControl.ScaleHeight - HScroll1.Height
    If cWidth < UserControl.ScaleWidth Then
        Container.Width = UserControl.ScaleWidth
    ElseIf Container.Width <> cWidth Then
        Container.Width = cWidth
    End If
    HScroll1.Top = UserControl.ScaleHeight - HScroll1.Height
    HScroll1.Width = UserControl.ScaleWidth
    PicHide.Width = UserControl.ScaleWidth
    PicHide.Height = UserControl.ScaleHeight
    PicHideContainer.Move (PicHide.ScaleWidth - PicHideContainer.Width) / 2, (PicHide.ScaleHeight - PicHideContainer.Height) / 2
    DoEvents
    Refresh True
End Sub


Public Property Let Node_Label(Index, txt)
    txt = Node(Index).Label
End Property

Public Property Set Node_Label(Index, txt)
    Node(Index).Label = txt
End Property

Public Property Set MenuIndex(Index, mnu_indx)
    Node(Index).MenuIndex = mnu_indx
End Property

Public Property Let MenuIndex(Index, mnu_indx)
    mnu_indx = Node(Index).MenuIndex
End Property
Public Function CreateNode(Label, Optional Parent As Integer, Optional MenuIndex, Optional Icon As Integer, Optional ArrayIndex, Optional Filetype As String, Optional Features As String = "")
    '----------------------------------
    'CREATE NODE
    '----------------------------------
    ReDim Preserve Node(UBound(Node) + 1)
    i = UBound(Node)
    'set parent setting
    If IsMissing(Parent) = False Then
        If UBound(Node) >= Parent Then
            Node(i).Parent = Parent
        End If
    End If
    'set icon setting
    If IsMissing(Icon) = False Then
        Node(i).Icon = Icon
    End If
    'set menu index setting
    If IsMissing(MenuIndex) = False Then
        Node(i).MenuIndex = MenuIndex
    End If
    'set array index setting
    If IsMissing(ArrayIndex) = False Then
        Node(i).ArrayIndex = ArrayIndex
    End If
    'set filetype setting
    If IsMissing(Filetype) = False Then
        Node(i).Filetype = Filetype
    End If
    'set features setting
    Node(i).Features = Features
    'set label setting
    Node(i).Label = Label
    
    CreateNode = i
    
End Function


Public Function CreateItem(Label, Optional Index, Optional MenuIndex, Optional Icon As Integer = 12, Optional ArrayIndex = 0, Optional TCPBufferIndex = 0, Optional msgtype As Integer = 0)
    '----------------------------------
    'CREATE FILE OBJECT
    '----------------------------------
    'ArrayIndex = a value indicating what array structure to use
    '   0 = Node()
    '   1 = PortItem()
    '   2 = Favorites()
    '   3 = BannedIP()
    
    'MenuIndex = the index # of the menu to use when right-licking the item
    
    'Icon = the index of the image in the ImageList control named Image1
    
    'TCPBufferIndex = the index of the TCPBuffer used to execute commands to
    '                   the port this item is linked to. For instance, closing
    '                   the specified port using the KillPort(index) function
    On Error Resume Next
    i = -1
    
    For X = 1 To File1.UBound
        If File1(X).Tag = "0" Then
        i = X
        Exit For
        End If
    Next
    
    If i = -1 Then
        Load File1(File1.Count)
        i = File1.UBound
    End If
    File1(i).Caption = Label
    File1(i).Indx = Index
    File1(i).ArrayIndex = ArrayIndex
    File1(i).MenuIndex = MenuIndex
    File1(i).TCPBufferIndex = TCPBufferIndex
    File1(i).ZOrder 0
    File1(i).Tag = ""
    If Icon > 0 Then File1(i).Icon = Image1.ListImages(Icon).ExtractIcon
    
    
    'reposition file
    If SpaceX < File1(i).Width + 6 Then SpaceX = File1(i).Width + 6
    File1(i).Top = NextY
    File1(i).Left = NextX
    If (File1(i).Top + (((2 * File1(i).Height) + (File1(i).Height / 2))) + SpaceY) > Container.ScaleHeight Then
        NextY = SpaceY
        NextX = NextX + SpaceX
        SpaceX = 3
    Else
        NextY = NextY + File1(i).Height + SpaceY
    End If
    tmp = NextX + SpaceX
    cWidth = tmp
    UserControl_Resize
            
            
    File1(i).Visible = True
    If msgtype > 0 Then File1(i).SetState msgtype
        
    
    CreateItem = i
End Function


Sub Refresh(Optional resized As Boolean = False)
    If NotDone = False Then
    SpaceY = 3
    SpaceX = 6
    MaxWidth = 250
    NextY = SpaceY
    NextX = SpaceX
    
    If File1.Count > 1 Then
        For X = 1 To File1.UBound
            DoEvents
            If File1(X).Tag <> "0" Then
                If SpaceX < File1(X).Width + 6 Then SpaceX = File1(X).Width + 6
                If File1(X).Left <> NextX Or File1(X).Top <> NextY Then File1(X).Move NextX, NextY
                If (File1(X).Top + (((2 * File1(X).Height) + (File1(X).Height / 2))) + SpaceY) > Container.ScaleHeight Then
                    NextY = SpaceY
                    NextX = NextX + SpaceX
                    SpaceX = 3
                    If X < File1.UBound Then bla = 1
                Else
                    NextY = NextY + File1(X).Height + SpaceY
                End If
                lastfile = X
            End If
            If PicHide.Visible = True Then PicMeter.Width = Percent(X, File1.UBound)
        Next
    End If
    tmp = File1(lastfile).Left + SpaceX
    cWidth = tmp
    
    If resized = False Then UserControl_Resize
    End If
End Sub

Private Sub GetAllActiveLocalPorts(Optional showHide As Boolean = True)

    Dim lngSize As Long
    Dim lngRetVal As Long
    Dim lngRows As Long
    Dim i As Long, ci As Byte
    Dim lvItem As ListItem
        
    'On Error GoTo fin
    
    If showHide = True Then
        lblMeter.Caption = "searching for open ports"
        PicHide.Visible = True
    End If
    
    NotDone = True
    
    indtab = 0
    '
    ' TCP
    lngSize = 0
    lngRetVal = GetTcpTable(ByVal 0&, lngSize, 0)

    If lngRetVal = ERROR_NOT_SUPPORTED Then
        'MsgBox "IP Helper non supporté !"
        Exit Sub
    End If
    ReDim arrTCPBuffer(0 To lngSize - 1) As Byte
    
    lngRetVal = GetTcpTable(arrTCPBuffer(0), lngSize, 0)
    
    For X = LBound(PortItem) To UBound(PortItem)
        DoEvents
        If PortItem(X).active < 2 Then PortItem(X).active = 0
    Next
    
    If lngRetVal = ERROR_SUCCESS Then
        CopyMem lngRows, arrTCPBuffer(0), 4
        lblMeter = "loading new open TCP ports"
        For i = 1 To lngRows
        PicMeter.Width = Percent(i, lngRows)
            DoEvents
            CopyMem TcpTableRow, arrTCPBuffer(4 + (i - 1) * Len(TcpTableRow)), Len(TcpTableRow)
                With TcpTableRow
                    ci = 0
                    If Left(GetIpFromLong(.dwRemoteAddr), 7) <> "192.168" And GetIpFromLong(.dwRemoteAddr) <> "127.0.0.1" And GetIpFromLong(.dwRemoteAddr) <> "0.0.0.0" And GetIpFromLong(.dwLocalAddr) <> "0.0.0.0" Then
                        For X = LBound(PortItem) To UBound(PortItem)
                            DoEvents
                            If PortItem(X).active < 2 And CStr(PortItem(X).remote_ip) = CStr(GetIpFromLong(.dwRemoteAddr)) And PortItem(X).type = "TCP" Then
                                'item already created
                                ci = 0
                                PortItem(X).active = 1
                                
                                If InStr(1, CStr(PortItem(X).remote_port), CStr(GetTcpPortNumber(.dwRemotePort)) & ", ") < 0 Then
                                    PortItem(X).remote_port = PortItem(X).remote_port & GetTcpPortNumber(.dwRemotePort) & ", "
                                    tmp = File1(PortItem(X).fileindex).Caption
                                    If Left(tmp, Len(PortItem(X).remote_ip)) = CStr(PortItem(X).remote_ip) Then
                                        File1(PortItem(X).fileindex).Caption = CStr(PortItem(X).remote_ip) & ":" & CStr(PortItem(X).remote_port)
                                    End If
                                End If
                                Exit For
                            Else
                                ci = 1
                            End If
                        Next
                        If ci = 1 Then
                            indtab = UBound(PortItem) + 1
                            ReDim Preserve PortItem(indtab)
                            PortItem(indtab).local_ip = GetIpFromLong(.dwLocalAddr)
                            PortItem(indtab).local_port = GetTcpPortNumber(.dwLocalPort)
                            PortItem(indtab).remote_ip = GetIpFromLong(.dwRemoteAddr)
                            PortItem(indtab).remote_port = GetTcpPortNumber(.dwRemotePort)
                            PortItem(indtab).state = GetState(.dwState)
                            PortItem(indtab).type = "TCP"
                            PortItem(indtab).active = 1
                            Index = CreateItem(PortItem(indtab).remote_ip & ":" & PortItem(indtab).remote_port, indtab, 0, 10, 1, i, 1)
                            PortItem(indtab).fileindex = Index
                        End If
                    End If
                End With
                ReDim Preserve TCPBuffer(i)
                TCPBuffer(i) = TcpTableRow
        Next
    End If
    
    
    
    ' UDP
    If 1 = 0 Then
        lngSize = 0
        lngRetVal = GetUdpTable(ByVal 0&, lngSize, 0)
    
        If lngRetVal = ERROR_NOT_SUPPORTED Then
            'MsgBox "IP Helper non supporté !"
            Exit Sub
        End If
        ReDim arrUDPBuffer(0 To lngSize - 1) ' As Byte
        
        lngRetVal = GetUdpTable(arrUDPBuffer(0), lngSize, 0)
    
        If lngRetVal = ERROR_SUCCESS Then
            CopyMem lngRows, arrUDPBuffer(0), 4
            lblMeter = "loading new open UPD ports"
            For i = 1 To lngRows
            PicMeter.Width = Percent(i, lngRows)
                DoEvents
                CopyMem UdpTableRow, arrUDPBuffer(4 + (i - 1) * Len(UdpTableRow)), Len(UdpTableRow)
                If GetIpFromLong(UdpTableRow.dwLocalAddr) <> "0.0.0.0" Then
                    With UdpTableRow
                        ci = 2
                        For X = LBound(PortItem) To UBound(PortItem)
                            DoEvents
                            If PortItem(X).active < 2 And CStr(PortItem(X).local_port) = CStr(GetUdpPortNumber(.dwLocalPort)) And PortItem(X).type = "UDP" Then
                                'item already created
                                PortItem(X).active = 1
                                ci = 2
                                Exit For
                            Else
                                ci = 1
                            End If
                        Next
                        
                        If ci = 1 Then
                            indtab = UBound(PortItem) + 1
                            ReDim Preserve PortItem(indtab)
                            PortItem(indtab).local_ip = GetIpFromLong(.dwLocalAddr)
                            PortItem(indtab).local_port = GetUdpPortNumber(.dwLocalPort)
                            PortItem(indtab).type = "UDP"
                            PortItem(indtab).active = 1
                            Index = CreateItem("UPD - Local Port " & PortItem(indtab).local_port, indtab, 0, 11, 1, i, 1)
                            PortItem(indtab).fileindex = Index
                        End If
                    End With
                End If
            Next
        End If
    End If
    
    For X = 1 To UBound(PortItem)
        DoEvents
        If X > UBound(PortItem) Then Exit For
        If PortItem(X).active = 0 Then
            PortItem(X).active = 3
            File1(PortItem(X).fileindex).SetState 2
            
            'move data over and shrink the array
            For Y = X To UBound(PortItem)
                If Y < UBound(PortItem) Then
                    PortItem(Y).active = PortItem(Y + 1).active
                    PortItem(Y).fileindex = PortItem(Y + 1).fileindex
                    PortItem(Y).local_ip = PortItem(Y + 1).local_ip
                    PortItem(Y).local_port = PortItem(Y + 1).local_port
                    PortItem(Y).remote_ip = PortItem(Y + 1).remote_ip
                    PortItem(Y).remote_port = PortItem(Y + 1).remote_port
                    PortItem(Y).state = PortItem(Y + 1).state
                    PortItem(Y).type = PortItem(Y + 1).type
                Else
                    PortItem(Y).active = 0
                    PortItem(Y).fileindex = 0
                    PortItem(Y).local_ip = ""
                    PortItem(Y).local_port = ""
                    PortItem(Y).remote_ip = ""
                    PortItem(Y).remote_port = ""
                    PortItem(Y).state = ""
                    PortItem(Y).type = ""
                    ReDim Preserve PortItem(Y - 1)
                    X = X - 1
                End If
            Next
        End If
    Next X
    NotDone = False
    If showHide = True Then lblMeter = "refreshing..."
    DoEvents
    Refresh
    DoEvents
    If showHide = True Then PicHide.Visible = False
fin:
NotDone = False
PicHide.Visible = False
Exit Sub
End Sub

Private Function Percent(cur, tot)
    On Error Resume Next
    Percent = Int((100 / tot) * cur)
End Function


Function Navigate(url)
    Dim URL2
    Dim arrurl As Variant
    'search for the folder by label starting with the last level
    'split up levels into array
    If url = "dx://my dx" Or url = "dx://" Then
        LoadFolderNode (0)
    Else
        URL2 = Mid(url, 6)
        URL2 = Replace(URL2, "\", "/")
        If Right(URL2, 1) <> "/" Then URL2 = URL2 & "/"
        arrurl = Split(URL2, "/")
        
        destination = -1
        lastselected = LBound(Node)
        linkedparent = -1
        If arrurl(UBound(arrurl)) = "" Then
        ReDim Preserve arrurl(UBound(arrurl) - 1)
        End If
        
        For X = UBound(arrurl) To LBound(arrurl) Step -1
            found = 0
            For Y = lastselected To UBound(Node)
                If LCase(Node(Y).Label) = LCase(arrurl(X)) Then
                    If linkedparent = -1 Or linkedparent = Y Then
                        If linkedparent = -1 Then destination = Y
                        linkedparent = Node(Y).Parent
                        found = 1
                        lastselected = Y
                    End If
                End If
            Next
            If found = 0 And Y = UBound(Node) And lastselected <> Y Then
                Exit For
            ElseIf found = 0 Then
                X = X - 1
                If X = UBound(arrurl) And UBound(arrurl) > 0 Then linkedparent = -1
            Else
                lastselected = LBound(Node)
                
            End If
        Next X
        If linkedparent = 0 And destination > -1 Then  '0 = root was found
            LoadFolderNode destination
        End If
    End If
err1:
  'file not found
End Function

Function LoadFolderNode(Index)
    PicHide.Visible = True
    
    '-------------------------------------------
    'reset all variables that are now uneeded
    '-------------------------------------------
    For X = 1 To File1.UBound
        File1(X).Tag = "0"
        File1(X).Visible = False
    Next
    
    TmrActPorts.Enabled = False
    ReDim PortItem(0)
    '-------------------------------------------
    '-------------------------------------------
    
    '-------------------------------------------
    'create the folder items from the array of nodes
    '-------------------------------------------
    For X = LBound(Node) To UBound(Node)
        If Node(X).Parent = Index Then
            CreateItem Node(X).Label, X, Node(X).MenuIndex, Node(X).Icon, Node(X).ArrayIndex
        End If
    Next
    '-------------------------------------------
    '-------------------------------------------
    
    
    '-------------------------------------------
    'load all other features depending on
    'what settings are made for that folder
    '-------------------------------------------
    If Node(Index).Features <> "" Then
        
        tmps = Split(Node(Index).Features, ".")
        For X = LBound(tmps) To UBound(tmps)
            Select Case tmps(X)
            Case "0"
                'settings items
            Case "1"
                'all open ports
                GetAllActiveLocalPorts
                TmrActPorts.Enabled = True
            End Select
        Next
    End If
    '-------------------------------------------
    '-------------------------------------------
    currnode = Node(Index).Parent
    url = Node(Index).Label
    Do
        If currnode = 0 Then Exit Do
        url = Node(currnode).Label & "/" & url
        currnode = Node(currnode).Parent
        
        
    Loop
    url = "dx://" & url
    RaiseEvent NavigateComplete(url)
    
    PicHide.Visible = False
End Function


Function LoadDXDrive()
    'load all folders and settings into the node array
    Dim Index(10) As Integer
    'HOME
    Index(0) = CreateNode("Home", 0, 1, 7, 0)
        'SETTINGS
        Index(1) = CreateNode("Settings", Index(0), 1, 13, 0, , "0")
        'Open Ports
        Index(2) = CreateNode("Open Ports", Index(0), 1, 3, 0, , "1")
            'Favorites
            Index(3) = CreateNode("Favorites", Index(2), 1, 14)
End Function
