VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "CONNECTION"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox chkAutoOpen 
      Caption         =   "&Open on Startup"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      ToolTipText     =   " apply and connect to selected server "
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      ToolTipText     =   " apply changes "
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Bac&k"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      ToolTipText     =   " do nothing "
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame fraUser 
      Caption         =   "[ User ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   4695
      Begin VB.CheckBox chkAutoConnect 
         Caption         =   "Auto Connect on Startup"
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox chkAutoLogin 
         Caption         =   "Auto Login on Connect"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login User"
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdChangePass 
         Caption         =   "C&hange Passw."
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdRegistration 
         Caption         =   "Regis&tration"
         Height          =   285
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtUser 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblFreeShortCuts 
         AutoSize        =   -1  'True
         Caption         =   "BDFJQXYZMI"
         Height          =   195
         Left            =   3840
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pass&word:"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Username:"
         Height          =   195
         Index           =   8
         Left            =   330
         TabIndex        =   12
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   1320
      Top             =   3960
   End
   Begin VB.Timer tmrStater 
      Interval        =   300
      Left            =   840
      Top             =   3960
   End
   Begin VB.Frame fraServer 
      Caption         =   "[ Server ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton cmdRemoveServer 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdEditServer 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdNewServer 
         Caption         =   "&New"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox lstGroups 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown-Liste
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox lstServers 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown-Liste
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label cmdGroupsMenu 
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   150
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "               "
         Height          =   195
         Left            =   3240
         TabIndex        =   36
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lblHost 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "                  "
         Height          =   195
         Left            =   3240
         TabIndex        =   35
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Index           =   3
         Left            =   2700
         TabIndex        =   5
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Host:"
         Height          =   195
         Index           =   2
         Left            =   2670
         TabIndex        =   4
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ser&ver:"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   840
         Width           =   540
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Group:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraEdit 
      Caption         =   "[ Edit Server ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1815
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdEditOk 
         Caption         =   "&Save"
         Height          =   375
         Left            =   840
         TabIndex        =   33
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtHost 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3120
         TabIndex        =   30
         Text            =   "8888"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Zentriert
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   32
         Text            =   "8888"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtDescription 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Text            =   "Description"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtGroup 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Text            =   "General"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdEditCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Group:"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ser&ver:"
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   27
         Top             =   840
         Width           =   540
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Host:"
         Height          =   195
         Index           =   5
         Left            =   2670
         TabIndex        =   29
         Top             =   360
         Width           =   390
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Port:"
         Height          =   195
         Index           =   4
         Left            =   2700
         TabIndex        =   31
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.Menu menuGroups 
      Caption         =   "Groups"
      Visible         =   0   'False
      Begin VB.Menu mnuRenameGroup 
         Caption         =   "&Rename Group"
      End
      Begin VB.Menu mnuRemoveGroup 
         Caption         =   "Re&move Group"
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelServer   As cxtServerType
Private SelIndex    As Long


Private Sub ApplySettings()
  Application.OpenConMgr = chkAutoOpen.Value
  Client.AutoConnect = chkAutoConnect.Value
  Client.AutoLogin = chkAutoLogin.Value
End Sub

Private Sub UpdateServerList()
  Dim i As Long
  lstServers.Clear
  For i = 0 To NumServers
    If Trim(Server(i).Group) = "" Then
      Server(i).Group = "General"
    End If
    If Server(i).Port < 1 Then
      Server(i).Port = 8888
    End If
    If UCase(Trim(lstGroups)) = UCase(Trim(Server(i).Group)) Then
      lstServers.AddItem Server(i).Description
    End If
  Next i
  If lstServers.ListCount Then lstServers.ListIndex = 0
  UpdateLabels
End Sub

Private Sub UpdateGroupsList()
  Dim i As Long
  lstGroups.Clear
  lstGroups.AddItem "General"
  For i = 0 To NumServers
    If Not GroupExist(Server(i).Group) Then
      lstGroups.AddItem Server(i).Group
    End If
  Next i
  If lstGroups.ListCount Then lstGroups.ListIndex = 0
  UpdateLabels
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdConnect_Click()
  ApplySettings
  ModifyClient
  LastServer = SelServer
  frmMain.ConnectToServer
  Unload Me
End Sub

Private Sub cmdEditCancel_Click()
  ShowFrame 0
End Sub

Private Sub cmdEditOk_Click()
  Dim i As Long
  If Trim(txtHost) = "" Then
    PlaySound "NOTREADY"
    Exit Sub
  End If
  ' Setup default Values if field is empty  '
  If Val(txtPort) < 1 Then txtPort = "8888"
  If Trim(txtDescription) = "" Then txtDescription = txtHost
  If Trim(txtGroup) = "" Then txtGroup = "General"
  ' New Server ?  '
  If fraEdit.Tag = "NEW" Or SelIndex = 0 Then
    NumServers = NumServers + 1
    SelIndex = NumServers + 1
  End If
  i = SelIndex
  Server(i - 1).Description = Trim(txtDescription)
  Server(i - 1).Host = Trim(txtHost)
  Server(i - 1).Group = Trim(txtGroup)
  Server(i - 1).Port = Val(txtPort)
  UpdateGroupsList
  UpdateServerList
  ActivateServer Server(i - 1)
  ShowFrame 0
End Sub

Private Sub cmdEditServer_Click()
  If SelIndex = 0 Then
    PlaySound "NOTREADY"
    Exit Sub
  End If
  txtDescription = SelServer.Description
  txtGroup = SelServer.Group
  txtHost = SelServer.Host
  txtPort = SelServer.Port
  fraEdit.Tag = "EDIT"
  fraEdit.Caption = "[ Edit Server ]"
  cmdEditOk.Caption = "&Save"
  ShowFrame 1
End Sub

Private Sub cmdGroupsMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PopupMenu menuGroups
End Sub

Private Function ModifyClient()
  If Client.User <> Trim(txtUser) And Client.Connected > 0 Then
    frmMain.Disconnect
  End If
  Client.User = Trim(txtUser)
  Client.Password = Trim(txtPassword)
End Function

Private Sub cmdLogin_Click()
  Static Tmr As Long
  Dim Tmp As String
  
  If Client.Connected > 0 And Client.LoggedIn > 0 Then
    frmMain.SendCL "LOGOUT"
    Exit Sub
  End If
  
  ModifyClient
  If Trim(Client.User) = "" Or Trim(Client.Password) = "" Then
    PlaySound "NOTREADY"
    Exit Sub
  End If
  
  bDontAutoLogin = True
  If Client.Connected = 0 Then frmMain.ConnectToServer
  
  Tmr = TickSeconds + 10
  Do: DoEvents
    If Tmr < TickSeconds Then
      MsgInformation "Can't connect to Server!", "Error", True
      Exit Sub
    End If
  Loop While Client.Connected < 0
  
  Tmr = TickSeconds + 2
  Do: DoEvents
  Loop While Tmr > TickSeconds
  If Client.LoggedIn > 0 Then Exit Sub
  
  If Len(Trim(Client.Password)) Then
    Tmp = "PASS " & Client.Password & vbCrLf
  End If
  Tmp = Tmp & "USER " & Client.User
  frmMain.SendCL Tmp
End Sub

Private Sub cmdNewServer_Click()
  txtDescription = "CNM Server"
  txtGroup = "General"
  txtHost = "localhost"
  txtPort = "8888"
  fraEdit.Tag = "NEW"
  fraEdit.Caption = "[ New Server ]"
  cmdEditOk.Caption = "&Add"
  ShowFrame 1
End Sub

Private Sub ShowFrame(Index As Long)
  fraServer.Visible = Index = 0
  fraEdit.Visible = Index = 1
  cmdOk.Enabled = Index = 0
  cmdConnect.Enabled = Index = 0
End Sub

Private Sub cmdOk_Click()
  ApplySettings
  ModifyClient
  If SelIndex = 0 Then SelServer = EmptyServer
  LastServer = SelServer
  Unload Me
End Sub

Private Sub cmdRegistration_Click()
  ShowRegistrationForm
End Sub

Private Sub cmdRemoveServer_Click()
  If SelIndex = 0 Then
    PlaySound "NOTREADY"
    Exit Sub
  End If
  If MsgQuestion("Do You realy want to remove Server:" & vbCrLf & _
                 Server(SelIndex - 1).Description & vbCrLf & _
                 "from Server list?", "Realy Remove?") = 2 Then
    Server(SelIndex - 1) = Server(NumServers)
    NumServers = NumServers - 1
    If NumServers < -1 Then NumServers = -1
    UpdateGroupsList
    UpdateServerList
  End If
End Sub

Private Sub Form_Load()
  chkAutoConnect.Value = Abs(Client.AutoConnect)
  chkAutoLogin.Value = Abs(Client.AutoLogin)
  chkAutoOpen.Value = Abs(Application.OpenConMgr)
  txtUser = Client.User
  txtPassword = Client.Password
  UpdateGroupsList
  UpdateServerList
  ActivateServer LastServer
End Sub

Private Function GroupExist(ByVal GroupName As String) As Boolean
  Dim i As Long
  If Trim(GroupName) = "" Then GroupName = "General"
  For i = 0 To lstGroups.ListCount - 1
    If UCase(Trim(lstGroups.List(i))) = UCase(Trim(GroupName)) Then
      GroupExist = True
      Exit For
    End If
  Next i
End Function

Private Sub Form_Unload(Cancel As Integer)
  ApplySettings
End Sub

Private Sub lstGroups_Click()
  UpdateServerList
End Sub

Private Sub SearchSelectedServer()
  Dim i As Long
  Dim n As Long
  Dim c As Long
  
  SelIndex = 0
  SelServer = EmptyServer
  For i = 0 To NumServers
    If Trim(Server(i).Group) = "" Then
      Server(i).Group = "General"
    End If
    If UCase(Trim(lstGroups)) = UCase(Trim(Server(i).Group)) Then
      c = c + 1
      If c = lstServers.ListIndex + 1 Then
        SelServer = Server(i)
        SelIndex = i + 1
        Exit For
      End If
    End If
  Next i
End Sub

Private Sub lstServers_Click()
  UpdateLabels
End Sub

Private Sub UpdateLabels()
  SearchSelectedServer
  lblHost = SelServer.Host
  lblPort = SelServer.Port
End Sub

Private Sub mnuRemoveGroup_Click()
  Dim GName As String
  Dim i As Long
  
  GName = lstGroups
  If Trim(GName) = "" Or UCase(Trim(GName)) = "GENERAL" Then
    PlaySound "DENY"
    Exit Sub
  End If
  
  If MsgQuestion("Do You realy want to remove this Group" & vbCrLf & _
                  "with all contents?", "Remove Group?") <> 2 Then
    Exit Sub
  End If
  
  While i <= NumServers
    If UCase(Trim(Server(i).Group)) = UCase(GName) Then
      Server(i) = Server(NumServers)
      Server(NumServers) = EmptyServer
      NumServers = NumServers - 1
      i = i - 1
    End If
    i = i + 1
  Wend
  
  UpdateGroupsList
  UpdateServerList
End Sub

Private Sub mnuRenameGroup_Click()
  Static bExit As Boolean
  Dim fInput As Form
  Dim NewGName As String
  Dim GName As String
  Dim i As Long
  
  If bExit Then Exit Sub
  
  GName = lstGroups
  If Trim(GName) = "" Or UCase(Trim(GName)) = "GENERAL" Then
    PlaySound "DENY"
    Exit Sub
  End If
  
  bExit = True
  Set fInput = New frmInputBox
  fInput.Caption = "Rename Group """ & GName & """ into:"
  fInput.lblCaption = "Enter new Group name:"
  fInput.Show
  Do While fInput.Visible: DoEvents: Loop
  bExit = False
  
  NewGName = Trim(fInput.txtInput)
  If NewGName <> "" Then
    For i = 0 To NumServers
      If UCase(Trim(Server(i).Group)) = UCase(GName) Then
        Server(i).Group = NewGName
      End If
    Next i
  End If
  
  Unload fInput
  UpdateGroupsList
  UpdateServerList
End Sub

Private Sub tmrBlink_Timer()
  If fraUser.ForeColor <> &H40C0& Then
    fraUser.ForeColor = &H40C0&
  Else
    fraUser.ForeColor = &H80000012
  End If
End Sub

Private Sub tmrStater_Timer()
  Dim Tmp As String
  cmdChangePass.Visible = Client.LoggedIn > 0
  If Client.LoggedIn > 0 Then
    Tmp = "&Logout User"
  Else
    Tmp = "&Login User"
  End If
  If cmdLogin.Caption <> Tmp Then cmdLogin.Caption = Tmp
  If Client.LoggedIn > 0 Then
    Tmp = "Update User Data"
  Else
    Tmp = "Registration"
  End If
  If cmdRegistration.Caption <> Tmp Then cmdRegistration.Caption = Tmp
End Sub

Private Sub txtHost_KeyPress(KeyAscii As Integer)
  If InStr("`~!@#$^&*()=+[]{}\|;:'"",<>/?" & Chr(9), Chr(KeyAscii)) Then KeyAscii = 0
  If KeyAscii = 0 Then PlaySound "DENY"
End Sub

Private Sub txtPassword_Change()
  txtPassword.ToolTipText = " " & txtPassword & " "
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 8
    Case Asc("0") To Asc("9")
    Case Else
      PlaySound "DENY"
      KeyAscii = 0
  End Select
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
  If InStr("@#$*+\|;:'"",<>/? " & Chr(9), Chr(KeyAscii)) Then KeyAscii = 0
  If KeyAscii = 0 Then PlaySound "DENY"
End Sub

Private Sub ActivateServer(Srv As cxtServerType)
  Dim i As Long, j As Long
  Dim aDesc As String
  If Trim(Srv.Host) = "" Then
    Exit Sub
  End If
  For i = 0 To NumServers
    If UCase(Trim(Server(i).Host)) = UCase(Trim(Srv.Host)) Then
      If Server(i).Port = LastServer.Port Then
        ActivateGroup Server(i).Group
        For j = 0 To lstServers.ListCount - 1
          aDesc = lstServers.List(j)
          If UCase(Trim(aDesc)) = UCase(Trim(Srv.Description)) Then
            lstServers.ListIndex = j
            Exit For
          End If
        Next j
        Exit For
      End If
    End If
  Next i
End Sub

Private Sub ActivateGroup(ByVal Group As String)
  Dim aGroup As String
  Dim i As Long
  aGroup = lstGroups
  If Trim(Group) = "" Then Group = "General"
  If UCase(Trim(aGroup)) = UCase(Trim(Group)) Then
    Exit Sub
  End If
  For i = 0 To lstGroups.ListCount - 1
    aGroup = lstGroups.List(i)
    If UCase(Trim(aGroup)) = UCase(Trim(Group)) Then
      lstGroups.ListIndex = i
      Exit For
    End If
  Next i
  UpdateServerList
End Sub
