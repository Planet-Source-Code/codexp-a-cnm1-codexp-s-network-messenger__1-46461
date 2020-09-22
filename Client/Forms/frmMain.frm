VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CNM"
   ClientHeight    =   4245
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "CNMClient"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdConnect"
            Object.ToolTipText     =   " Connect to Server "
            Object.Tag             =   "connect"
            ImageKey        =   "connect"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdJoin"
            Object.ToolTipText     =   " Join a Channel "
            Object.Tag             =   "join disabled"
            ImageKey        =   "join disabled"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOpenConsole"
            Object.ToolTipText     =   " Open Console "
            Object.Tag             =   "console"
            ImageKey        =   "console"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3990
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "00:00:00"
            TextSave        =   "00:00:00"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   0
      Left            =   600
      ScaleHeight     =   2145
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      Begin MSComctlLib.ImageList imlBuddy 
         Left            =   480
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0A02
               Key             =   "standard"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0F9C
               Key             =   "me"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   120
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1536
               Key             =   "connect"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":188A
               Key             =   "connect disabled"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1BDE
               Key             =   "disconnect"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F32
               Key             =   "join"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2286
               Key             =   "join disabled"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25DA
               Key             =   "console"
            EndProperty
         EndProperty
      End
      Begin VB.Timer tFastControl 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   600
         Top             =   600
      End
      Begin VB.Timer tAppControl 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   600
         Top             =   120
      End
      Begin VB.Timer tWSControl 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   600
      End
      Begin MSWinsockLib.Winsock WS 
         Left            =   120
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.PictureBox picBuddys 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   1935
      TabIndex        =   3
      Top             =   480
      Width           =   1935
      Begin MSComctlLib.TreeView tvUsers 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   5953
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   0
         LabelEdit       =   1
         Style           =   1
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "imlBuddy"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Line shpLine 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   0
      X2              =   2040
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line shpLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   0
      X2              =   2040
      Y1              =   3975
      Y2              =   3975
   End
   Begin VB.Line shpLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   2040
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line shpLine 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   2040
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Menu menuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenConsole 
         Caption         =   "&Open Console"
      End
      Begin VB.Menu mnuOpenPrivate 
         Caption         =   "&Message to..."
      End
      Begin VB.Menu mnuJoinChannel 
         Caption         =   "&Join Channel..."
      End
      Begin VB.Menu mnuLine12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuConnection 
         Caption         =   "&Connection..."
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu mnuRegistration 
         Caption         =   "&Register..."
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu menuUser 
      Caption         =   "User Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuUserInfo 
         Caption         =   "User Info..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************'
'  TO DO:                     '
' - History                   '
' - Plugins                   '
'*****************************'

Private bFormMoved    As Boolean



Public Sub ConnectToServer(Optional ByVal OpenConnections As Boolean = True, Optional ByVal JustConnect As Boolean)
  If Not JustConnect And Trim(Client.User) = "" Then
    If OpenConnections Then mnuConnection_Click
    Exit Sub
  End If
  If Trim(LastServer.Host) = "" Then
    If OpenConnections Then mnuConnection_Click
    Exit Sub
  End If
  If LastServer.Port < 1 Then LastServer.Port = 8888
  CurrentServer = LastServer
  If WS.State Then WS.Close
  If tWSControl.Enabled Then tWSControl_Timer
  If bDontAutoLogin Then
    bAutoLogin = False
    bDontAutoLogin = False
  Else
    bAutoLogin = Client.AutoLogin
  End If
  WS.Connect CurrentServer.Host, CurrentServer.Port
  AutoConnectEnabled = True
End Sub

Private Sub Form_Load()
  If UCase(Trim(Command)) <> "MULTIUSE" Then
    If App.PrevInstance Then
      Unload Me
      Exit Sub
    End If
  End If
  Me.Show
  Init
  EnableTimers
  Me.Move Screen.Width - Me.Width, 0
  If Application.OpenConMgr Then mnuConnection_Click
End Sub

Public Sub ToggleConnection()
  If WS.State = sckConnected Then
    Disconnect
  Else
    ConnectToServer
  End If
End Sub

Private Sub Form_Resize()
  shpLine(2).Y1 = Me.ScaleHeight - sbStatus.Height - 15
  shpLine(2).Y2 = Me.ScaleHeight - sbStatus.Height - 15
  shpLine(3).Y1 = Me.ScaleHeight - sbStatus.Height - 30
  shpLine(3).Y2 = Me.ScaleHeight - sbStatus.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Disconnect
  DisableTimers
  CleanUp
  End
End Sub

Public Sub Disconnect()
  AutoConnectEnabled = False
  If Client.LoggedIn > 0 Then
    SendCL ":" & Client.User & " LOGOUT"
  End If
  DoEvents
  If WS.State Then WS.Close
  If tWSControl.Enabled Then tWSControl_Timer
End Sub

Private Sub mnuAbout_Click()
  Static bExit As Boolean
  Dim Tmp As String
  If bExit Then Exit Sub
  bExit = True
  Tmp = Application.Caption & " " & _
        Application.Version & vbCrLf & _
        Application.Copyright
  MsgInformation Tmp, "About"
  bExit = False
End Sub

Private Sub mnuConnect_Click()
  If WS.State = sckConnected Then
    Disconnect
  Else
    ConnectToServer
  End If
End Sub

Private Sub mnuConnection_Click()
  Dim sX As Single
  Dim sY As Single
  If frmConnect Is Nothing Then Load frmConnect
  sY = Me.Top
  sX = Me.Left - frmConnect.Width
  If sX < 0 Then
    sX = Me.Left + Me.Width
  End If
  frmConnect.Move sX, sY
  If frmConnect.Visible Then
    frmConnect.SetFocus
  Else
    frmConnect.Show
  End If
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub EnableTimers()
  tAppControl.Enabled = True
  tFastControl.Enabled = True
  tWSControl.Enabled = True
End Sub

Private Sub DisableTimers()
  tAppControl.Enabled = False
  tFastControl.Enabled = False
  tWSControl.Enabled = False
End Sub

Private Sub mnuJoinChannel_Click()
  Static bExit As Boolean
  Static fInp As New frmInputBox
  Dim Chan As String
  
  If fInp Is Nothing Then Set fInp = New frmInputBox
  If fInp.Visible Then fInp.SetFocus
  If Client.LoggedIn = 0 Or bExit Then Exit Sub
  bExit = True
  
  fInp.Caption = "Input Channel:"
  fInp.lblCaption = "Type Channel name here to Join it:"
  
  fInp.Show
  Do While fInp.Visible: DoEvents: Loop
  bExit = False
  
  Chan = Trim(fInp.txtInput)
  If Len(Chan) Then SendCL ":" & Client.User & " JOIN " & Chan
  Unload fInp
End Sub

Private Sub mnuOpenConsole_Click()
  If Not fConsole.Visible Then fConsole.Show
End Sub

Private Sub mnuOpenPrivate_Click()
  Static bExit As Boolean
  Static fInp As New frmInputBox
  Dim User As String
  
  If fInp Is Nothing Then Set fInp = New frmInputBox
  If fInp.Visible Then fInp.SetFocus
  If Client.LoggedIn = 0 Or bExit Then Exit Sub
  bExit = True
  
  fInp.Caption = "Input User ID:"
  fInp.lblCaption = "Type User ID here to send message to:"
  
  fInp.Show
  Do While fInp.Visible: DoEvents: Loop
  bExit = False
  
  User = Trim(fInp.txtInput)
  If Len(User) Then OpenPrivate User
  
  Unload fInp
End Sub

Private Sub mnuRegistration_Click()
  ShowRegistrationForm
End Sub

Private Sub mnuUserInfo_Click()
  Dim NNode As Node
  Set NNode = tvUsers.SelectedItem
  If NNode Is Nothing Or Client.Connected = 0 Then Exit Sub
  ShowRegistrationForm
  frmRegForm.ClearForm
  SendCL "QUERY USER " & NNode.Key & " INFO"
End Sub

Private Sub tAppControl_Timer()
  Static bFirst As Boolean
  Static Tmr As Long
  Dim Tmp As String
  
  ' Setup "Connect" Menu Caption  '
  If WS.State = 7 Then
    Tmp = "Dis&connect"
  Else
    Tmp = "&Connect"
  End If
  If mnuConnect.Caption <> Tmp Then
    mnuConnect.Caption = Tmp
  End If
  
  If AutoConnectEnabled Then
    If Client.Connected = 0 And Client.AutoConnect Then
      If Tmr = 0 Then Tmr = TickSeconds + 5
      If Tmr < TickSeconds Or Not bFirst Then
        bFirst = True
        ConnectToServer False
        Tmr = 0
      End If
    End If
  End If
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "cmdConnect"
      ToggleConnection
    Case "cmdJoin"
      mnuJoinChannel_Click
    Case "cmdOpenConsole"
      mnuOpenConsole_Click
  End Select
End Sub

Private Sub tFastControl_Timer()
  Static fX As Single
  Static fY As Single
  Dim Tmp As String
  Dim TmpX As Single
  Dim lTmp As Long
  
  ' Setup WinSock State in window Caption '
  Tmp = "CNM " & Version
  If Me.Caption <> Tmp Then
    Me.Caption = Tmp
  End If
  ' Setup UserID in Status Bar '
  Tmp = " " & Client.User
  If sbStatus.Panels(2).Text <> Tmp Then
    sbStatus.Panels(2).Text = Tmp
  End If
  ' Setup Connected Time in Status Bar '
  If Client.Connected > 0 Then
    Tmp = TimeLeftAfter(Client.Connected)
  Else
    Tmp = "0:00:00"
  End If
  If sbStatus.Panels(1).Text <> Tmp Then
    sbStatus.Panels(1).Text = Tmp
  End If
  
  ' Setup Connect Button Image  '
  If tbToolbar.Buttons("cmdConnect").Tag <> CStr(WS.State) Then
    Select Case WS.State
      Case 7: tbToolbar.Buttons("cmdConnect").Image = "disconnect"
      Case 0: tbToolbar.Buttons("cmdConnect").Image = "connect"
      Case Else: tbToolbar.Buttons("cmdConnect").Image = "connect disabled"
    End Select
    tbToolbar.Buttons("cmdConnect").Tag = CStr(WS.State)
  End If
  ' Setup Join Channel Button Image  '
  lTmp = Client.LoggedIn > 0
  If tbToolbar.Buttons("cmdJoin").Tag <> CStr(lTmp) Then
    If Client.LoggedIn > 0 Then
      tbToolbar.Buttons("cmdJoin").Image = "join"
    Else
      tbToolbar.Buttons("cmdJoin").Image = "join disabled"
    End If
    tbToolbar.Buttons("cmdJoin").Tag = CStr(lTmp)
  End If
  
  ' Dock Log Window '
  If Me.Left <> fX Or Me.Top <> fY Then bFormMoved = True
  If bFormMoved Then
    If Application.ShowLogger Then
      If frmLogViewer.Visible Then
        If frmLogViewer.mnuDock.Checked Then
          frmLogViewer.Top = Me.Top + Me.Height
          TmpX = Me.Left
          If TmpX > Screen.Width - frmLogViewer.Width Then
            TmpX = Screen.Width - frmLogViewer.Width
          End If
          frmLogViewer.Left = TmpX
        End If
      End If
    End If
    bFormMoved = False
  End If
  
  fX = Me.Left
  fY = Me.Top
End Sub

Private Sub tvUsers_DblClick()
  If Not tvUsers.SelectedItem Is Nothing Then
    OpenPrivate tvUsers.SelectedItem.Key
  End If
End Sub

Private Sub tvUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim NNode As Node
  If Button = vbRightButton Then
    Set NNode = tvUsers.SelectedItem
    If Not NNode Is Nothing Then
      PopupMenu menuUser
    End If
  End If
End Sub

Private Sub tWSControl_Timer()
  Select Case WS.State
    Case sckConnecting, sckConnectionPending, sckResolvingHost ' do nothing '
    Case sckConnected
      If Client.Connected = 0 Then
        ' Event: Client Connected to Server '
        Client.Connected = TimeLong
        Client.LoggedIn = 0
        Client.DataBuffer = ""
        ' Send Login Informations '
        If Trim(Client.User) <> "" And bAutoLogin Then
          If Trim(Client.Password) <> "" Then
            SendCL "PASS " & Client.Password
          End If
          SendCL "USER " & Client.User
        End If
        EventRaised "Connected to Server"
      End If
    Case sckClosed
      If Client.Connected > 0 Then
        ' Event: Client Disconnected from Server  '
        Client.Connected = 0
        Client.LoggedIn = 0
        Client.DataBuffer = ""
        NumUsers = 0
        UpdateUserList
        EventRaised "Disconnected from Server"
      End If
      If Client.LoggedIn > 0 Then Client.LoggedIn = 0
    Case Else
      WS.Close
      If tWSControl.Enabled Then tWSControl_Timer
  End Select
End Sub

Public Sub SendCL(ByVal CommandLine As String, Optional PrependUserID As Boolean)
  Dim i As Long
  Dim Tmp As String
  If WS.State = 7 And Len(Trim(CommandLine)) Then
    If PrependUserID And Len(Trim(Client.User)) Then Tmp = ":" & Trim(Client.User)
    WS.SendData Tmp & CommandLine & vbCrLf
    For i = 0 To 10: DoEvents: Next i
  End If
End Sub

Private Sub WS_Close()
  If tWSControl.Enabled Then tWSControl_Timer
End Sub

Private Sub WS_Connect()
  If tWSControl.Enabled Then tWSControl_Timer
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
  Dim Buffer  As String
  Dim Temp    As String
  Dim n       As Long
  Dim i       As Long
  
  ' Get Buffer and new Data '
  Buffer = Client.DataBuffer
  WS.GetData Temp
  
  ' add buffer if exists  '
  If Len(Buffer) Then Temp = Buffer & Temp
  n = DelimiterCount(Temp, vbCrLf, True) + 1
  Buffer = GetToken(Temp, vbCrLf, n, True) ' save last line to buffer  '
  
  ' Get each line delimited by CrLf (except the last line)  '
  For i = 1 To n - 1
    Call ExecuteRemoteCommand(GetToken(Temp, vbCrLf, i, True))
  Next i
  
  Client.DataBuffer = Buffer
End Sub

Public Sub ExecuteRemoteCommand(ByVal CommandLine As String)
  ' expected: CommandLine String (single string line without CrLf)  '
  ' CommandLine Syntax: (IRC norm)                                  '
  ' [:user] command [param1] [paramX] [...][:Message String]        '
  ' CommandLine Rules:                                              '
  ' 1. If the first character of the CommandLine is a colon ':'     '
  '    then the first word is the user name.                        '
  ' 2. In the CommandLine a Message String begins after the first   '
  '    colon ':' but if it's not the first character in the         '
  '    CommandLine, in other case begins it after second colon.     '
  ' 3. The Command is the first word in CommandLine without leading '
  '    colon ':', in other case is this second word.                '
  ' 4. Each word before Message String and after Command is a       '
  '    Parameter.                                                   '
  ' 5. Command, Parameter and User are alphanumeric strings without '
  '    space like TAB or WHITE.                                     '
  ' 6. User begins with colon ':' without spaces between them.      '
  Dim Original As String
  Dim Params(10) As String
  Dim User As String
  Dim Msg As String
  Dim Cmd As String
  Dim Tmp As String
  Dim TmpA As String
  Dim TmpB As String
  Dim lTmpA As String
  Dim lTmpB As String
  Dim i As Long
  
  If Application.ShowLogger Then
    frmLogViewer.AddLogMessage CommandLine
  End If
  Original = CommandLine  ' ;) '
  CommandLine = LTrim(CommandLine)
  ' Get and remove Message from CommandLine '
  If Left(CommandLine, 1) = ":" Then i = 1
  Msg = GetToken(CommandLine, ":", 2 + i, True, 2 + i)
  CommandLine = GetToken(CommandLine, ":", 1 + i, True, 2 + i)
  ' Get and remove User if exists  '
  If i Then
    User = GetToken(CommandLine, " ", 1, , 2)
    CommandLine = GetToken(CommandLine, " ", 2, , 2)
  End If
  ' Get and remove Command  '
  Cmd = GetToken(CommandLine, " ", 1, , 2)
  CommandLine = GetToken(CommandLine, " ", 2, , 2)
  ' Get 10 Params (if exists) '
  Params(0) = CommandLine
  For i = 1 To DelimiterCount(CommandLine, " ") + 1
    Params(i) = GetToken(CommandLine, " ", i)
  Next i
  
  ' Select Command  '
  Select Case UCase(Cmd)
    Case "QUERY"
      Select Case UCase(Params(1))
        Case "USER"
          Select Case UCase(Params(3))
            Case "USERID", "NICKNAME", "USERNAME", "ADDRESS1", _
                 "ADDRESS2", "ADDRESS3", "EMAIL", "MSNID", _
                 "PHONE", "GENDER", "BDATE", "ICQN"
              If UCase(Params(4)) = "INFO" Then ShowRegistrationForm
          End Select
          Select Case UCase(Params(3))
            Case "ERROR"
              MsgInformation Msg, "Query Error:", True
            Case "USERID"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtUserID = Msg
              Else
                
                'userlist()
              End If
            Case "NICKNAME"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtNick = Msg
              Else
              End If
            Case "USERNAME"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtUserName = Msg
              Else
              End If
            Case "ADDRESS1"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtAddress1 = Msg
              Else
              End If
            Case "ADDRESS2"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtAddress2 = Msg
              Else
              End If
            Case "ADDRESS3"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtAddress3 = Msg
              Else
              End If
            Case "EMAIL"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtEMail = Msg
              Else
              End If
            Case "MSNID"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtMSN = Msg
              Else
              End If
            Case "PHONE"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtPhone = Msg
              Else
              End If
            Case "GENDER"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.SetGenderValue (Val(Msg))
              Else
              End If
            Case "BDATE"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtBDate = Msg
              Else
              End If
            Case "ICQN"
              If UCase(Params(4)) = "INFO" Then
                frmRegForm.txtICQN = Msg
              Else
              End If
            Case Else
              Debug.Print "NOT IMPLEMENTED: QUERY USER " & Params(3)
          End Select
      End Select
      
    Case "PRIVMSG"
      If Left(Params(1), 1) = "#" Then
        TmpA = Trim(Mid(Params(1), 2))
        If Channels.Exist(TmpA) Then
          Channels(TmpA).OpenWindow
          Channels(TmpA).Window.AddMessage "<" & User & "> " & Msg
        End If
      Else
        If Len(User) Then
          lTmpA = OpenPrivate(User)
          If lTmpA Then
            i = lTmpA - 1
            fPrivate(i).AddMessage "<" & User & "> " & Msg
          End If
        Else
          fConsole.AddMessage "<Unknown> " & Msg
        End If
      End If
      
    Case "LOGIN"
      If Len(Params(1)) Then
        If UCase(Params(1)) = UCase(Client.User) Then
          ' Event: User logged in '
          Client.LoggedIn = TimeLong
          EventRaised "Login successfuly!"
          SendCL ":" & Client.User & " USERS"
        End If
        AddUserToList Params(1), Val(Params(2))
      End If
    
    Case "LOGOUT"
      If Len(Params(1)) Then
        If UCase(Params(1)) = UCase(Client.User) Then
          Client.LoggedIn = 0
        End If
        If UCase(Params(1)) = UCase(Trim(Client.User)) Then
          RemoveAllUsersFromList
        Else
          RemoveUserFromList Params(1)
        End If
      End If
      
    Case "JOIN"
      TmpA = Trim(Params(1))
      If Left(TmpA, 1) = "#" Then TmpA = Trim(Mid(TmpA, 2))
      If Len(TmpA) Then
        If Not Channels.Exist(TmpA) Then
          If UCase(User) = UCase(Client.User) Then
            If Channels.AddAs(TmpA, User) = 0 Then
              Channels(TmpA).OpenWindow
            End If
          End If
        End If
        If Channels.Exist(TmpA) Then
          Channels(TmpA).Window.AddUser User
          Channels(TmpA).Window.AddMessage User & " has joined #" & Channels(TmpA).Caption
        End If
      End If
    
    Case "PART"
      TmpA = Trim(Params(1))
      If Left(TmpA, 1) = "#" Then TmpA = Trim(Mid(TmpA, 2))
      If Channels.Exist(TmpA) Then
        Channels(TmpA).Window.RemoveUser User
        Channels(TmpA).Window.AddMessage User & " has left #" & Channels(TmpA).Caption
        If UCase(User) = UCase(Client.User) Then
          Set Channels(TmpA).Window.Chan = Nothing
          Channels.Remove TmpA
        End If
      End If
    
    Case "USERS"
      TmpA = Params(1)
      If Left(TmpA, 1) = "#" Then TmpA = Trim(Mid(TmpA, 2))
      If Len(TmpA) Then
        If Not Channels.Exist(TmpA) Then
          Channels.AddAs TmpA, Client.User
          If Channels.Exist(TmpA) Then
            Channels(TmpA).Window.txtLine.Enabled = False
          End If
        End If
        If Channels.Exist(TmpA) Then
          Channels(TmpA).Window.ParseUserList Msg
        End If
      Else
        SplitUsers Msg
      End If
      
    Case "REG"
      Select Case UCase(Params(1))
        Case "DONE"
          MsgInformation Msg, "Registration done!"
          lTmpA = MsgQuestion("Do you want to apply new Account?", "Apply new Account?")
          If lTmpA = 2 Then ' YES '
            If Len(frmRegForm.txtUserID) And _
               Len(frmRegForm.txtPassword) Then
              Client.User = frmRegForm.txtUserID
              Client.Password = frmRegForm.txtPassword
            End If
          End If
        
        Case "UPDATE"
          Select Case UCase(Params(2))
            Case "DONE"
              MsgInformation Msg, "Update done!"
              If Len(frmRegForm.txtUserID) And _
                 Len(frmRegForm.txtPassword) Then
                Client.User = frmRegForm.txtUserID
                Client.Password = frmRegForm.txtPassword
              End If
          End Select
      End Select
    
    Case "0", "000" ' Server Command failed  '
      fConsole.AddMessage "*** ERROR: " & Params(0) & ":" & Msg
      Select Case UCase(Params(1))
        Case "REG"
          Select Case UCase(Params(2))
            Case "USERID", "PASSWORD", "NICKNAME"
              MsgInformation "Registration Error in Field " & Params(2) & ":" & vbCrLf & Msg
          End Select
        Case "PRIVMSG"
          Select Case UCase(Params(2))
            Case "CHANNEL"
              If Len(Params(3)) Then
                MsgInformation "Server sent this Error Message:" & vbCrLf & Msg, "Error!", True
              End If
          End Select
        Case "JOIN" ' Message Failed '
          Select Case UCase(Params(2))
            Case "JOINED"
              TmpA = Params(3)
              If Left(TmpA, 1) = "#" Then TmpA = Trim(Mid(TmpA, 2))
              If Channels.Exist(TmpA) Then
                Channels(TmpA).OpenWindow
              End If
            Case "INVALID"
              TmpB = "Your desired Channel Name is invalid!" & vbCrLf
              TmpB = TmpB & "Please try to join another Channel!"
              MsgInformation TmpB, "Invalid Channel!", True
            Case "ERROR"
              TmpB = "Server cant create this Channel!" & vbCrLf
              TmpB = TmpB & "Reason is unknown. :("
              MsgInformation TmpB, "Server Error:", True
            Case "LIMIT"
              TmpB = "You can't create any more Channels!" & vbCrLf
              TmpB = TmpB & "You may Join maximal 10 Channels!"
              MsgInformation TmpB, "Limit reached!", True
          End Select
        Case "LOGIN"  ' Login failed  '
          PlaySound "ERROR"
          Disconnect
          EventRaised "Login failed! Please check Your Account!", "ERROR"
          MsgInformation "Your Login was rejected by Server!" & vbCrLf & _
                         "Error Message from Server:" & vbCrLf & _
                         Msg, "Login Error!", True
          frmConnect.tmrBlink.Enabled = True
          mnuConnection_Click
        Case "USER" ' User ID is Invalid  '
          PlaySound "ERROR"
          If UCase(Params(2)) <> "DENY" Then
            Disconnect
            EventRaised "User ID not accepted! Please check User ID!", "ERROR"
            MsgInformation "Your User ID was not accepted!" & vbCrLf & _
                           "Error Message from Server:" & vbCrLf & _
                           Msg, "User ID Error!", True
            frmConnect.tmrBlink.Enabled = True
            mnuConnection_Click
          End If
        Case "JOIN"
          Select Case UCase(Params(2))
            Case "JOINED"
          End Select
        Case Else
          LogMessage "Server Error: " & Msg, "ERROR"
      End Select
    
    Case "1", "001" ' Server Command done '
      Select Case UCase(Params(1))
        Case "PRIVMSG"
        Case "JOIN"
        Case "LOGIN"
        Case "USER" ' User ID is Ok!  '
          SendCL "LOGIN"
        Case Else
          If Len(User) Then
            TmpA = User & ": "
          Else
            TmpA = "MSG: "
          End If
          LogMessage TmpA & Msg
      End Select
      
  End Select
End Sub

Private Sub AddUserToList(ByVal User As String, ByVal Rights As Long)
  Dim i As Long
  If Trim(User) = "" Then Exit Sub
  For i = 0 To NumUsers - 1
    If UCase(Trim(UserList(i).User)) = UCase(Trim(User)) Then
      ' Don't add existing user '
      Exit Sub
    End If
  Next i
  NumUsers = NumUsers + 1
  UserList(NumUsers - 1) = EmptyUser
  UserList(NumUsers - 1).User = Trim(User)
  UpdateUserList
End Sub

Private Sub RemoveUserFromList(ByVal User As String)
  Dim i As Long
  For i = 0 To NumUsers - 1
    If UCase(Trim(UserList(i).User)) = UCase(Trim(User)) Then
      UserList(i) = UserList(NumUsers - 1)
      NumUsers = NumUsers - 1
      Exit For
    End If
  Next i
  UpdateUserList
End Sub

Private Sub RemoveAllUsersFromList()
  NumUsers = 0
  UpdateUserList
End Sub

Private Sub SplitUsers(ByVal Users As String)
  'Dim cRights As Long
  Dim CUser As String
  Dim i As Long
  
  NumUsers = 0
  For i = 1 To DelimiterCount(Users, ",") + 1
    CUser = Trim(GetToken(GetToken(Users, ",", i), "|", 1))
    'cRights = Val(GetToken(GetToken(Users, ",", i), "|", 2))
    If Len(CUser) Then
      NumUsers = NumUsers + 1
      UserList(NumUsers - 1) = EmptyUser
      UserList(NumUsers - 1).User = CUser
    End If
  Next i
  
  UpdateUserList
End Sub

Private Sub UpdateUserList()
  Dim i As Long
  
  tvUsers.Nodes.Clear
  For i = 0 To NumUsers - 1
    Set UserList(i).NNode = tvUsers.Nodes.Add(, , UserList(i).User, UserList(i).User)
    UserList(i).NNode.Tag = i
    UserList(i).NNode.ForeColor = vbBlue
    If UCase(UserList(i).User) = UCase(Client.User) Then
      UserList(i).NNode.Image = "me"
    Else
      UserList(i).NNode.Image = "standard"
    End If
  Next i
  
End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  If tWSControl.Enabled Then tWSControl_Timer
End Sub
