VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CodeXP's Net Messenger Server"
   ClientHeight    =   3960
   ClientLeft      =   150
   ClientTop       =   435
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
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "CNMServer"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picStats 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   4665
      TabIndex        =   9
      Top             =   240
      Width           =   4695
      Begin VB.CheckBox btnLogger 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "&Logger"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox btnMore 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "&More Details"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdOptions 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   4200
         Picture         =   "frmMain.frx":030A
         Style           =   1  'Grafisch
         TabIndex        =   27
         ToolTipText     =   " Server Options "
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox cmdLockServer 
         Appearance      =   0  '2D
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4200
         Picture         =   "frmMain.frx":074C
         Style           =   1  'Grafisch
         TabIndex        =   25
         ToolTipText     =   " Lock Server GUI "
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes"
         Height          =   195
         Index           =   7
         Left            =   3360
         TabIndex        =   23
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes"
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   22
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "In:"
         Height          =   315
         Index           =   6
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblDataIn 
         Alignment       =   2  'Zentriert
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblDataOut 
         Alignment       =   2  'Zentriert
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2400
         TabIndex        =   19
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Out:"
         Height          =   315
         Index           =   5
         Left            =   1920
         TabIndex        =   18
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Channels:"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblChannelsCount 
         Alignment       =   2  'Zentriert
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1215
         TabIndex        =   16
         Top             =   600
         Width           =   345
      End
      Begin VB.Label lblUpTime 
         Alignment       =   2  'Zentriert
         Caption         =   "000:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2400
         TabIndex        =   15
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Uptime:"
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   14
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblLoggedCount 
         Alignment       =   2  'Zentriert
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1215
         TabIndex        =   13
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblUsersCount 
         Alignment       =   2  'Zentriert
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1215
         TabIndex        =   12
         Top             =   120
         Width           =   345
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Logged in:"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Users:"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox picHolder 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1680
      ScaleHeight     =   1185
      ScaleWidth      =   3105
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Timer tFastControl 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1560
         Top             =   600
      End
      Begin VB.Timer tAppControl 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1560
         Top             =   120
      End
      Begin VB.PictureBox picIconContainer 
         Appearance      =   0  '2D
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2040
         ScaleHeight     =   825
         ScaleWidth      =   945
         TabIndex        =   1
         Top             =   120
         Width           =   975
         Begin VB.PictureBox picIcons 
            Appearance      =   0  '2D
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   120
            Picture         =   "frmMain.frx":0896
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   5
            Top             =   120
            Width           =   270
         End
         Begin VB.PictureBox picIcons 
            Appearance      =   0  '2D
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   480
            Picture         =   "frmMain.frx":09E0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   4
            Top             =   120
            Width           =   270
         End
         Begin VB.PictureBox picIcons 
            Appearance      =   0  '2D
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   480
            Picture         =   "frmMain.frx":0B2A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   3
            Top             =   480
            Width           =   270
         End
         Begin VB.PictureBox picIcons 
            Appearance      =   0  '2D
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   120
            Picture         =   "frmMain.frx":0C74
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   2
            Top             =   480
            Width           =   270
         End
      End
      Begin VB.Timer tWSControl 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   300
         Left            =   1080
         Top             =   600
      End
      Begin VB.Timer tWSSControl 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   600
         Top             =   600
      End
      Begin VB.Timer tWSCControl 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   600
      End
      Begin MSWinsockLib.Winsock WSC 
         Left            =   120
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock WSS 
         Left            =   600
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock WS 
         Index           =   0
         Left            =   1080
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin MSComctlLib.TreeView tvwDetails 
      Height          =   1935
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
   Begin VB.PictureBox picDetails 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2145
      ScaleWidth      =   4665
      TabIndex        =   6
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   45
   End
   Begin VB.Menu menuServer 
      Caption         =   "Main Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuStartStop 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu menuClient 
      Caption         =   "Client Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuClientDisconnect 
         Caption         =   "&Disconnect"
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "Options Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAccountManager 
         Caption         =   "&Account Manager"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
' _________________________________________________________________________________ '
'                                                                                   '
' CodeXP's Net Messenger Server Copyright (C)2003 by CodeXP                         '
' Mailto: CodeXP@Lycos.de (i prefer german language, thx)                           '
' _________________________________________________________________________________ '
'                                                                                   '
' !ATENTION! I've some little troubles with english language (i'm german), for this '
' reason i beg you to forgive me and don't start a panic! Ok? Thank you!            '
' But if you want to help me then you can send me some suggestions for correction.  '
' _________________________________________________________________________________ '
'                                                                                   '
'  Description                                                                      '
'  ===========                                                                      '
'  This Application is a Net Messaging Server which distributes messages to clients.'
'  Net Messaging allows communication between two or more users connected to LAN or '
'  Internet. This application uses the TCP/IP Protocol for data transmitions.       '
'  This application was created on a request of Julian Kyll for the LAN Partys.     '
' _________________________________________________________________________________ '
'                                                                                   '
' Used Controls    (Documentation)                                                  '
' =============                                                                     '
'_*** WinSocks ***__________________________________________________________________'
' [ WSC ]                                                                           '
' Used for server sided conections. For example to connect to other server or for   '
' remote server administrations.                                                    '
'                                                                                   '
' [ WSS ]                                                                           '
' This is the server socket which waits (listening) for connection requests.        '
' If connection request is raised then will a free WS() socket accept this request. '
'                                                                                   '
' [ WS(Index) ]                                                                     '
' Are used for user (client) connections.                                           '
'                                                                                   '
'_*** Timers ***____________________________________________________________________'
' [ tAppControl ]                                                                   '
' This timer is for controling application processes, events and changed states.    '
' Time Interval: slowly (1 Sec)                                                     '
'                                                                                   '
' [ tFastControl ]                                                                  '
' The same as tAppControl but for faster operations like refresh important states.  '
' Time Interval: fast (300 ms)                                                      '
' _________________________________________________________________________________ '
'                                                                                   '
'  Notices                                                                          '
'  =======                                                                          '
'  WSC Socket is currently not used because it's not needed by server.              '
' _________________________________________________________________________________ '
'                                                                                   '


Private AutoStartDone   As Boolean
Private OldAutoReStart  As Long
Private bFormMoved      As Boolean
Private SelectedClient  As Long
Private bRecallUnload   As Boolean

Public WithEvents ServerTrayIcon As clsTrayIcon
Attribute ServerTrayIcon.VB_VarHelpID = -1



' --------------------------------------------------------------------------------- '
' LOCAL COMMAND LINE INTERPRETER                      '
Public Sub ExecuteLocalCommand(CommandLine As String)
  ' to do: Local Command interpreter (implement server console) '
End Sub
' --------------------------------------------------------------------------------- '
' REMOTE COMMAND LINE INTERPRETER                     '
Public Sub ExecuteClientCommand(ByVal CommandLine As String, ByVal Index As Long)
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
  Dim CChanUser As clsChanUser
  Dim cChan As clsChannel
  Dim Original As String
  Dim Params(10) As String
  Dim User As String
  Dim Msg As String
  Dim Cmd As String
  Dim TmpA As String
  Dim TmpB As String
  Dim lTmpA As Long
  Dim lTmpB As Long
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
  If Trim(User) = "" Then User = Client(Index).User
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
    ' QUERY                                                                       '
    Case "QUERY"
      Select Case UCase(Params(1))
        Case "USER"
          ' Registration Commands '
          RegExecute Original, Index
        Case Else
          SendServerError "QUERY NOOP " & Params(1) & ":Unknown Query Operation!", Index
      End Select
    
    ' REG                                                                         '
    Case "REG"
      ' Registration Commands '
      RegExecute Original, Index
      
    ' PING                                                                        '
    Case "PING"
      SendToClient ":Server PONG " & Params(0), Index
    
    ' PONG                                                                        '
    Case "PONG"
      ' Seconds = TimeLong - Params(1) '
    
    ' PART                                                                        '
    Case "PART"
      Channels.PartUser Client(Index), Params(1)
    
    ' JOIN                                                                        '
    Case "JOIN"
      Channels.JoinUser Client(Index), Params(1)
      
    ' PRIVMSG                                                                     '
    Case "PRIVMSG"
      If Client(Index).Loggedin > 0 Then
        TmpA = Trim(Params(1))
        If UCase(TmpA) = "SERVER" Then
          frmLogViewer.AddLogMessage "<" & Client(Index).User & "> " & Msg, mMessage
        Else
          ' Send Message to Param(1) (to User/Channel if # is prepended) '
          If Len(TmpA) Then
            If Left(TmpA, 1) = "#" Then
              TmpA = Trim(Mid(TmpA, 2))
              ' Send Message to Channel '
              lTmpA = 1
              TmpB = ""
              Select Case ChanMsg(User, TmpA, Msg)
                Case 1  ' Empty User  '
                  TmpB = "RECIPIENT:No Sender!"
                Case 2  ' Empty Message '
                  TmpB = "MESSAGE:No Message!"
                Case 3  ' Empty Channelname '
                  TmpB = "CHANNEL:No Channel!"
                Case 4  ' Invalid Channelname '
                  TmpB = "CHANNEL #" & TmpA & ":Invalid Channel!"
                Case 5  ' Channel does not exist '
                  TmpB = "CHANNEL #" & TmpA & ":Channel does not exist!"
                Case Else
                  lTmpA = 0
              End Select
              If lTmpA Then
                SendServerError Cmd & " " & TmpB, Index
              End If
            Else
              ' Send Message to User    '
              lTmpA = 0
              For i = 0 To MaxSockets
                If Client(i).Loggedin > 0 Then
                  If UCase(TmpA) = UCase(Trim(Client(i).User)) Then
                    lTmpA = i + 1
                    Exit For
                  End If
                End If
              Next i
              If lTmpA = 0 Then
              '***************************************************'
              ' To Do:                                            '
              ' 1. If User is not Connected then Save Message in  '
              '    a Database for resend if user goes online.     '
              '***************************************************'
              Else
                i = lTmpA - 1
                SendToClient ":" & User & " PRIVMSG " & TmpA & ":" & Msg, i
              End If
            End If
          Else
            SendServerError Cmd & " RECIPIENT:Wrong Recipient!", Index
          End If
        End If
      Else
        SendServerError Cmd & " DENY:You are not logged in!", Index
      End If
    
    ' USERS                                                                       '
    Case "USERS"
      If Client(Index).Loggedin > 0 Then
        lTmpA = 0
        TmpA = Params(1)
        If Left(TmpA, 1) = "#" Then TmpA = Trim(Mid(TmpA, 2))
        If Len(TmpA) Then
          ' List Users from Channel Param(1)  '
          If Channels.Exist(TmpA) Then
            TmpA = Channels(TmpA).Caption
            SendToClient ":Server USERS #" & TmpA & ":" & Channels(TmpA).MakeUserList, Index
          Else
            SendServerError "USERS CHANNEL:Channel does not exist!", Index
          End If
        Else
          ' List Logged in Users  '
          For i = 0 To MaxSockets
            If Client(i).Loggedin Then
              If Len(TmpA) Then
                TmpA = TmpA & ", " & Client(i).User
              Else
                TmpA = Client(i).User
              End If
              lTmpA = lTmpA + 1
            End If
          Next i
          TmpA = ":Server USERS:" & TmpA
          SendToClient TmpA, Index
        End If
      Else
        SendServerError "USERS DENY:You are not logged in!", Index
      End If
      
    ' PASS                                                                        '
    Case "PASS", "PASSWORD"
      If Client(Index).Loggedin > 0 Then
        SendServerError "PASS DENY:Password not needed!", Index
      Else
        TmpA = Params(1)
        If CheckPassword(TmpA) = 0 Then
          Client(Index).Password = TmpA
          SendToClient ":Server 001 PASS OK:Password Ok!", Index
        Else
          SendServerError "PASS INVALID:Invalid Password!", Index
        End If
      End If
      
    ' USER                                                                        '
    Case "USER"
      If Client(Index).Loggedin = 0 Then
        TmpA = Trim(Params(1))
        If Len(TmpA) Then
          If CheckUserID(TmpA) = 0 Then
            If UserIDIsUsed(TmpA) Then
              SendServerError "USER INUSE:User ID is already in use!", Index
            Else
              Client(Index).User = TmpA
              SendToClient ":Server 001 USER OK:User ID Ok!", Index
            End If
          Else
            SendServerError "USER INVALID:Invalid User ID!", Index
          End If
        Else
          SendServerError "USER EMPTY:User ID required!", Index
        End If
      Else
        SendServerError "USER DENY:You are already logged in!", Index
      End If
    
    ' LOGIN                                                                       '
    Case "LOGIN"
      Client(Index).LoginUser
      
    ' LOGOUT                                                                      '
    Case "LOGOUT"
      Client(Index).LogoutUser
    
    ' LEAVE / DISCONNECT                                                          '
    Case "LEAVE", "DISCONNECT"
      If Client(Index).Loggedin > 0 Then Client(Index).LogoutUser
      WS(Index).Close
      tWSControl_Timer Val(Index)
    
    ' Unknown Command                                                             '
    Case Else
      If Len(Cmd) Then
        ' Unknown Command '
        SendServerError Cmd & ":""" & Cmd & """ Unknown Command!", Index
      Else
        ' No Command '
        SendServerError ":No Command!", Index
      End If
      
  End Select
End Sub
' --------------------------------------------------------------------------------- '
Private Sub DoCountConnections()
  Dim LUsers As Long, CUsers As Long
  Dim i As Long
  
  For i = 0 To MaxSockets
    If Client(i).Connected > 0 And WS(i).State = 7 Then
      CUsers = CUsers + 1
      If Client(i).Loggedin > 0 Then
        LUsers = LUsers + 1
      End If
    End If
  Next i
  
  Server.UsersCount = CUsers
  Server.LUsersCount = LUsers
End Sub
' --------------------------------------------------------------------------------- '
Private Sub LoadControls()
  Dim i As Long
  On Error GoTo LoadControls_Error
  For i = 1 To MaxSockets
    Load WS(i)
    Load tWSControl(i)
  Next i
  Exit Sub
LoadControls_Error:
  Call ErrorRaised("LoadControls()", 0, "Error occurs while loading Controls")
  MsgInformation "Application can not create some Controls!" & vbCrLf & _
                 "Application will be terminated now!", "CRITICAL ERROR!"
  Unload Me
End Sub
' --------------------------------------------------------------------------------- '
Private Sub btnLogger_Click()
  Application.ShowLogger = btnLogger.Value
  If btnLogger.Value Then
    If Not frmLogViewer.Visible Then frmLogViewer.Show vbModeless
  Else
    If frmLogViewer.Visible Then frmLogViewer.Hide
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub btnMore_Click()
  If btnMore.Value Then
    picDetails.Visible = True
    Me.Height = picDetails.Top + picDetails.Height + Screen.TwipsPerPixelY * 25 + picStats.Left
    If Me.Visible Then picDetails.SetFocus
  Else
    Me.Height = picStats.Top + picStats.Height + Screen.TwipsPerPixelY * 25 + picStats.Left
    picDetails.Visible = False
    If Me.Visible Then picStats.SetFocus
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub cmdLockServer_Click()
  Static bUnlock As Boolean
  Static btnMoreValue As Long
  Static bIn As Boolean
  Dim fInp As frmInputBox
  Dim Pwd As String
  Dim bOn As Boolean
  
  If bUnlock And cmdLockServer.Value = vbUnchecked Then
    cmdLockServer.Value = vbChecked
    If Not bIn Then
      bIn = True
      Set fInp = New frmInputBox
      fInp.Caption = " PASSWORD (Pwd: UNLOCK)"
      fInp.lblCaption = "Please input the Unlock Password:"
      fInp.txtInput.PasswordChar = "*"
      fInp.Show vbModal
      Pwd = fInp.txtInput
      Unload fInp: Set fInp = Nothing
      bIn = False
    End If
    If Pwd = "UNLOCK" Then
      bUnlock = False
      cmdLockServer.Value = vbUnchecked
    End If
    Exit Sub
  End If
  If bUnlock Then Exit Sub
  
  bOn = cmdLockServer.Value = vbChecked
  If bOn Then
    bUnlock = True
    cmdLockServer.ToolTipText = " Unlock Server GUI "
    btnMoreValue = btnMore.Value
    btnMore.Value = vbUnchecked
  Else
    cmdLockServer.ToolTipText = " Lock Server GUI "
    btnMore.Value = btnMoreValue
  End If
  
  ' Enable/Disable Controls in Lock Mode  '
  btnMore.Enabled = Not bOn
  btnLogger.Enabled = Not bOn
  cmdOptions.Enabled = Not bOn
    
  On Error Resume Next
  picStats.SetFocus
End Sub
' --------------------------------------------------------------------------------- '
Private Sub cmdOptions_Click()
  Dim sX As Single, sY As Single
  If cmdLockServer.Value = False Then
    sX = cmdOptions.Left + picStats.Left
    sY = cmdOptions.Top + picStats.Top + cmdOptions.Height / 2
    PopupMenu menuOptions, , sX, sY
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub cmdOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then Call cmdOptions_Click
End Sub
' --------------------------------------------------------------------------------- '
Private Sub Form_DblClick()
  Shell "explorer " & App.Path, vbNormalFocus
End Sub
' --------------------------------------------------------------------------------- '
Private Sub lblCaption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMainMenu
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub mnuAccountManager_Click()
  Dim sX As Single
  sX = Me.Left + Me.Width
  If sX > Screen.Width - frmUserDB.Width Then
    sX = Screen.Width - frmUserDB.Width
  End If
  UserDB.RSClients.Requery
  frmUserDB.UpdateForm
  frmUserDB.Move sX, Me.Top
  frmUserDB.Show vbModeless
End Sub
' --------------------------------------------------------------------------------- '
Private Sub mnuShowHide_Click()
  ToggleUIVisibility
End Sub
' --------------------------------------------------------------------------------- '
Public Sub ToggleUIVisibility()
  If Me.Visible And Me.WindowState = vbNormal Then
    UIHide
  Else
    UIShow
  End If
End Sub
' --------------------------------------------------------------------------------- '
Public Sub UIShow()
  On Error Resume Next
  Me.Show
  Me.WindowState = vbNormal
  Me.SetFocus
End Sub
' --------------------------------------------------------------------------------- '
Public Sub UIHide()
  On Error Resume Next
  Me.WindowState = vbMinimized
  Me.Hide
End Sub
' --------------------------------------------------------------------------------- '
Private Sub picDetails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMainMenu
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub picStats_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMainMenu
  End If
End Sub
' --------------------------------------------------------------------------------- '
Public Sub PopupMainMenu()
  If cmdLockServer.Value = False Then
    PopupMenu menuServer
  End If
End Sub
' --------------------------------------------------------------------------------- '
Public Sub EnableTimers()
  Dim i As Long
  tFastControl.Enabled = True
  tAppControl.Enabled = True
  tWSCControl.Enabled = True
  tWSSControl.Enabled = True
  For i = 0 To MaxSockets
    tWSControl(i).Enabled = True
  Next i
End Sub
' --------------------------------------------------------------------------------- '
Public Sub DisableTimers()
  Dim i As Long
  tFastControl.Enabled = False
  tAppControl.Enabled = False
  tWSCControl.Enabled = False
  tWSSControl.Enabled = False
  For i = 0 To MaxSockets
    tWSControl(i).Enabled = False
  Next i
End Sub
' --------------------------------------------------------------------------------- '
Public Sub CloseAllSockets()
  Dim i As Long
  On Error Resume Next
  If WSS.State Then
    WSS.Close
    DoEvents
    tWSSControl_Timer
  End If
  If WSC.State Then
    WSC.Close
    DoEvents
    tWSCControl_Timer
  End If
  For i = 0 To MaxSockets
    If WS(i).State Then
      WS(i).Close
      DoEvents
      tWSControl_Timer Val(i)
    End If
  Next i
End Sub
' --------------------------------------------------------------------------------- '
Public Sub ToggleServerState()
  If Server.State = iStoped Then
    StartServer
  Else
    StopServer
  End If
End Sub
' --------------------------------------------------------------------------------- '
Public Sub StopServer()
  OldAutoReStart = Server.AutoRestart
  Server.AutoRestart = 0
  CloseAllSockets
End Sub
' --------------------------------------------------------------------------------- '
Public Sub StartServer()
  If OldAutoReStart Then
    Server.AutoRestart = OldAutoReStart
    Server.AutoStart = True
    AutoStartDone = False
  End If
  DoEvents
  Call tWSSControl_Timer
End Sub
' --------------------------------------------------------------------------------- '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' Form Events                                                                       '
Private Sub Form_Load()
  Init
  LoadControls
  EnableTimers
  AddStatsNodes
  btnMore_Click
  btnLogger.Value = Abs(Application.ShowLogger)
  Set ServerTrayIcon = New clsTrayIcon
  With ServerTrayIcon
    Set .Icon = Me.Icon
    .hWnd = Me.hWnd
    .Tip = " " & App.ProductName & " "
    '.AddIcon
  End With
End Sub
' --------------------------------------------------------------------------------- '
Private Sub AddStatsNodes()
  Set NNodes = tvwDetails.Nodes
  ' Server Node '
  Set NStats.RNServer = tvwDetails.Nodes.Add(, , "RNServer", "Server")
  Set NStats.NServerState = tvwDetails.Nodes.Add("RNServer", tvwChild, "NServerState")
  Set NStats.NUpTime = tvwDetails.Nodes.Add("RNServer", tvwChild, "NUpTime")
  Set NStats.NUsersCount = tvwDetails.Nodes.Add("RNServer", tvwChild, "NUsersCount")
  Set NStats.NLUsersCount = tvwDetails.Nodes.Add("RNServer", tvwChild, "NLUsersCount")
  Set NStats.NChannels = tvwDetails.Nodes.Add("RNServer", tvwChild, "NChannels")
  Set NStats.NDataIn = tvwDetails.Nodes.Add("RNServer", tvwChild, "NDataIn")
  Set NStats.NDataOut = tvwDetails.Nodes.Add("RNServer", tvwChild, "NDataOut")
  
  ' Clients Node  '
  Set NStats.RNClients = tvwDetails.Nodes.Add(, , "RNClients", "Clients")
  
  ' Channels Node '
  Set NStats.RNChannels = tvwDetails.Nodes.Add(, , "RNChannels", "Channels")
End Sub
' --------------------------------------------------------------------------------- '
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMainMenu
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub Form_Resize()
  bFormMoved = True
End Sub
' --------------------------------------------------------------------------------- '
Private Sub Form_Unload(Cancel As Integer)
  Dim bMin As Boolean
  If Not bForceExit Then
    bMin = Me.WindowState = vbMinimized
    If bMin Then Me.WindowState = vbNormal
    If MsgQuestion("You are about to close Net Messenger Server" & vbCrLf & _
                   "Application, do you realy want to do this?", "Confirm") <> 2 Then
      Cancel = True
      If bMin Then Me.WindowState = vbMinimized
      Exit Sub
    End If
  End If
  StopServer
  frmUserDB.Hide
  CloseAllSockets
  If Not bRecallUnload Then
    ServerTrayIcon.RemoveIcon
    bRecallUnload = True
    Cancel = True
    Me.Hide
    Exit Sub
  End If
  DisableTimers
  CleanUp
  End
End Sub
' --------------------------------------------------------------------------------- '
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Select Case UnloadMode
    Case 0, 1
      If cmdLockServer.Value And bForceExit = False Then
        Cancel = True
        If Not Me.Visible Or Me.WindowState = vbMinimized Then
          ToggleUIVisibility
        End If
        cmdLockServer.Value = vbUnchecked
      End If
    Case Else
      ServerTrayIcon.RemoveIcon
      bForceExit = True
  End Select
End Sub
' Form Events  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' Application Menu                                                                  '
Private Sub mnuQuit_Click()
  Unload Me
End Sub
' --------------------------------------------------------------------------------- '
Private Sub mnuStartStop_Click()
  ToggleServerState
End Sub
' --------------------------------------------------------------------------------- '
Private Sub mnuPause_Click()
  If Server.State <> iStoped Then
    Server.Paused = Not Server.Paused
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub mnuClientDisconnect_Click()
  Dim i As Long
  If SelectedClient > 0 Then
    i = SelectedClient - 1
    WS(i).Close
  End If
  SelectedClient = 0
End Sub
' Application Menu ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' Tray Icon Handler                                                                 '
Private Sub ServerTrayIcon_MouseUp(ByVal Button As Long)
  On Error Resume Next
  Select Case Button
    Case vbLeftButton
      ToggleUIVisibility
    Case vbRightButton
      If cmdLockServer.Value Then
        ToggleUIVisibility
      Else
        PopupMainMenu
      End If
  End Select
End Sub
' Tray Icon Handler  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' Application Control                                                               '
Private Sub tAppControl_Timer()
  Static lRecallUnload As Long
  Static OldAppState As cxeAppState
  Dim Tmp As String
  
  If bRecallUnload Then
    If lRecallUnload > 0 Then
      bForceExit = True
      Unload Me
    End If
    lRecallUnload = lRecallUnload + 1
  End If
  
  ' Update some Details in the TreeView  '
  If btnMore.Value Then
    If Server.Listening > 0 Then
      If Server.Paused Then
        Tmp = "State: Paused"
      Else
        Tmp = "State: Listening"
      End If
    Else
      Tmp = "State: Stopped"
    End If
    ChangeNStat "NServerState", Tmp
    ChangeNStat "NUsersCount", "Connections: " & Server.UsersCount
    ChangeNStat "NLUsersCount", "Users: " & Server.LUsersCount
    ChangeNStat "NChannels", "Channels: " & Channels.Count
    ChangeNStat "NDataOut", "Data Out: " & Server.DataOut & " Bytes"
    ChangeNStat "NDataIn", "Data In: " & Server.DataIn & " Bytes"
    If Server.Listening > 0 Then
      Tmp = "Uptime: " & TimeLeftAfter(Server.Listening)
    Else
      Tmp = "Uptime: down"
    End If
    ChangeNStat "NUpTime", Tmp
  End If
  
  DoWatchClients
  DoCountConnections

  If Server.Listening <> 0 And Server.Paused Then Server.State = iPaused
  If Server.Listening <> 0 And Not Server.Paused Then Server.State = iListening
  If Server.Listening = 0 Then Server.State = iStoped
  
  ' Event: Application State changed  '
  If OldAppState <> Server.State Then
    If Server.State = iStoped Then
      Set Me.Icon = picIcons(2)
      mnuStartStop.Caption = "&Start"
      mnuPause.Enabled = False
    ElseIf Server.State = iListening Then
      Set Me.Icon = picIcons(1)
      mnuStartStop.Caption = "&Stop"
      mnuPause.Enabled = True
      If OldAppState = iPaused Then
        EventRaised ServerResumed
      End If
    ElseIf Server.State = iPaused Then
      Set Me.Icon = picIcons(3)
      mnuStartStop.Caption = "&Stop"
      mnuPause.Enabled = True
      If OldAppState <> iPaused Then
        EventRaised ServerPaused
      End If
    End If
    If Server.Listening Then Tmp = "Listening"
    If Server.Paused Then Tmp = "Paused"
    If Server.Listening = 0 Then Tmp = "Stopped"
    ServerTrayIcon.Tip = " " & App.ProductName & " - " & Tmp & " "
    Set ServerTrayIcon.Icon = Me.Icon
    ServerTrayIcon.UpdateIcon
    mnuPause.Checked = Server.Paused
    OldAppState = Server.State
  End If
  
  bFormMoved = True
End Sub
' --------------------------------------------------------------------------------- '
Private Sub tFastControl_Timer()
  Static fX As Single, fY As Single
  Dim Tmp As String
  
  ' Refresh Stats   '
  If picStats.Visible Then
    ' Setup Uptime  '
    If Server.Listening = 0 Then
      Tmp = "down"
    Else
      Tmp = TimeLeftAfter(Server.Listening)
    End If
    If lblUpTime <> Tmp Then lblUpTime = Tmp
    ' Users Count '
    If lblUsersCount <> Server.UsersCount Then
      lblUsersCount = Server.UsersCount
    End If
    If lblLoggedCount <> Server.LUsersCount Then
      lblLoggedCount = Server.LUsersCount
    End If
    ' Channels Count '
    If lblChannelsCount <> Channels.Count Then
      lblChannelsCount = Channels.Count
    End If
    ' Data In/Out   '
    If lblDataIn <> Server.DataIn Then
      lblDataIn = Server.DataIn
    End If
    If lblDataOut <> Server.DataOut Then
      lblDataOut = Server.DataOut
    End If
  End If
  
  ' Dock Log Window '
  If Me.Left <> fX Or Me.Top <> fY Then bFormMoved = True
  If bFormMoved Then
    If Application.ShowLogger Then
      If frmLogViewer.Visible Then
        If frmLogViewer.mnuDock.Checked Then
          frmLogViewer.Top = Me.Top + Me.Height
          frmLogViewer.Left = Me.Left
        End If
      End If
    End If
    bFormMoved = False
  End If
  
  ' Update Menu Captions  '
  If Me.Visible Then
    Tmp = "&Hide"
  Else
    Tmp = "S&how"
  End If
  If mnuShowHide.Caption <> Tmp Then mnuShowHide.Caption = Tmp
  
  fX = Me.Left
  fY = Me.Top
End Sub
' --------------------------------------------------------------------------------- '
Private Sub tvwDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim NNode As Node
  Set NNode = tvwDetails.HitTest(X, Y)
  If Not NNode Is Nothing Then
    lblState = NNode.Key
  Else
    lblState = ""
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub tvwDetails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim NNode As Node
  If Button = vbRightButton Then
    Set NNode = tvwDetails.HitTest(X, Y)
    If Not NNode Is Nothing Then
      lblState = NNode.Key
      Select Case NNode.Key
        Case "RNServer"
          PopupMainMenu
        Case Else
          If Left(NNode.Key, 7) = "Client\" Then
            If DelimiterCount(NNode.Key, "\") = 1 Then
              If Left(GetToken(NNode.Key, "\", 2), 2) = "WS" Then
                SelectedClient = Val(Mid(GetToken(NNode.Key, "\", 2), 3)) + 1
                PopupClientMenu
              End If
            End If
          End If
      End Select
    Else
      lblState = ""
    End If
  End If
End Sub
' --------------------------------------------------------------------------------- '
Public Sub PopupClientMenu()
  If cmdLockServer.Value = False Then
    PopupMenu menuClient
  End If
End Sub
' --------------------------------------------------------------------------------- '
Private Sub tWSCControl_Timer()
  On Error GoTo tWSCControl_Error
  
  Exit Sub
tWSCControl_Error:
  Call ErrorRaised("tWSCControl_Timer()", Err.Number, Err.Description)
  Resume Next
End Sub
' --------------------------------------------------------------------------------- '
Private Sub tWSControl_Timer(Index As Integer)
  Static bIn As Boolean
  Dim i As Long
  
  If bIn Then Exit Sub
  On Error GoTo tWSControl_Error
  bIn = True
  
  i = Index
  Select Case WS(i).State
    Case sckClosed
      If Client(i).Connected > 0 Then
        ' Event: Client disconnected  '
        If Client(i).Loggedin > 0 Then
          ' Event: User Log out '
          Client(i).LogoutUser
        End If
        RemoveClientNode i
        Client(i).Connected = 0 ' for Watcher '
        Set Client(i) = New clsClient
        Client(i).Index = i + 1
      End If
    Case sckConnected
      If Client(i).Connected = 0 Then
        Set Client(i) = New clsClient
        Client(i).Index = i + 1
        Client(i).Connected = TimeLong
        ' Event: Client connected  '
        AddClientNode i
      End If
    Case sckConnecting        ' nothing '
    Case sckConnectionPending ' nothing '
    Case sckResolvingHost     ' nothing '
    Case Else
      If WS(i).State Then WS(i).Close
  End Select
  
  bIn = False
  Exit Sub
tWSControl_Error:
  Call ErrorRaised("tWSControl_Timer(" & Index & ")", Err.Number, Err.Description)
  Resume Next
End Sub
' --------------------------------------------------------------------------------- '
Sub AddClientWatch(ByVal Index As Long)
  'On Error Resume Next
  WatchClients.Add Client(Index)
  If Err Then Debug.Print "AddClientWatch() Error: " & Err.Description
End Sub
' --------------------------------------------------------------------------------- '
Sub DoWatchClients()
  Dim i As Long
  Dim Cli As clsClient
  Dim Temp As String
  'On Error Resume Next
  i = 1
  While i <= WatchClients.Count
    Set Cli = WatchClients(i)
    If Cli Is Nothing Then
      WatchClients.Remove i
    Else
      If Cli.Connected = 0 Or Cli.NNode Is Nothing Then
        WatchClients.Remove i
      Else
        If btnMore.Value Then
          If Len(Cli.User) Then
            Temp = Cli.User
          Else
            Temp = "Client " & Cli.Index
          End If
          ChangeNStat Cli.NNode.Key, Temp
          Temp = "Connected: " & TimeLeftAfter(Cli.Connected)
          ChangeNStat Cli.NConnected.Key, Temp
          If Cli.Loggedin > 0 Then
            Temp = "Loggedin: " & TimeLeftAfter(Cli.Loggedin)
          Else
            Temp = "Loggedin: No"
          End If
          ChangeNStat Cli.NLoggedin.Key, Temp
          Temp = Trim(WS(Cli.Index - 1).RemoteHostIP)
          If Len(Temp) Then
            Temp = "IP-Address: " & Temp
          Else
            Temp = "IP-Address: unknown"
          End If
          ChangeNStat Cli.NIPAddress.Key, Temp
          Temp = Trim(WS(Cli.Index - 1).RemoteHost)
          If Len(Temp) Then
            Temp = "Hostname: " & Temp
          Else
            Temp = "Hostname: unknown"
          End If
          ChangeNStat Cli.NHostName.Key, Temp
        End If
      End If
    End If
    i = i + 1
  Wend
  If Err Then Debug.Print "DoWatchClient() Error: " & Err.Description
End Sub
' --------------------------------------------------------------------------------- '
Public Sub AddClientNode(ByVal Index As Long)
  Dim Key As String
  On Error Resume Next
  RemoveClientNode Index
  Key = "Client\WS" & Index
  Set Client(Index).NNode = NNodes.Add("RNClients", tvwChild, Key, "Client " & Index)
  Set Client(Index).NConnected = NNodes.Add(Key, tvwChild, Key & "\NConnected", "Connected: Yes")
  Set Client(Index).NLoggedin = NNodes.Add(Key, tvwChild, Key & "\NLoggedin", "Loggedin: No")
  Set Client(Index).NIPAddress = NNodes.Add(Key, tvwChild, Key & "\NIPAddress", "IP-Address: unknown")
  Set Client(Index).NHostName = NNodes.Add(Key, tvwChild, Key & "\NHostName", "Hostname: unknown")
  If Err = 0 Then
    AddClientWatch Index
  End If
  If Err Then Debug.Print "AddClientNode() Error: " & Err.Description
End Sub
' --------------------------------------------------------------------------------- '
Public Sub RemoveClientNode(ByVal Index As Long)
  On Error Resume Next
  If Not Client(Index).NNode Is Nothing Then
    NNodes.Remove Client(Index).NNode.Key
    Set Client(Index).NNode = Nothing
  End If
  If Err Then Debug.Print "RemoveClientNode() Error: " & Err.Description
End Sub
' --------------------------------------------------------------------------------- '
Private Sub tWSSControl_Timer()
  Static AutoRestartTmr As Long
  Dim EventNr As cxeLogEvents
    
  On Error GoTo tWSSControl_Error
  
  Select Case WSS.State
    Case sckListening
      If Server.Listening = 0 Then  ' Event: Server Started '
        Server.Listening = TimeLong
        EventNr = ServerStarted
      End If
    Case sckClosed
      If Server.Listening Then
        Server.Listening = 0
        EventNr = ServerClosed
      End If
      If Not AutoStartDone And (Server.AutoStart Or Server.AutoRestart) Then
        If Server.Port Then
          WSS.LocalPort = Server.Port
          WSS.Listen  ' Start the Server  '
        End If
        AutoStartDone = True
      Else
        If Server.AutoRestart Then
          If AutoRestartTmr = 0 Then AutoRestartTmr = Tick + Server.AutoRestart * 1000
          If AutoRestartTmr < Tick Then
            AutoRestartTmr = Tick + Server.AutoRestart * 1000
            AutoStartDone = False
          End If
        End If
      End If
    Case Else
      WSS.Close ' if Server is not Listening and not Closed then Close it '
  End Select
  
  If EventNr Then EventRaised EventNr
  
  Exit Sub
tWSSControl_Error:
  Call ErrorRaised("tWSSControl_Timer()", Err.Number, Err.Description)
  Resume Next
End Sub
' Application Control  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' WinSock Close                                                                     '
Private Sub WS_Close(Index As Integer)
  Call tWSControl_Timer(Index)
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSS_Close()
  ' disabled '
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSC_Close()
  Call tWSCControl_Timer
End Sub
' WinSock Close  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' WinSock Connect                                                                   '
Private Sub WS_Connect(Index As Integer)
  Call tWSControl_Timer(Index)
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSS_Connect()
  ' disabled '
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSC_Connect()
  Call tWSCControl_Timer
End Sub
' WinSock Connect  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' WinSock Connection Request                                                        '
Private Sub WS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
  ' disabled '
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSC_ConnectionRequest(ByVal requestID As Long)
  ' disabled '
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSS_ConnectionRequest(ByVal requestID As Long)
  Dim i As Long, bAccepted As Boolean
  
  On Error GoTo WSS_ConnectionRequest_Resume
  
  If Server.Paused Then Exit Sub
  
  For i = 0 To MaxSockets
    If Client(i).Connected = 0 And WS(i).State = 0 Then
      Call tWSControl_Timer(Val(i))
      WS(i).Accept requestID
      bAccepted = True
      Exit For
    End If
  Next i
  
WSS_ConnectionRequest_Resume:
  If Not bAccepted Then
    If Err Then
      Call ErrorRaised("WSS_ConnectionRequest()", Err.Number, Err.Description)
    Else
      Call ErrorRaised("WSS_ConnectionRequest()", 0, "Connection request rejected. Server full!")
    End If
  End If
End Sub
' WinSock Connection Request ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' WinSock Data Arrival                                                              '
Private Sub WS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  Dim n As Long, i As Long
  Dim Buffer As String
  Dim Temp As String
  Dim CmdL As String
  
  ' Be sure if user is connected  '
  If Client(Index).Connected = 0 Then
    Call tWSControl_Timer(Index)
  End If
  
  ' Get Buffer and new Data  '
  Buffer = Client(Index).DataBuffer
  WS(Index).GetData Temp
  Server.DataIn = Server.DataIn + Len(Temp)
  
  ' add buffer if exists  '
  If Len(Buffer) Then Temp = Buffer & Temp
  n = DelimiterCount(Temp, vbCrLf, True) + 1
  Buffer = GetToken(Temp, vbCrLf, n, True) ' save last line to buffer  '
  
  ' Get each line delimited by CrLf (except the last line)  '
  For i = 1 To n - 1
    CmdL = GetToken(Temp, vbCrLf, i, True)
    If Len(CmdL) < Server.MaxCmdLLen Then
      Call ExecuteClientCommand(CmdL, Index)
    Else
      Client(Index).LeaveReason = "ERROR: Command Lenght overflow!"
      WS(Index).Close
    End If
  Next i
  
  ' Check Buffer overflow   '
  If Len(Buffer) > Server.MaxBuffer Then
    ' Error: Buffer overflow (BAD CLIENT) '
    Client(Index).LeaveReason = "ERROR: Data Buffer overflow!"
    WS(Index).Close
  End If
  
  '*******************************************'
  ' To Do:                                    '
  ' 1. Flood protection                       '
  ' 1.2 automatic Host/IP ban on permanent    '
  '     flooding or buffer overflowing        '
  ' 2. Buffer overflow and Flood logging      '
  '    with some informations (USER/IP/REASON)'
  '*******************************************'
  Client(Index).DataBuffer = Buffer
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSC_DataArrival(ByVal bytesTotal As Long)
  ' may not recieve something (this Socket is just for Output) '
  Dim Temp As String
  WSC.GetData Temp
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSS_DataArrival(ByVal bytesTotal As Long)
  ' disabled '
End Sub
' WinSock Data Arrival ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '



'/:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::/'
' WinSock Error                                                                     '
Private Sub WS_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call ErrorRaised("WS_Error(" & Index & ")", Val(Number), Description)
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call ErrorRaised("WSC_Error()", Val(Number), Description)
End Sub
' --------------------------------------------------------------------------------- '
Private Sub WSS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call ErrorRaised("WSS_Error()", Val(Number), Description)
End Sub
' WinSock Error  ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^ ^  '


