Attribute VB_Name = "basMain"
Option Explicit

Public Const MaxSockets = 255

Public Enum cxeMsgType
  mMessage = 0
  mEvent = 1
  mError = 2
End Enum

Public Enum cxeAppState
  iStoped = 0
  iListening = 1
  iPaused = 2
End Enum

Public Enum cxeLogEvents
  ApplicationStart = 100
  ApplicationExit = 101
  ServerStarted = 102
  ServerClosed = 103
  ServerPaused = 104
  ServerResumed = 105
  UserDefinedMessage = 106
End Enum

Public Type ServerNodes
  NUpTime       As Node
  NUsersCount   As Node
  NLUsersCount  As Node
  NDataIn       As Node
  NDataOut      As Node
  NServerState  As Node
  NChannels     As Node
  RNClients     As Node
  RNChannels    As Node
  RNServer      As Node
End Type

Public Type cxtServerType
  Port        As Long     ' Port Number         '
  Listening   As Double   ' Listening TimeStamp (0 if iddle) '
  Paused      As Boolean  ' Don't accept requests if true '
  MaxUsers    As Long     ' Maximal count of users  '
  MaxBuffer   As Long     ' Maximal Data Buffer Lenght  '
  MaxCmdLLen  As Long     ' Maximal Command Line Lenght '
  UsersCount  As Long     ' Connected users count '
  LUsersCount As Long     ' Logged in users count '
  AutoStart   As Boolean  ' Start Server on Application Start '
  AutoRestart As Long     ' Automaticaly Restart if Idle '
  State       As cxeAppState ' Server Status    '
  DataIn      As Long     ' Received Data Bytes '
  DataOut     As Long     ' Sent Data Bytes     '
End Type

Public Type cxtApplicationType
  Started     As Double   ' Timestamp (when started)  '
  Caption     As String   ' Application's Caption '
  Version     As String   ' Application's Version '
  Copyright   As String   ' Copyright (C)       '
  ShowLogger  As Boolean  ' Show Log Window ?   '
End Type

Public Server             As cxtServerType
Public Application        As cxtApplicationType
Public Client(MaxSockets) As New clsClient
Public bForceExit         As Boolean
Public NStats             As ServerNodes
Public NNodes             As Nodes
Public WatchClients       As New Collection



Public Sub Init()
  ' Change Dir to <App.Path>  '
  ChDrive App.Path
  ChDir App.Path
  ' Create required Dirs '
  If Dir(AddBackslash(App.Path) & "Data", vbDirectory) = "" Then
    MkDir AddBackslash(App.Path) & "Data"
  End If
  ' Setup standard Values                             '
  Application.Caption = "CodeXP's Net Messenger Server"
  Application.Copyright = "Copyright (C)2003 by CodeXP"
  Application.Version = "V" & Version
  Application.Started = TimeLong
  Application.ShowLogger = False
  Server.MaxBuffer = &HFFF
  Server.MaxCmdLLen = &H1FF
  Server.MaxUsers = MaxSockets
  Server.Port = 8888
  Server.AutoStart = True
  Server.AutoRestart = 5
  EventRaised ApplicationStart  'Event: AppStart'
  ' Open Database Connection  '
  UserDB_Init AddBackslash(App.Path) & "Data\users.db"
End Sub


Public Sub CleanUp()
  UserDB_CleanUp
  EventRaised ApplicationExit   'Event: AppExit'
End Sub


Public Function Version() As String
  Version = App.Major & "." & App.Minor & App.Revision
End Function


Public Sub EventRaised(EventID As cxeLogEvents, Optional uMsg As String)
  Dim EventMsg As String, ResNum As Long
  
  ' Get Event Message from Resource '
  ResNum = Val(LoadResString(0))
  If EventID >= 100 And EventID <= 100 + ResNum Then
    EventMsg = LoadResString(EventID)
  End If
  ' Parse Event Message (replace Variables) '
  EventMsg = Replace(EventMsg, "%AppName%", Application.Caption)
  EventMsg = Replace(EventMsg, "%Port%", Server.Port)
  EventMsg = Replace(EventMsg, "%UMsg%", uMsg)
  
  If Len(EventMsg) Then
    ' Add to Logging File: (TimeStamp & EventMsg) '
  End If
  
  ' Show the Logger Window if required  '
  If Application.ShowLogger Then
    frmLogViewer.AddLogMessage EventMsg, mEvent
  End If
End Sub


Public Sub ErrorRaised(Module As String, ErrNr As Long, ErrDesc As String)
  If Application.ShowLogger Then
    frmLogViewer.AddLogMessage "Error Nr " & ErrNr & " raised in Module " & Module & vbCrLf & _
                               "Description: " & ErrDesc, mError
  End If
End Sub


Public Function MsgQuestion(Message As String, Optional Title As String = "-1") As Long
  Dialog.ResetForm
  Dialog.picInformation.Visible = False
  Dialog.picQuestion.Visible = True
  Dialog.cmdButton1.Caption = "&No"
  Dialog.cmdButton2.Caption = "&Yes"
  Dialog.cmdButton2.Visible = True
  Dialog.cmdButton2.Default = True
  Dialog.Message = Message
  If Title <> "-1" Then Dialog.Caption = Title
  Dialog.Show vbModal
  MsgQuestion = Dialog.ReturnValue
End Function


Public Function MsgYesNoCancel(Message As String, Optional Title As String = "-1") As Long
  Static Dialog As Dialog
  
  If Dialog Is Nothing Then Set Dialog = New Dialog
  Do While Dialog.Visible: DoEvents: Loop
  Dialog.ResetForm
  Dialog.picInformation.Visible = False
  Dialog.picQuestion.Visible = True
  Dialog.cmdButton1.Caption = "&Cancel"
  Dialog.cmdButton2.Caption = "&No"
  Dialog.cmdButton3.Caption = "&Yes"
  Dialog.cmdButton2.Visible = True
  Dialog.cmdButton3.Visible = True
  Dialog.cmdButton3.Default = True
  Dialog.Message = Message
  If Title <> "-1" Then Dialog.Caption = Title
  Dialog.Show
  Do While Dialog.Visible: DoEvents: Loop
  MsgYesNoCancel = Dialog.ReturnValue
End Function


Public Sub MsgInformation(Message As String, Optional Title As String = "-1", Optional Critical As Boolean)
  Dialog.ResetForm
  If Critical Then
    Dialog.picInformation.Visible = False
    Dialog.picCritical.Visible = True
  End If
  Dialog.Message = Message
  If Title <> "-1" Then Dialog.Caption = Title
  Dialog.Show vbModal
End Sub


Public Function UserIDReserved(ByVal User As String) As Boolean
  Dim bV As Boolean
  bV = True
  Select Case UCase(Trim(User))
    Case "SERV"
    Case "SERVER"
    Case "SERVICE"
    Case "ADMIN"
    Case "ADMINISTRATOR"
    Case "CHANBOT"
    Case "SERVBOT"
    Case "BOT"
    Case "CODEXP"
    Case "WESSELOH"
    Case "EUGEN"  ' <- Indiziert!!!  '
    Case Else
      bV = False
  End Select
  UserIDReserved = bV
End Function


Public Function UserIDVorbidden(ByVal User As String) As Boolean
  Dim bV As Boolean
  bV = True
  Select Case UCase(Trim(User))
    Case "BINLADEN"
    Case "SEX"
    Case "PENIS"
    Case "WAGINA"
    Case "FOTZE"
    Case "SCHORN"
    Case Else
      If InStr(UCase(User), "FUCK") Then
      ElseIf InStr(UCase(User), "WAREZ") Then
      ElseIf InStr(UCase(User), "SUCK") Then
      ElseIf InStr(UCase(User), "PORN") Then
      Else
        bV = False
      End If
  End Select
  UserIDVorbidden = bV
End Function


Public Function UserIDIsUsed(ByVal User As String) As Boolean
  Dim i As Long
  For i = 0 To MaxSockets
    If Client(i).Loggedin Then
      If UCase(Trim(Client(i).User)) = UCase(Trim(User)) Then
        UserIDIsUsed = True
        Exit For
      End If
    End If
  Next i
End Function


Public Function UserGetIndex(ByVal User As String) As Long
  Dim i As Long
  User = Trim(User)
  If User = "" Then Exit Function
  For i = 0 To MaxSockets
    If Client(i).Loggedin > 0 Then
      If UCase(Trim(Client(i).User)) = UCase(User) Then
        UserGetIndex = i + 1
        Exit For
      End If
    End If
  Next i
End Function


Public Sub ChangeNStat(ByVal Key As String, ByVal Text As String)
  On Error Resume Next
  If frmMain.tvwDetails.Nodes(Key).Text <> Text Then
    frmMain.tvwDetails.Nodes(Key).Text = Text
  End If
End Sub

        
Public Sub SendServerError(ByVal ErrLine As String, ByVal Index As Long, Optional ByVal ErrCmd As String = "000", Optional ByVal CountFailure As Boolean = True)
  If Trim(ErrCmd) = "" Then ErrCmd = "000"
  Call SendToClient(":Server " & ErrCmd & " " & ErrLine, Index)
  If CountFailure Then Client(Index).BadCommand
End Sub


Public Sub SendToAllUsers(CommandLine As String, Optional Except As Long = -1)
  Dim i As Long
  For i = 0 To MaxSockets
    If Client(i).Loggedin > 0 Then
      If Except <> i Then
        SendToClient CommandLine, i
      End If
    End If
  Next i
End Sub


Public Sub SendToClient(CommandLine As String, Index As Long)
  Dim i As Long
  Dim Temp As String
  With frmMain.WS(Index)
    If .State = 7 Then
      Temp = CommandLine & vbCrLf
      .SendData Temp
      Server.DataOut = Server.DataOut + Len(Temp)
      For i = 0 To 10: DoEvents: Next i
    End If
  End With
End Sub

