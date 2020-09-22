Attribute VB_Name = "basMain"
Option Explicit

Public Const INIFile = "CNM.INI"
Public Const MaxServers = 255
Public Const MaxUsers = 255
Public Const MaxPrivates = 20
Public Const MaxChannels = 20

Public Enum cxeAppState
  Disconnected = 0
  Connected = 1
  Waiting = 2
End Enum

Public Enum cxeMsgType
  mMessage = 0
  mEvent = 1
  mError = 2
End Enum

Public Type cxtServerType
  Description As String   ' Server Name (short Description) '
  Group       As String   ' Server Group        '
  Host        As String   ' Server HOST/IP      '
  Port        As Long     ' Port Number         '
End Type

Public Type cxtClientType
  User        As String   ' User (Account ID)   '
  UserName    As String   ' User's Name (Nick)  '
  Password    As String   ' User's Password     '
  Connected   As Double   ' Timestamp for Connection start (0 if disconnected) '
  LoggedIn    As Double   ' Timestamp for Login time (0 if not logged in) '
  DataBuffer  As String   ' Data Buffer         '
  LeaveReason As String   ' Reason for leaving  '
  AutoConnect As Boolean  ' Autoconnect to Server '
  AutoLogin   As Boolean  ' Auto Login on Connect '
  NNode       As Node
End Type

Public Type cxtApplicationType
  Started     As Double   ' Timestamp (when started)  '
  Caption     As String   ' Application's Caption '
  Version     As String   ' Application's Version '
  Copyright   As String   ' Copyright (C)       '
  ShowLogger  As Boolean  ' Show Log Window ?   '
  OpenConMgr  As Boolean  ' Auto open Connections window on Startup?  '
End Type

Public bDontAutoLogin     As Boolean
Public bAutoLogin         As Boolean
Public AutoConnectEnabled As Boolean
Public NumServers         As Long
Public Application        As cxtApplicationType
Public Server(MaxServers) As cxtServerType
Public LastServer         As cxtServerType
Public EmptyServer        As cxtServerType
Public CurrentServer      As cxtServerType
Public Client             As cxtClientType
Public UserList(MaxUsers) As cxtClientType
Public EmptyUser          As cxtClientType
Public NumUsers           As Long
Public Channels           As New clsChannels
Public fConsole           As New frmConsole
Public fPrivate(MaxPrivates)  As New frmConsole
'Public fChannel(MaxChannels)  As New frmConsole



Public Sub Init()
  fConsole.lstCon.BackColor = vbBlack
  fConsole.lstCon.ForeColor = vbWhite
  fConsole.bConsole = True
  fConsole.AddMessage "CNM Client " & Version & " Console"
  ' Setup standard Values                             '
  Application.Caption = "CodeXP's Net Messenger"
  Application.Copyright = "Copyright (C)2003 by CodeXP"
  Application.Version = "V" & Version
  Application.Started = TimeLong
  'Application.ShowLogger = True
  AutoConnectEnabled = True
  Client.AutoLogin = True
  
  LoadAppSettings
  EventRaised "CNM Client Started"
End Sub


Public Sub CleanUp()
  SaveAppSettings
  EventRaised "CNM Client Terminated"
End Sub


Public Function Version() As String
  Version = App.Major & "." & App.Minor & App.Revision
End Function


Public Sub EventRaised(ByVal Message As String, Optional ByVal Handler As String)
  
  Handler = Trim(UCase(Handler))
  Select Case Handler
    Case "EXAMPLE"
      ' DO EXAMPLE  '
  End Select
  
  If Len(Message) Then
    ' Add to Logging File: (TimeStamp & EventMsg) '
  End If
  
  ' Show the Logger Window if required  '
  If Application.ShowLogger Then
    If Handler = "ERROR" Then
      frmLogViewer.AddLogMessage Message, mError
    Else
      frmLogViewer.AddLogMessage Message, mEvent
    End If
  End If
End Sub


Public Sub LogMessage(ByVal Message As String, Optional ByVal Handler As String)
  
  Handler = Trim(UCase(Handler))
  Select Case Handler
    Case "EXAMPLE"
      ' DO EXAMPLE  '
  End Select
  
  If Len(Message) Then
    ' Add to Logging File: (TimeStamp & EventMsg) '
  End If
  
  ' Show the Logger Window if required  '
  If Application.ShowLogger Then
    If Handler = "ERROR" Then
      frmLogViewer.AddLogMessage Message, mError
    Else
      frmLogViewer.AddLogMessage Message, mMessage
    End If
  End If
End Sub


Public Sub ErrorRaised(Module As String, ErrNr As Long, ErrDesc As String)
  If Application.ShowLogger Then
    frmLogViewer.AddLogMessage "Error Nr " & ErrNr & " raised in Module " & Module & vbCrLf & _
                               "Description: " & ErrDesc, mError
  End If
  ' To Do: Logging Errors to File '
End Sub


Public Function MsgQuestion(Message As String, Optional Title As String = "-1")
  Static fDia As Dialog
  
  If Not fDia Is Nothing Then
    Do While fDia.Visible: DoEvents: Loop
  End If
  Set fDia = New Dialog
  With fDia
    .ResetForm
    .picInformation.Visible = False
    .picQuestion.Visible = True
    .cmdButton1.Caption = "&No"
    .cmdButton2.Caption = "&Yes"
    .cmdButton2.Visible = True
    .cmdButton2.Default = True
    .Message = Message
    If Title <> "-1" Then .Caption = Title
    .Show
    Do While .Visible: DoEvents: Loop
    MsgQuestion = .ReturnValue
  End With
  Unload fDia
  Set fDia = Nothing
End Function


Public Sub MsgInformation(Message As String, Optional Title As String = "-1", Optional Critical As Boolean)
  Static lCount As Long
  Dim fDia As Dialog
  
  If lCount > 9 Then Exit Sub
  lCount = lCount + 1
  Set fDia = New Dialog
  With fDia
    .ResetForm
    If Critical Then
      .picInformation.Visible = False
      .picCritical.Visible = True
    End If
    .Message = Message
    If Title <> "-1" Then .Caption = Title
    .Show
    Do While .Visible: DoEvents: Loop
  End With
  Unload fDia
  Set fDia = Nothing
  lCount = lCount - 1
End Sub


Public Sub LoadAppSettings()
  Dim i As Long
  For i = 0 To MaxServers
    Server(i).Description = GetINI(AppPath & INIFile, "Servers", "Server(" & i & ").Description")
    Server(i).Group = GetINI(AppPath & INIFile, "Servers", "Server(" & i & ").Group")
    Server(i).Host = GetINI(AppPath & INIFile, "Servers", "Server(" & i & ").Host")
    Server(i).Port = Val(GetINI(AppPath & INIFile, "Servers", "Server(" & i & ").Port"))
    If Trim(Server(i).Host) = "" Then
      NumServers = i - 1
      Exit For
    End If
    If Trim(Server(i).Description) = "" Then Server(i).Description = Server(i).Host
    If Trim(Server(i).Group) = "" Then Server(i).Group = "General"
    If Server(i).Port < 1 Then Server(i).Port = 8888
  Next i
  LastServer.Description = GetINI(AppPath & INIFile, "Connection", "Server.Description")
  LastServer.Group = GetINI(AppPath & INIFile, "Connection", "Server.Group")
  LastServer.Host = GetINI(AppPath & INIFile, "Connection", "Server.Host")
  LastServer.Port = Val(GetINI(AppPath & INIFile, "Connection", "Server.Port"))
  Client.User = GetINI(AppPath & INIFile, "User", "User")
  Client.Password = GetINI(AppPath & INIFile, "User", "Password")
  Client.AutoConnect = Val(GetINI(AppPath & INIFile, "User", "AutoConnect", Abs(Client.AutoConnect)))
  Client.AutoLogin = Val(GetINI(AppPath & INIFile, "User", "AutoLogin", Abs(Client.AutoLogin)))
  Application.ShowLogger = Val(GetINI(AppPath & INIFile, "Settings", "LogWindow", "0"))
  Application.OpenConMgr = Val(GetINI(AppPath & INIFile, "Settings", "OpenConMgr", "1"))
End Sub


Public Sub SaveAppSettings()
  Dim i As Long
  ' Remove all old Servers from file  '
  SaveINI AppPath & INIFile, "Servers", Chr(0), Chr(0)
  ' Save Server List  '
  For i = 0 To NumServers
    SaveINI AppPath & INIFile, "Servers", "Server(" & i & ").Description", Server(i).Description
    SaveINI AppPath & INIFile, "Servers", "Server(" & i & ").Group", Server(i).Group
    SaveINI AppPath & INIFile, "Servers", "Server(" & i & ").Host", Server(i).Host
    SaveINI AppPath & INIFile, "Servers", "Server(" & i & ").Port", Server(i).Port
  Next i
  ' Save Last Connection and User '
  SaveINI AppPath & INIFile, "Connection", "Server.Description", LastServer.Description
  SaveINI AppPath & INIFile, "Connection", "Server.Group", LastServer.Group
  SaveINI AppPath & INIFile, "Connection", "Server.Host", LastServer.Host
  SaveINI AppPath & INIFile, "Connection", "Server.Port", LastServer.Port
  SaveINI AppPath & INIFile, "User", "User", Client.User
  SaveINI AppPath & INIFile, "User", "Password", Client.Password
  SaveINI AppPath & INIFile, "User", "AutoConnect", Abs(Client.AutoConnect)
  SaveINI AppPath & INIFile, "User", "AutoLogin", Abs(Client.AutoLogin)
  ' Save Settings '
  SaveINI AppPath & INIFile, "Settings", "LogWindow", Abs(Application.ShowLogger)
  SaveINI AppPath & INIFile, "Settings", "OpenConMgr", Abs(Application.OpenConMgr)
End Sub


Public Function AppPath() As String
  AppPath = AddBackslash(App.Path)
End Function


Public Sub PlaySound(ByVal SndID As String)
  Beep
  Select Case UCase(Trim(SndID))
    Case "ERROR"
    Case "NOTREADY"
    Case "DENY"
  End Select
End Sub


Public Function OpenPrivate(ByVal User As String) As Long
  Dim bFound As Boolean
  Dim i As Long
  If Trim(User) = "" Then Exit Function
  For i = 0 To MaxPrivates
    If Len(Trim(fPrivate(i).User)) Then
      If UCase(Trim(fPrivate(i).User)) = UCase(Trim(User)) Then
        If fPrivate(i).Visible Then
          If Len(fPrivate(i).txtLine) = 0 Then
            fPrivate(i).SetFocus
          End If
        Else
          fPrivate(i).Show
        End If
        bFound = True
        OpenPrivate = i + 1
        Exit For
      End If
    End If
  Next i
  If Not bFound Then
    For i = 0 To MaxPrivates
      If Len(Trim(fPrivate(i).User)) = 0 Then
        fPrivate(i).User = Trim(User)
        If fPrivate(i).Visible Then
          If Len(fPrivate(i).txtLine) = 0 Then
            fPrivate(i).SetFocus
          End If
        Else
          fPrivate(i).Show
        End If
        bFound = True
        OpenPrivate = i + 1
        Exit For
      End If
    Next i
  End If
End Function


Public Sub ShowRegistrationForm()
  Dim sX As Single
  Dim sY As Single
  On Error Resume Next
  If Not frmConnect Is Nothing Then Unload frmConnect
  If frmRegForm Is Nothing Then Load frmRegForm
  sY = frmMain.Top
  sX = frmMain.Left - frmRegForm.Width
  If sX < 0 Then
    sX = frmMain.Left + frmMain.Width
  End If
  frmRegForm.Move sX, sY
  If frmRegForm.Visible Then
    frmRegForm.SetFocus
  Else
    frmRegForm.Show
  End If
End Sub


Public Function FindUser(ByVal UserID As String)

End Function
