VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public User         As String   ' User (Account ID)   '
Public UserName     As String   ' User's Name         '
Public NickName     As String   ' User's Nickname     '
Public Password     As String   ' User's Password     '
Public Address1     As String   ' Address Part 1      '
Public Address2     As String   ' Address Part 2      '
Public Address3     As String   ' Address Part 3      '
Public EMail        As String   ' E-Mail Address      '
Public Telephone    As String   ' Telephone           '
Public ICQN         As String   ' ICQ Number          '
Public MSNID        As String   ' MSN Messenger ID    '
Public BDate        As Double   ' Date of Birth       '
Public Gender       As Integer  ' Gender              '
Public Connected    As Double   ' Timestamp for Connection start (0 if disconnected) '
Public Loggedin     As Double   ' Timestamp for Login time (0 if not logged in) '
Public LastLogin    As Double   ' Timestamp for the last Login  '
Public Registered   As Double   ' Timestamp for Registration Date '
Public DataBuffer   As String   ' Data Buffer         '
Public LeaveReason  As String   ' Reason for leaving  '
Public Flags        As Long     ' User Flags (Rights) '
Public Index        As Long     ' Socket Index + 1    '
Public NNode        As Node
Public NLoggedin    As Node
Public NConnected   As Node
Public NIPAddress   As Node
Public NHostName    As Node
Public NewUser      As clsClient

Private BCPI        As Long     ' Bad Commands per Interval (Count) '
Private BCPIT       As Long     ' Bad Commands per Interval (Timer) '
Private BCPII       As Long     ' BCPI - Time Interval in Seconds   '
Private BCPIM       As Long     ' BCPI - Count of Allowed Failures  '



Public Function Clone() As clsClient
  Dim NewClient As New clsClient
  NewClient.Address1 = Me.Address1
  NewClient.Address2 = Me.Address2
  NewClient.Address3 = Me.Address3
  NewClient.BDate = Me.BDate
  NewClient.Connected = Me.Connected
  NewClient.DataBuffer = Me.DataBuffer
  NewClient.EMail = Me.EMail
  NewClient.Flags = Me.Flags
  NewClient.Gender = Me.Gender
  NewClient.ICQN = Me.ICQN
  NewClient.Index = Me.Index
  NewClient.LastLogin = Me.LastLogin
  NewClient.LeaveReason = Me.LeaveReason
  NewClient.Loggedin = Me.Loggedin
  NewClient.MSNID = Me.MSNID
  NewClient.NickName = Me.NickName
  NewClient.Password = Me.Password
  NewClient.Registered = Me.Registered
  NewClient.Telephone = Me.Telephone
  NewClient.User = Me.User
  NewClient.UserName = Me.UserName
  Set Clone = NewClient
End Function


Public Function BadCommand()
  If BCPII < 1 Then BCPII = 10
  If BCPIM < 1 Then BCPIM = 20
  BCPI = BCPI + 1
  If BCPIT < TickSeconds Then
    BCPIT = TickSeconds + BCPII
    If BCPI > BCPIM Then
      ' Maximum of Failures reached '
      If Index Then
        LeaveReason = "ERROR: Server Flooding!"
        frmMain.WS(Index - 1).Close
      End If
    End If
  End If
End Function


Public Sub SetBCPIInterval(ByVal Interval As Long)
  BCPII = Interval
End Sub


Public Sub SetBCPIMaxFailures(ByVal MaxFailures As Long)
  BCPIM = MaxFailures
End Sub


Public Function LoginUser() As Boolean
  If Index = 0 Then Exit Function
  
  If Loggedin = 0 Then
    If Len(Trim(User)) Then
      If UserIDReserved(User) Then
        SendServerError "LOGIN RESERVED:User ID is reserved!", Index - 1
      Else
        If UserIDIsUsed(User) Then
          SendServerError "LOGIN INUSE:User ID is already in use!", Index - 1
        Else
          If UserIDVorbidden(User) Then
            SendServerError "LOGIN VORBIDDEN:User ID is vorbidden!", Index - 1
          Else
            If UserDB.UserExists(Client(Index - 1).User) Then
              If UserDB.GetUserData(Client(Index - 1)) Then
                ' Event: User Logged in  '
                Loggedin = TimeLong
                EventRaised UserDefinedMessage, "User """ & User & """ is logged in!"
                SendToAllUsers ":Server LOGIN " & User
                LoginUser = True
              Else
                SendServerError "LOGIN DENY:UserID or Password is wrong!", Index - 1
              End If
            Else
              SendServerError "LOGIN DENY:UserID is not registered!", Index - 1
            End If
          End If
        End If
      End If
    Else
      SendServerError "LOGIN NOID:User ID is missing!", Index - 1
    End If
  Else
    SendServerError "LOGIN DENY:You are already logged in!", Index - 1
  End If
End Function


Public Sub LogoutUser()
  If Index = 0 Then Exit Sub
  If Loggedin = 0 Then
    SendServerError "LOGOUT DENY:You are not logged in!", Index - 1
  Else
    If Len(User) Then
      ChanUserLogout User
      EventRaised UserDefinedMessage, "User """ & User & """ is logged out!"
      SendToAllUsers ":Server LOGOUT " & User & _
              IIf(Len(LeaveReason), ":" & LeaveReason, "")
    End If
  End If
  Registered = 0
  Loggedin = 0
  Flags = 0
End Sub


