VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsChannel"
Attribute VB_Ext_KEY = "Member0" ,"clsChannel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"clsChanUsers"
Option Explicit

Private mCol As Collection

Public Function Add(Channel As clsChannel, Optional ByVal Key As String) As Long
  On Error Resume Next
    
  Channel.Caption = Trim(Channel.Caption)
  Channel.Created = TimeLong
  If Len(Channel.Caption) Then
    If Len(Key) = 0 Then
      Key = UCase(Channel.Caption)
    End If
    mCol.Add Channel, Key
    If Err = 0 Then
      Channel.Users.ChanKey = Channel.Caption
      Set Channel.NNode = NNodes.Add("RNChannels", tvwChild, "NChan\" & Channel.Caption, "#" & Channel.Caption)
      If Err Then Err.Clear
    End If
  Else
    Err.Raise 8, "Add()", "Unnamed Channel!"
  End If
  
  Add = Err
End Function


Public Function AddAs(ByVal Caption As String, Optional ByVal Creator As String, Optional ByVal Password As String) As Long
  Dim NewChan As clsChannel
  
  Set NewChan = New clsChannel
  NewChan.Caption = Trim(Caption)
  NewChan.Creator = Trim(Creator)
  NewChan.Password = Trim(Password)
  
  AddAs = Me.Add(NewChan)
      
  Set NewChan = Nothing
End Function


Public Function Channel(ByVal IndexKey As Variant) As clsChannel
Attribute Channel.VB_UserMemId = 0
  On Error Resume Next
  
  If VarType(IndexKey) = vbString Then IndexKey = Trim(IndexKey)
  Set Channel = mCol(IndexKey)
  
  If Err Then Set Channel = Nothing
End Function


Public Function Exist(ByVal IndexKey As Variant) As Boolean
  Dim cChan As clsChannel
  On Error Resume Next
  
  If VarType(IndexKey) = vbString Then IndexKey = Trim(IndexKey)
  Set cChan = mCol(IndexKey)
  
  If Err = 0 Then Exist = True
  Set cChan = Nothing
End Function


Public Property Get Count() As Long
  Count = mCol.Count
End Property


Public Function Remove(ByVal IndexKey As Variant) As Long
  On Error Resume Next

  If VarType(IndexKey) = vbString Then IndexKey = Trim(IndexKey)
  mCol.Remove IndexKey
  
  Remove = Err
End Function


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
  Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub


Public Function UserMayCreate(ByVal UserID As String) As Boolean
  Dim cChan As clsChannel
  Dim lCount As Long
  UserID = Trim(UserID)
  If UserID = "" Then Exit Function
  For Each cChan In mCol
    If cChan.Users.Exist(UserID) Then
      lCount = lCount + 1
    End If
  Next cChan
  If lCount < 10 And mCol.Count < MaxChannels Then UserMayCreate = True
End Function


Public Function PartUser(Client As clsClient, ByVal tChannel As String) As Boolean
  Dim tChan As clsChannel
  tChannel = Trim(tChannel)
  If Client.Loggedin > 0 Then
    If Left(tChannel, 1) = "#" Then tChannel = Trim(Mid(tChannel, 2))
    If Len(tChannel) Then
      If Exist(tChannel) Then
        Set tChan = Channel(tChannel)
        If tChan.Users.Exist(Client.User) Then
          tChan.SendToUsers ":" & Client.User & " PART #" & tChannel
          tChan.Users.Remove Client.User
          If tChan.Users.Count < 1 Then Remove tChannel
        Else
          SendServerError "PART NOUSER:You are not on this Channel!", Client.Index - 1
        End If
      Else
        SendServerError "PART NOCHANNEL:Channel does not exist!", Client.Index - 1
      End If
    Else
      SendServerError "PART EMPTY:No Channel!", Client.Index - 1
    End If
  Else
    SendServerError "PART DENY:You are not logged in!", Client.Index - 1
  End If
End Function


Public Function JoinUser(Client As clsClient, ByVal tChannel As String) As Boolean
  Dim tChan   As clsChannel
  Dim uFlags  As Long
  
  tChannel = Trim(tChannel)
  If Client.Loggedin > 0 Then
    If Left(tChannel, 1) = "#" Then tChannel = Trim(Mid(tChannel, 2))
    If Len(tChannel) Then
      If ChanIsValid(tChannel) Then
        If UserMayCreate(Client.User) Then
          If Exist(tChannel) = False Then
            If AddAs(tChannel, Client.User) = 0 Then
              ' Event: New Channel Created  '
            End If
          End If
          If Exist(tChannel) Then
            Set tChan = Channel(tChannel)
            If Not tChan.Users.Exist(Client.User) Then
              uFlags = 0
              If tChan.Users.Count = 0 Then uFlags = 1
              ' Event: User Joined a Channel  '
              Call tChan.Users.AddAs(Client.User, Client.Index - 1, uFlags)
              tChan.SendToUsers ":" & Client.User & " JOIN #" & tChan.Caption
              SendToClient ":Server USERS #" & tChan.Caption & ":" & tChan.MakeUserList, Client.Index - 1
            Else
              SendServerError "JOIN JOINED #" & tChan.Caption & ":You are already joined #" & tChan.Caption, Client.Index - 1
            End If
          Else
            SendServerError "JOIN ERROR:Can not create Channel!", Client.Index - 1
          End If
        Else
          SendServerError "JOIN LIMIT:You may not create more Channels!", Client.Index - 1
        End If
      Else
        SendServerError "JOIN INVALID:Invalid Channel!", Client.Index - 1
      End If
    Else
      SendServerError "JOIN EMPTY:No Channel!", Client.Index - 1
    End If
  Else
    SendServerError "JOIN DENY:You are not logged in!", Client.Index - 1
  End If
End Function

