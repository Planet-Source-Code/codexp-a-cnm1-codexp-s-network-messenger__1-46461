Attribute VB_Name = "basUserDB"
Option Explicit

Public UserDB As clsUserDB

Public Function FileExists(ByVal FileName As String) As Boolean
  Dim lAttr As Long
  FileName = Trim(FileName)
  If FileName = "" Then Exit Function
  lAttr = vbArchive Or vbHidden Or vbReadOnly Or vbSystem
  If UCase(Dir(FileName, lAttr)) = UCase(GetLastToken(FileName, "\")) Then
    FileExists = True
  End If
End Function

Public Sub UserDB_Init(Optional ByVal DBFile As String)
  Set UserDB = New clsUserDB
  UserDB.DBPassword = "opera"
  UserDB.OpenUserDB DBFile
  Load frmUserDB
End Sub

Public Sub UserDB_CleanUp()
  Unload frmUserDB
End Sub

Public Sub RegExecute(ByVal CommandLine As String, ByVal Index As Long)
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
        ' QUERY USER                                                              '
        Case "USER"
          If UserDB.UserExists(Params(2)) Then
            SendToClient UserDB.MakeQueryUserInformations(Params(2), Params(3)), Index
          Else
            SendServerError "QUERY USER NOUSER " & Params(2) & ":User does not exist in DataBase!", Index
          End If
      End Select
      
    ' REG                                                                         '
    Case "REG"
      Select Case UCase(Params(1))
        Case "UPDATE"
          If RegisterClient(Index, True) Then
            SendToClient ":Server REG UPDATE DONE:Your Informations are successfuly updated!", Index
          End If
        Case "FILL", "SET"
          Msg = Trim(Msg)
          Dim NewUser As clsClient
          If Client(Index).NewUser Is Nothing Then
            Set Client(Index).NewUser = Client(Index).Clone
          End If
          Set NewUser = Client(Index).NewUser
          Select Case UCase(Params(2))
            ' REG FILL USER                                   '
            Case "USER", "USERID", "ID"
              NewUser.User = Msg
            ' REG FILL PASSWORD                               '
            Case "PASSWORD", "PASS", "PWD", "PASSWD"
              NewUser.Password = Msg
            ' REG FILL USERNAME                               '
            Case "USERNAME", "FULLNAME", "NAME"
              NewUser.UserName = Msg
            ' REG FILL NICKNAME                               '
            Case "NICKNAME", "NICK"
              NewUser.NickName = Msg
            ' REG FILL ADDRESS1                               '
            Case "ADDRESS1", "ADDR1"
              NewUser.Address1 = Msg
            ' REG FILL ADDRESS2                               '
            Case "ADDRESS2", "ADDR2"
              NewUser.Address2 = Msg
            ' REG FILL ADDRESS3                               '
            Case "ADDRESS3", "ADDR3"
              NewUser.Address3 = Msg
            ' REG FILL BDATE                                  '
            Case "BDATE", "BDAY", "BIRTHDATE", "BIRTHDAY"
              If BDateIsValid(Msg) Then
                NewUser.BDate = TimeToLong(Msg)
              Else
                NewUser.BDate = 0
              End If
            ' REG FILL EMAIL                                  '
            Case "EMAIL", "E-MAIL", "MAIL"
              NewUser.EMail = Msg
            ' REG FILL PHONE                                  '
            Case "PHONE", "TELEPHONE"
              NewUser.Telephone = Msg
            ' REG FILL ICQN                                   '
            Case "ICQN", "ICQ"
              NewUser.ICQN = Msg
            ' REG FILL MSNID                                  '
            Case "MSNID", "MSN"
              NewUser.MSNID = Msg
            ' REG FILL GENDER                                 '
            Case "GENDER", "SEX"
              NewUser.Gender = GenderValue(Msg)
            Case Else ' Syntax Error  '
              SendServerError "REG NOFIELD " & Params(2) & _
                           ":This field is not valid!", Index
          End Select
        Case Else
          If Client(Index).Loggedin > 0 Then
            ' DENY - Loggedin users are not allowed to register '
            SendServerError "REG DENY " & Params(2) & _
                            ":You are loggedin and may not register!", Index
          Else
            Select Case UCase(Params(1))
              Case "REGISTER"
                If RegisterClient(Index) Then
                  SendToClient ":Server REG DONE:You are successfuly registered!", Index
                End If
              Case Else ' Syntax Error  '
                SendServerError "REG NOOP " & Params(2) & _
                             ":Unknown Register Operation!", Index
            End Select
          End If
      End Select
  End Select
End Sub


Private Function GenderValue(ByVal Gender As String) As Integer
  Dim vRet As Integer
  Gender = Trim(Gender)
  Select Case UCase(Gender)
    Case "M", "MALE", "MAN", "1": vRet = 1
    Case "F", "FEMALE", "WOMAN", "2": vRet = 2
  End Select
  GenderValue = vRet
End Function


Public Function CheckUserID(ByVal UserID As String) As Long
  ' Return Values:                '
  ' 0  User ID is ok!             '
  ' 1  User ID is empty           '
  ' 2  User ID is too long        '
  ' 3  User ID is invalid         '
  Dim uidErr As Long
  Dim Tmp As String
  Dim i As Long
  Tmp = "`!@#$&*+\|;:'"",.<>/? "
  
  UserID = Trim(UserID)
  If UserID = "" Then uidErr = 1
  If Len(UserID) > 20 Then uidErr = 2
  If uidErr = 0 Then
    ' Check if user id is valid '
    For i = 1 To Len(UserID)
      If InStr(Tmp, Mid(UserID, i, 1)) Then
        uidErr = 3
        Exit For
      End If
    Next i
  End If
  
  CheckUserID = uidErr
End Function


Public Function CheckPassword(ByVal Pass As String) As Long
  ' Return Values:                '
  ' 0  Password is ok!            '
  ' 1  Password is empty          '
  ' 2  Password is too long       '
  ' 3  Password is invalid        '
  Dim uidErr As Long
  Dim Tmp As String
  Dim i As Long
  Tmp = "`!@#$&*+\|;:'"",<>/?"
  
  Pass = Trim(Pass)
  If Pass = "" Then uidErr = 1
  If Len(Pass) > 30 Then uidErr = 2
  If uidErr = 0 Then
    ' Check if user id is valid '
    For i = 1 To Len(Pass)
      If InStr(Tmp, Mid(Pass, i, 1)) Then
        uidErr = 3
        Exit For
      End If
    Next i
  End If
  
  CheckPassword = uidErr
End Function


' Open Recordset with SQL - Returns True if any Recordset was found else False  '
Public Function OpenSQL(Query As QueryDef, RSReturn As Recordset) As Boolean
  On Error Resume Next
  Set RSReturn = Query.OpenRecordset(dbOpenSnapshot)
  If Not RSReturn Is Nothing Then
    If RSReturn.RecordCount > 0 Then
      RSReturn.MoveFirst
      OpenSQL = True
    End If
  End If
End Function


Public Function RegisterClient(ByVal Index As Long, Optional bUpdate As Boolean) As Boolean
  ' Return Value: True if success else False  '
  Dim ErrMsg As String
  Dim NewUser As clsClient
  
  Set NewUser = Client(Index).NewUser
  If NewUser Is Nothing Then
    ErrMsg = "REG USERDATA EMPTY:User Data is empty"
    SendServerError ErrMsg, Index
    Exit Function
  End If
  
  If Client(Index).Loggedin = 0 Then bUpdate = False
  
  ' Check UserID Validness      '
  Select Case CheckUserID(NewUser.User)
    Case 1: ErrMsg = "REG USERID EMPTY:User ID is empty"
    Case 2: ErrMsg = "REG USERID LARGE:User ID is too long"
    Case 3: ErrMsg = "REG USERID INVALID:User ID is invalid"
    Case Else
      If UserDB.UserExists(NewUser.User, IIf(bUpdate, Client(Index).User, "")) Then
        ErrMsg = "REG USERID REGISTERED:User ID is already registered"
      End If
  End Select
  If Len(ErrMsg) Then
    SendServerError ErrMsg, Index
    Exit Function
  End If
  
  ' Check Password Validness    '
  Select Case CheckPassword(NewUser.Password)
    Case 1: ErrMsg = "REG PASSWORD EMPTY:Password is empty"
    Case 2: ErrMsg = "REG PASSWORD LARGE:Password is too long"
    Case 3: ErrMsg = "REG PASSWORD INVALID:Password is invalid"
  End Select
  If Len(ErrMsg) Then
    SendServerError ErrMsg, Index
    Exit Function
  End If
  
  ' Truncate overlenght       '
  NewUser.Address1 = Left(NewUser.Address1, 50)
  NewUser.Address2 = Left(NewUser.Address2, 50)
  NewUser.Address3 = Left(NewUser.Address3, 50)
  NewUser.EMail = Left(NewUser.EMail, 50)
  NewUser.MSNID = Left(NewUser.MSNID, 50)
  NewUser.ICQN = Left(NewUser.ICQN, 10)
  NewUser.Telephone = Left(NewUser.Telephone, 30)
  NewUser.UserName = Left(NewUser.UserName, 50)
  NewUser.NickName = Left(NewUser.NickName, 30)
  NewUser.Gender = NewUser.Gender Mod 3
  NewUser.Registered = TimeLong
  NewUser.LastLogin = TimeLong
  
  ' Check if Nickname already registered  '
  If Len(NewUser.NickName) Then
    If UserDB.NickNameExists(NewUser.NickName, Client(Index).User) Then
      ErrMsg = "REG NICKNAME REGISTERED:Nickname is already used by other User!"
      SendServerError ErrMsg, Index
      Exit Function
    End If
  End If
  
  ' if Update then Delete User first    '
  If bUpdate Then
    If Client(Index).Registered > 0 Then
      ' Try to Update   '
      Call UserDB.UpdateUser(Client(Index), NewUser, ErrMsg)
      If Len(ErrMsg) Then ErrMsg = "REG UPDATE ERROR:" & ErrMsg
    Else
      ErrMsg = "REG UPDATE DANY:You may not modify anything!"
      SendServerError ErrMsg, Index
      Exit Function
    End If
  Else
    ' Try to register   '
    Call UserDB.RegisterUser(NewUser, ErrMsg)
    If Len(ErrMsg) Then ErrMsg = "REG REGISTER ERROR:" & ErrMsg
  End If
  
  
  If Len(ErrMsg) Then
    SendServerError ErrMsg, Index
  Else
    RegisterClient = True
  End If
End Function
