VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DBase      As Database
Public DBFile     As String
Public DBPassword As String
Public RSClients  As Recordset



Public Sub OpenUserDB(Optional ByVal DBFile As String)
  Dim ConStr As String
  On Error Resume Next
  If Len(Trim(DBFile)) Then Me.DBFile = Trim(DBFile)
  If Not FileExists(Me.DBFile) Then CreateUserDB
  If Not IsOpen And FileExists(Me.DBFile) Then
    If Len(DBPassword) Then ConStr = ";PWD=" & DBPassword
    Set DBase = OpenDatabase(Me.DBFile, False, False, ConStr)
  End If
  If IsOpen Then
    Set RSClients = DBase.OpenRecordset("Accounts", dbOpenDynaset)
  End If
  If Err Then Debug.Print Err.Description
End Sub


Private Sub CreateUserDB(Optional ByVal ReplaceOldDB As Boolean)
  On Error Resume Next
  If IsOpen Then CloseUserDB
  If ReplaceOldDB Then DeleteUserDB
  If DBase Is Nothing And Len(DBFile) Then
    Set DBase = CreateDatabase(DBFile, dbLangGeneral)
    BuildUserDB
  End If
End Sub


Private Sub BuildUserDB()
  Dim tmpIndex As Index
  Dim tmpTable As TableDef
  
  If Not IsOpen Then Exit Sub
  On Error GoTo BuildUserDB_Error
  
  ' Erstelle eine neue Datenbank  '
  With DBase
    If Len(DBPassword) Then
      .NewPassword "", DBPassword
    End If
    ' Erstelle Tabelle "Accounts" '
    Set tmpTable = .CreateTableDef("Accounts")
    With tmpTable
      ' Felder in der Tabelle erstellen  '
      .Fields.Append .CreateField("ID", dbLong)
      .Fields.Append .CreateField("User", dbText, 20)
      .Fields.Append .CreateField("Password", dbText, 30)
      .Fields.Append .CreateField("Nickname", dbText, 30)
      .Fields.Append .CreateField("Username", dbText, 50)
      .Fields.Append .CreateField("Address1", dbText, 50)
      .Fields.Append .CreateField("Address2", dbText, 50)
      .Fields.Append .CreateField("Address3", dbText, 50)
      .Fields.Append .CreateField("EMail", dbText, 50)
      .Fields.Append .CreateField("ICQN", dbText, 10)
      .Fields.Append .CreateField("MSNID", dbText, 50)
      .Fields.Append .CreateField("Phone", dbText, 30)
      .Fields.Append .CreateField("Registered", dbDouble)
      .Fields.Append .CreateField("LastLogin", dbDouble)
      .Fields.Append .CreateField("ClanID", dbLong)
      .Fields.Append .CreateField("Flags", dbLong)
      .Fields.Append .CreateField("Gender", dbInteger)
      .Fields.Append .CreateField("BDate", dbDouble)
      ' Eigenschaften der Felder festlegen  '
      .Fields("ID").Attributes = dbAutoIncrField
      .Fields("User").Required = True
      .Fields("Password").Required = True
      .Fields("Nickname").AllowZeroLength = True
      .Fields("Username").AllowZeroLength = True
      .Fields("Address1").AllowZeroLength = True
      .Fields("Address2").AllowZeroLength = True
      .Fields("Address3").AllowZeroLength = True
      .Fields("ICQN").AllowZeroLength = True
      .Fields("MSNID").AllowZeroLength = True
      .Fields("EMail").AllowZeroLength = True
      .Fields("Phone").AllowZeroLength = True
      .Fields("Registered").DefaultValue = 0
      .Fields("LastLogin").DefaultValue = 0
      .Fields("Flags").DefaultValue = 0
      .Fields("Gender").DefaultValue = 0
      .Fields("BDate").DefaultValue = 0
      ' Primären Schlüssel erstellen (geht nur beim erstellen neuer Tabelle) '
      .Indexes.Append .CreateIndex("PrimaryKey")                            ''
      .Indexes("PrimaryKey").Unique = True                                  ''
      .Indexes("PrimaryKey").Primary = True                                 ''
      .Indexes("PrimaryKey").Fields.Append .CreateField("ID")  '<            '
      ' Felder Indizieren   '
      .Indexes.Append .CreateIndex("User")
      .Indexes("User").Unique = True
      .Indexes("User").Fields.Append .CreateField("User")
    End With
    ' Tabelle zu der Datenbank anfügen  '
    .TableDefs.Append tmpTable
    .TableDefs.Refresh  ' DB aktualisieren '
    
    ' Erstelle Tabelle "Clans"    '
    Set tmpTable = .CreateTableDef("Clans")
    With tmpTable
      ' Felder in der Tabelle erstellen  '
      .Fields.Append .CreateField("ID", dbLong)
      .Fields.Append .CreateField("Clan", dbText, 50)
      .Fields.Append .CreateField("Description", dbText, 255)
      .Fields.Append .CreateField("Leader", dbText, 30)
      .Fields.Append .CreateField("EMail", dbText, 50)
      ' Eigenschaften der Felder festlegen  '
      .Fields("ID").Attributes = dbAutoIncrField
      .Fields("Clan").Required = True
      .Fields("Description").AllowZeroLength = True
      .Fields("Leader").AllowZeroLength = True
      .Fields("EMail").AllowZeroLength = True
      ' Primären Schlüssel erstellen (geht nur beim erstellen neuer Tabelle) '
      .Indexes.Append .CreateIndex("PrimaryKey")                            ''
      .Indexes("PrimaryKey").Unique = True                                  ''
      .Indexes("PrimaryKey").Primary = True                                 ''
      .Indexes("PrimaryKey").Fields.Append .CreateField("ID")  '<            '
      ' Felder Indizieren   '
      .Indexes.Append .CreateIndex("Clan")
      .Indexes("Clan").Unique = True
      .Indexes("Clan").Fields.Append .CreateField("Clan")
    End With
    ' Tabelle zu der Datenbank anfügen  '
    .TableDefs.Append tmpTable
    .TableDefs.Refresh  ' DB aktualisieren '
  End With
  
  Exit Sub
BuildUserDB_Error:
  Debug.Print "BuildUserDB() Error: " & Err.Description
End Sub


Public Function IsOpen() As Boolean
  IsOpen = Not DBase Is Nothing
End Function


Public Sub RebuildUserDB()
  If DeleteUserDB Then Exit Sub
  OpenUserDB
End Sub


Public Function DeleteUserDB()
  On Local Error Resume Next
  If IsOpen Then CloseUserDB
  If FileExists(DBFile) Then Kill DBFile
  If Err Then Debug.Print Err.Description
  DeleteUserDB = Err
End Function


Public Sub CloseUserDB()
  If Not RSClients Is Nothing Then RSClients.Close
  Set RSClients = Nothing
  If Not DBase Is Nothing Then DBase.Close
  Set DBase = Nothing
End Sub


Private Sub Class_Terminate()
  CloseUserDB
End Sub


Public Function UserExists(ByVal UserID As String, Optional ByVal ExceptUserID As String) As Boolean
  Dim qdfSQL As QueryDef
  Dim rstRS As Recordset
  Dim Tmp     As String
  
  UserID = Trim(UserID)
  ExceptUserID = Trim(ExceptUserID)
  If UserID = "" Or DBase Is Nothing Then Exit Function
  If UCase(UserID) = UCase(ExceptUserID) Then Exit Function
  
  Set qdfSQL = DBase.CreateQueryDef("")
  Tmp = "SELECT User FROM Accounts WHERE User LIKE '" & UserID & "'"
  If Len(ExceptUserID) Then
    Tmp = Tmp & " AND User NOT LIKE '" & ExceptUserID & "'"
  End If
  qdfSQL.SQL = Tmp
  UserExists = OpenSQL(qdfSQL, rstRS)
End Function


Public Function NickNameExists(ByVal NickName As String, Optional ByVal ExceptUserID As String) As Boolean
  Dim qdfSQL  As QueryDef
  Dim rstRS   As Recordset
  Dim Tmp     As String
  
  NickName = Trim(NickName)
  ExceptUserID = Trim(ExceptUserID)
  If NickName = "" Then Exit Function
  
  Set qdfSQL = DBase.CreateQueryDef("")
  Tmp = "SELECT Nickname FROM Accounts WHERE Nickname LIKE '" & NickName & "'"
  If Len(ExceptUserID) Then
    Tmp = Tmp & " AND User NOT LIKE '" & ExceptUserID & "'"
  End If
  qdfSQL.SQL = Tmp
  NickNameExists = OpenSQL(qdfSQL, rstRS)
End Function


Public Function GetUserData(Client As clsClient) As Boolean
  Dim qdfSQL  As QueryDef
  Dim rstRS   As Recordset
  Dim tSQL    As String
  
  On Error GoTo GetUserData_Error
  
  If Client Is Nothing Or DBase Is Nothing Then Exit Function
  If Trim(Client.User) = "" Or Trim(Client.Password) = "" Then Exit Function
  
  Set qdfSQL = DBase.CreateQueryDef("")
  qdfSQL.MaxRecords = 1
  tSQL = "SELECT * FROM Accounts WHERE User LIKE '" & Trim(Client.User) & "'"
  tSQL = tSQL & " AND Password LIKE '" & Trim(Client.Password) & "'"
  qdfSQL.SQL = tSQL
  If OpenSQL(qdfSQL, rstRS) Then
    Client.Address1 = rstRS.Fields("Address1").Value
    Client.Address2 = rstRS.Fields("Address2").Value
    Client.Address3 = rstRS.Fields("Address3").Value
    Client.BDate = rstRS.Fields("BDate").Value
    Client.EMail = rstRS.Fields("EMail").Value
    Client.Flags = rstRS.Fields("Flags").Value
    Client.Gender = rstRS.Fields("Gender").Value
    Client.ICQN = rstRS.Fields("ICQN").Value
    Client.LastLogin = rstRS.Fields("LastLogin").Value
    Client.MSNID = rstRS.Fields("MSNID").Value
    Client.NickName = rstRS.Fields("Nickname").Value
    Client.Registered = rstRS.Fields("Registered").Value
    Client.Telephone = rstRS.Fields("Phone").Value
    Client.UserName = rstRS.Fields("Username").Value
    'Client.clanid = = rstRS.Fields("ClanID").Value
    GetUserData = True
  End If
  
GetUserData_Error:
  If Err Then
    Debug.Print "GetUserData() Error:" & Err.Description
    GetUserData = False
  End If
End Function


Public Function RemoveUser(ByVal UserID As String) As Long
  Dim tSQL As String
  On Error GoTo RemoveUser_Error
  
  UserID = Trim(UserID)
  If UserExists(UserID) Then
    tSQL = "DELETE * FROM Accounts WHERE User LIKE '" & UserID & "'"
    DBase.Execute tSQL
    If UserExists(UserID) Then Error 1
  End If
  
RemoveUser_Error:
  If Err Then RemoveUser = Err
End Function


Public Function UpdateUser(User As clsClient, NewUser As clsClient, Optional ByRef ErrMsg As String) As Long
  ' Return Values:              '
  ' 0 - Update success          '
  '!0 - Update Fails            '
  '     Value is Error Number   '
  ' 1 - Database is not open    '
  ' 2 - User already registered '
  Dim rstNewUser  As Recordset
  Dim qdfSQL      As QueryDef
  Dim tSQL        As String
  
  On Error GoTo UpdateUser_Error
  
  If UserExists(NewUser.User, User.User) Then
    ErrMsg = "UserID is already in use!"
    UpdateUser = 2
    Exit Function
  End If
  
  If DBase Is Nothing Then
    ErrMsg = "DataBase is not available!"
    UpdateUser = 1
    Exit Function
  End If
  
  
  tSQL = "UPDATE Accounts SET " & _
         "  User = '" & NewUser.User & "', " & _
         "  Password = '" & NewUser.Password & "', " & _
         "  Nickname = '" & NewUser.NickName & "', " & _
         "  Username = '" & NewUser.UserName & "', " & _
         "  Address1 = '" & NewUser.Address1 & "', " & _
         "  Address2 = '" & NewUser.Address2 & "', " & _
         "  Address3 = '" & NewUser.Address3 & "', " & _
         "  EMail    = '" & NewUser.EMail & "', " & _
         "  ICQN     = '" & NewUser.ICQN & "', " & _
         "  MSNID    = '" & NewUser.MSNID & "', " & _
         "  Phone    = '" & NewUser.Telephone & "', " & _
         "  Flags    = " & NewUser.Flags & ", " & _
         "  Gender   = " & NewUser.Gender & ", " & _
         "  BDate    = " & NewUser.BDate & " " & _
         "WHERE User LIKE '" & User.User & "' " & _
         "  AND Password LIKE '" & User.Password & "'"
  Set qdfSQL = DBase.CreateQueryDef("")
  qdfSQL.SQL = tSQL
  qdfSQL.Execute
  
  If qdfSQL.RecordsAffected = 0 Then
    ErrMsg = "UserID or Password is wrong!"
    UpdateUser = 3
    Exit Function
  End If
  
UpdateUser_Error:
  If Err Then
    UpdateUser = Err.Number
    ErrMsg = "SERVER ERROR(" & Err.Number & "): " & Err.Description
  End If
End Function


Public Function RegisterUser(Client As clsClient, Optional ByRef ErrMsg As String) As Long
  ' Return Values:            '
  ' 0 - Registration success  '
  '!0 - Registration Fails    '
  '     Value is Error Number '
  ' 1 - Database is not open  '
  Dim rstNewUser As Recordset
  On Error GoTo RegisterUser_Error
  
  If DBase Is Nothing Then Error 1
  Set rstNewUser = DBase.OpenRecordset("Accounts")
  With rstNewUser
    .AddNew
    .Fields("User") = Client.User
    .Fields("Password") = Client.Password
    .Fields("Nickname") = Client.NickName
    .Fields("Username") = Client.UserName
    .Fields("Address1") = Client.Address1
    .Fields("Address2") = Client.Address2
    .Fields("Address3") = Client.Address3
    .Fields("EMail") = Client.EMail
    .Fields("ICQN") = Client.ICQN
    .Fields("MSNID") = Client.MSNID
    .Fields("Phone") = Client.Telephone
    .Fields("Registered") = Client.Registered
    .Fields("LastLogin") = Client.LastLogin
    '.Fields ("ClanID") = Client.clanid
    .Fields("Gender") = Client.Gender
    .Fields("BDate") = Client.BDate
    .Update
  End With
  
RegisterUser_Error:
  If Err Then
    RegisterUser = Err.Number
    ErrMsg = "SERVER ERROR(" & Err.Number & "): " & Err.Description
  End If
End Function


Public Function MakeQueryUserInformations(ByVal UserID As String, Optional QID As String) As String
  Dim qdfSQL  As QueryDef
  Dim rstRS   As Recordset
  Dim Pre     As String
  Dim Tmp     As String
  Dim BDate   As String
  Dim bUE     As Boolean
  
  On Error GoTo MKQUI_ERROR
  
  UserID = Trim(UserID)
  If UserID = "" Then
    UserID = "''"
    Tmp = "QUERY USER " & UserID & " ERROR:UserID is required for this Query!"
  Else
    Set qdfSQL = DBase.CreateQueryDef("")
    qdfSQL.SQL = "SELECT * FROM Accounts WHERE User LIKE '" & UserID & "'"
    If OpenSQL(qdfSQL, rstRS) Then
      If rstRS.RecordCount Then
        rstRS.MoveFirst
        
        If rstRS.AbsolutePosition >= 0 Then
          Pre = "QUERY USER " & UserID & " "
          If Val(rstRS.Fields("BDate")) Then
            BDate = LongToTime(rstRS.Fields("BDate"))
          End If
          Tmp = Pre & "NICKNAME " & QID & ":" & rstRS.Fields("Nickname") & vbCrLf & _
                Pre & "USERNAME " & QID & ":" & rstRS.Fields("Username") & vbCrLf & _
                Pre & "USERID " & QID & ":" & rstRS.Fields("User") & vbCrLf & _
                Pre & "ADDRESS1 " & QID & ":" & rstRS.Fields("Address1") & vbCrLf & _
                Pre & "ADDRESS2 " & QID & ":" & rstRS.Fields("Address2") & vbCrLf & _
                Pre & "ADDRESS3 " & QID & ":" & rstRS.Fields("Address3") & vbCrLf & _
                Pre & "EMAIL " & QID & ":" & rstRS.Fields("EMail") & vbCrLf & _
                Pre & "MSNID " & QID & ":" & rstRS.Fields("MSNID") & vbCrLf & _
                Pre & "PHONE " & QID & ":" & rstRS.Fields("Phone") & vbCrLf & _
                Pre & "GENDER " & QID & ":" & rstRS.Fields("Gender") & vbCrLf & _
                Pre & "BDATE " & QID & ":" & BDate & vbCrLf & _
                Pre & "ICQN " & QID & ":" & rstRS.Fields("ICQN")
        Else
          bUE = True
        End If
      Else
        bUE = True
      End If
    Else
      bUE = True
    End If
    If bUE Then
      Tmp = "QUERY USER " & UserID & " ERROR:User does not exist in DataBase! [UE]"
    End If
  End If
  
MKQUI_ERROR:
  If Err Then
    Tmp = "QUERY USER " & UserID & " ERROR:" & Err.Description
  End If
  MakeQueryUserInformations = Tmp
End Function

