VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChanUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsChanUser"
Attribute VB_Ext_KEY = "Member0" ,"clsChanUser"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mCol    As Collection
Public ChanKey  As String
Public NNode    As Node

Public Function Add(User As clsChanUser, Optional ByVal Key As String) As Long
  On Error Resume Next
  
  User.User = Trim(User.User)
  User.Entered = TimeLong
  If Len(User.User) Then
    If Len(Key) = 0 Then
      Key = UCase(User.User)
    End If
    mCol.Add User, Key
    If Err = 0 Then
      User.ChanKey = Me.ChanKey
      If NNode Is Nothing Then
        Set NNode = NNodes.Add("NChan\" & ChanKey, tvwChild, "NChan\" & ChanKey & "\Users", "Users")
      End If
      Set User.NNode = NNodes.Add("NChan\" & ChanKey & "\Users", tvwChild, "NChan\" & ChanKey & "\User\" & User.User, User.User)
      If Err Then Err.Clear
    End If
  Else
    Err.Raise 8, "Add()", "Unnamed User!"
  End If
  
  Add = Err
End Function


Public Function AddAs(ByVal User As String, ByVal WSIndex As Long, Optional ByVal Flags As Long) As Long
  Dim NewUser As clsChanUser
  
  Set NewUser = New clsChanUser
  NewUser.User = Trim(User)
  NewUser.WSIndex = WSIndex
  NewUser.Flags = Flags
  
  AddAs = Me.Add(NewUser)
  
  Set NewUser = Nothing
End Function


Public Function User(ByVal IndexKey As Variant) As clsChanUser
Attribute User.VB_UserMemId = 0
  On Error Resume Next
  
  If VarType(IndexKey) = vbString Then IndexKey = UCase(Trim(IndexKey))
  Set User = mCol(IndexKey)
  
  If Err Then Set User = Nothing
End Function


Public Function Exist(ByVal IndexKey As Variant) As Boolean
  Dim CUsr As clsChanUser
  On Error Resume Next
  
  If VarType(IndexKey) = vbString Then IndexKey = UCase(Trim(IndexKey))
  Set CUsr = mCol(IndexKey)
  
  If Err = 0 Then Exist = True
  Set CUsr = Nothing
End Function


Public Property Get Count() As Long
  Count = mCol.Count
End Property


Public Function Remove(ByVal IndexKey As Variant) As Long
  On Error Resume Next

  If VarType(IndexKey) = vbString Then IndexKey = UCase(Trim(IndexKey))
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

