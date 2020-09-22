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
  NewChan.Created = TimeLong
  
  On Error Resume Next
  
  If Len(NewChan.Caption) Then
    mCol.Add NewChan, NewChan.Caption
  Else
    Err.Raise 8, "AddAs()", "Unnamed Channel!"
  End If
    
  Set NewChan = Nothing
  AddAs = Err
End Function


Public Function Channel(ByVal IndexKey As Variant) As clsChannel
Attribute Channel.VB_UserMemId = 0
  On Error Resume Next
  
  If VarType(IndexKey) = vbString Then IndexKey = Trim(IndexKey)
  Set Channel = mCol(IndexKey)
  
  If Err Then Set Channel = Nothing
End Function


Public Function Exist(ByVal IndexKey As Variant) As Boolean
  Dim CChan As clsChannel
  On Error Resume Next
  
  If VarType(IndexKey) = vbString Then IndexKey = Trim(IndexKey)
  Set CChan = mCol(IndexKey)
  
  If Err = 0 Then Exist = True
  Set CChan = Nothing
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

