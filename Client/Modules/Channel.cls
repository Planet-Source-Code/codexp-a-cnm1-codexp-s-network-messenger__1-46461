VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Caption      As String           ' Name of the Channel         '
Attribute Caption.VB_VarUserMemId = 0
Public Topic        As String           ' Topic in this Channel       '
Public Creator      As String           ' User ID who has created     '
Public Created      As Double           ' Timestamp when it created   '
Public Password     As String           ' Password for this Channel   '
Public Registered   As Double           ' When was registered ?       '
Public Flags        As Long             ' Flags for this Chan         '
Public Users        As clsChanUsers     ' Users in this Channel       '
Public Window       As frmConsole

Private Sub Class_Initialize()
  Set Users = New clsChanUsers
  Set Window = New frmConsole
  Set Window.Icon = Window.picIcon(1).Picture
  Set Window.Chan = Me
  Window.bChannel = True
End Sub

Private Sub Class_Terminate()
  Set Users = Nothing
  Window.Hide
  Unload Window
  Set Window = Nothing
End Sub

Public Function MakeUserList() As String
  Dim CUser As clsChanUser
  Dim sDel As String
  Dim sFlg As String
  Dim sUL As String
  
  For Each CUser In Users
    If Len(sUL) Then sDel = ","
    If CUser.Flags Then
      sFlg = "|" & CUser.Flags
    Else
      sFlg = ""
    End If
    sUL = sUL & sDel & CUser.User & sFlg
  Next CUser
  
  MakeUserList = sUL
End Function

Public Sub OpenWindow()
  On Error Resume Next
  If Window.WindowState = vbMinimized Then Window.WindowState = vbNormal
  If Not Window.Visible Then Window.Show
  If Len(Window.txtLine) = 0 Then Window.SetFocus
End Sub

