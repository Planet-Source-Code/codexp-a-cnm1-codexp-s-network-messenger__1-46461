VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChanUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public User     As String     ' User ID                     '
Attribute User.VB_VarUserMemId = 0
Public Flags    As Long       ' Flags for this User         '
Public Entered  As Double     ' Login Timestamp             '
Public WSIndex  As Long       ' WinSock Index
Public ChanKey  As String
Public NNode    As Node

Private Sub Class_Terminate()
  On Error Resume Next
  NNodes.Remove NNode.Key
End Sub
