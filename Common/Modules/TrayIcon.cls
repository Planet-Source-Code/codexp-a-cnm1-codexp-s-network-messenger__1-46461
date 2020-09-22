VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Icon As StdPicture
Public hWnd As Long
Public Tip  As String
Public ID   As Long

Private NID As NOTIFYICONDATA

Public Event Click()
Public Event DblClick(ByVal Button As Long)
Public Event MouseMove()
Public Event MouseDown(ByVal Button As Long)
Public Event MouseUp(ByVal Button As Long)



' Load Icon from File                                         '
Public Function LoadIcon(ByVal FileName As String) As Boolean
  On Error Resume Next
  Set Icon = LoadPicture(FileName)
  If Icon.Type <> vbPicTypeIcon Then
    Set Icon = Nothing
  End If
  If Err Then
    LoadIcon = True
    Err.Clear
  End If
End Function

Public Sub AddIcon()
  'prepare the NOTIFYICONDATA type with the
  'required parameters:
  
  '.cbSize: Size of this structure, in bytes.
  '
  '.hwnd:   Handle of the window that will receive
  '         notification messages associated with
  '         an icon in the taskbar status area.
  '
  'uID:     Application-defined identifier of
  '         the taskbar icon. In an application
  '         with a single tray icon, this can be
  '         an arbitrary number.  For apps with
  '         multiple icons, each icon ID must be
  '         different as this member identifies
  '         which of the icons was selected.
  '
  '.uFlags: flags that indicate which of the other
  '         members contain valid data. This member
  '         can be a combination of the following:
  '         NIF_ICON    hIcon member is valid.
  '         NIF_MESSAGE uCallbackMessage member is valid.
  '         NIF_TIP     szTip member is valid.
  '
  'uCallbackMessage: Application-defined message identifier.
  '         The system uses this identifier for
  '         notification messages that it sends
  '         to the window identified in hWnd.
  '         These notifications are sent when a
  '         mouse event occurs in the bounding
  '         rectangle of the icon. (Note: 'callback'
  '         is a bit misused here (in the context of
  '         other callback demonstrations); there is
  '         no systray-specific callback defined -
  '         instead the form itself must be subclassed
  '         to respond to this message.
  '
  'hIcon:   Handle to the icon to add, modify, or delete.
  '
  'szTip:   Tooltip text to display for the icon. Must
  '         be terminated with a Chr$(0).
  
  On Error Resume Next
  
  If NID.uID = 0 Then
    With NID
      .cbSize = LenB(NID)
      .hWnd = Me.hWnd
      .uID = Me.ID
      .uFlags = NIF_ICON Or NIF_MESSAGE Or IIf(Len(Me.Tip), NIF_TIP, 0)
      .uCallbackMessage = WM_MYHOOK
      .hIcon = Me.Icon
      .szTip = Me.Tip & Chr$(0)
    End With
    
    ShellTrayAdd NID
    colSubClassed.Add Me, "C" & Me.ID
    RegisterNewTrayIcon
  End If
End Sub

Public Sub RemoveIcon()
  On Error Resume Next
  If NID.uID Then
    With NID
      ShellTrayRemove NID
      colSubClassed.Remove "C" & Me.ID
      UnRegisterTrayIcon
      .uID = 0
    End With
  End If
End Sub


Public Sub UpdateIcon()
  If NID.uID Then
    With NID
      .hIcon = Me.Icon
      .uFlags = NIF_ICON Or NIF_MESSAGE Or IIf(Len(Me.Tip), NIF_TIP, 0)
      .szTip = Me.Tip & Chr$(0)
    End With
    ShellTrayModify NID
  End If
End Sub


Public Sub IconProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Static MD As Boolean
  If wParam <> Me.ID Then Exit Sub
  
  Select Case lParam
    Case WM_LBUTTONDOWN
      MD = True
      RaiseEvent MouseDown(1)
    Case WM_LBUTTONUP
      RaiseEvent MouseUp(1)
      If MD Then RaiseEvent Click
      MD = False
    Case WM_LBUTTONDBLCLK
      RaiseEvent DblClick(1)
    Case WM_MBUTTONDOWN
      RaiseEvent MouseDown(4)
    Case WM_MBUTTONUP
      RaiseEvent MouseUp(4)
    Case WM_MBUTTONDBLCLK
      RaiseEvent DblClick(4)
    Case WM_RBUTTONDOWN
      RaiseEvent MouseDown(2)
    Case WM_RBUTTONUP
      RaiseEvent MouseUp(2)
    Case WM_RBUTTONDBLCLK
      RaiseEvent DblClick(2)
    Case 512
      RaiseEvent MouseMove
    Case Else
      Debug.Print lParam
  End Select

End Sub

Private Sub Class_Initialize()
  lTINumCreated = lTINumCreated + 1
  ID = lTINumCreated
End Sub

Private Sub Class_Terminate()
  RemoveIcon
End Sub
