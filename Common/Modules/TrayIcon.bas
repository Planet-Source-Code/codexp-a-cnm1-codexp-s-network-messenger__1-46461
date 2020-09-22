Attribute VB_Name = "basTrayIcon"
Option Explicit

' Some Types for SysTray API  '
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

' Constants of SysTray API  '
Public Const NIM_ADD As Long = &H0
Public Const NIM_MODIFY As Long = &H1
Public Const NIM_DELETE As Long = &H2

Public Const NIF_ICON As Long = &H2
Public Const NIF_TIP As Long = &H4
Public Const NIF_MESSAGE As Long = &H1

' Constants of Windows Messaging for the Callback '
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203

Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209

Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206

' Get/SetWindowLong Messages  '
Public Const GWL_WNDPROC As Long = (-4)
Public Const GWL_HWNDPARENT As Long = (-8)
Public Const GWL_ID As Long = (-12)
Public Const GWL_STYLE As Long = (-16)
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_USERDATA As Long = (-21)

' General Windows Messages    '
Public Const WM_USER As Long = &H400
Public Const WM_MYHOOK As Long = WM_USER + 1
Public Const WM_NOTIFY As Long = &H4E
Public Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10

' SysTray Icon API    '
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
' Windows Messaging and Windows API Declarations  '
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' defWindowProc: Variable to hold the ID of the        '
'                default Window Message processing     '
'                Procedure. Returned by SetWindowLong. '
Public defWindowProc As Long

' isSubClassed: Flag indicating that SubClassing       '
'               has been done. Provides the Means      '
'               to call the correct Message-Handler.   '
Public isSubClassed As Boolean

' colSubClassed: Collection of the TrayIcon Classes    '
Public colSubClassed As New Collection
Public lTINumCreated As Long



' Register new TrayIcon in the Collection           '
Public Sub RegisterNewTrayIcon()
  If colSubClassed.Count Then SubClass frmMain.hWnd
End Sub


' UnRegister TrayIcon in the Collection           '
Public Sub UnRegisterTrayIcon()
  If colSubClassed.Count = 0 Then UnSubClass frmMain.hWnd
End Sub



' Our own Window Message Procedure                                          '
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim bPassThrough  As Boolean
  Dim CTIcon        As clsTrayIcon
  ' Window Message Procedure                    '
  '                                             '
  ' If the Handle returned is to our Form,      '
  ' call a Form-specific Message Handler to     '
  ' deal with the Tray notifications.  If it    '
  ' is a General System Message, pass it on to  '
  ' the default Window Procedure.               '
  '                                             '
  ' If its ours, we look at lParam for the      '
  ' Message generated, and react appropriately. '
  On Error Resume Next
  
  bPassThrough = True
  For Each CTIcon In colSubClassed
    If CTIcon.hWnd = hWnd And CTIcon.ID = wParam Then
      ' Check uMsg for the Application-defined    '
      ' Identifier (NID.uID) assigned to the      '
      ' SysTray Icon in NOTIFYICONDATA (NID).     '
      ' WM_MYHOOK was defined as the Message sent '
      ' as the .uCallbackMessage Member of        '
      ' NOTIFYICONDATA the SysTray Icon           '
      If uMsg = WM_MYHOOK Then
        ' lParam is the Value of the Message      '
        ' that generated the Tray notification.   '
        ' It will be passed to the TrayIcon Class.'
        CTIcon.IconProc uMsg, wParam, lParam
        bPassThrough = False
      End If
    End If
  Next CTIcon
  
  If bPassThrough Then
    ' Handle any other Form Messages by   '
    ' passing to the default Message Proc '
    WindowProc = CallWindowProc(defWindowProc, hWnd, uMsg, wParam, lParam)
    ' This takes care of Messages when the      '
    ' Handle specified is not that of the Form  '
  End If
  
End Function

Public Function ShellTrayAdd(NID As NOTIFYICONDATA) As Long
  'Shell_NotifyIcon Messages:
  'dwMessage: Message value to send. This parameter
  '           can be one of these values:
  '           NIM_ADD     Adds icon to status area
  '           NIM_DELETE  Deletes icon from status area
  '           NIM_MODIFY  Modifies icon in status area
  '
  'pnid:      Address of the prepared NOTIFYICONDATA.
  '           The content of the structure depends
  '           on the value of dwMessage.
 
  ShellTrayAdd = Shell_NotifyIcon(NIM_ADD, NID)
End Function


Public Function ShellTrayModify(NID As NOTIFYICONDATA) As Long
  ' Modify the Icon in the Taskbar  '
  ShellTrayModify = Shell_NotifyIcon(NIM_MODIFY, NID)
End Function


Public Sub ShellTrayRemove(NID As NOTIFYICONDATA)
  ' Remove the Icon from the Taskbar  '
  Call Shell_NotifyIcon(NIM_DELETE, NID)
End Sub


Public Sub UnSubClass(ByVal hWnd As Long)
  If defWindowProc Then
    ' Restore the default Message handling before exiting '
    SetWindowLong hWnd, GWL_WNDPROC, defWindowProc
    defWindowProc = 0
  End If
End Sub


Public Sub SubClass(ByVal hWnd As Long)
  If defWindowProc = 0 Then
    ' Assign our own Window Message Procedure (WindowProc)  '
    defWindowProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
  End If
End Sub

