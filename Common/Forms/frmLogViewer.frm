VERSION 5.00
Begin VB.Form frmLogViewer 
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Logging Window"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstErrors 
      Appearance      =   0  '2D
      Height          =   1395
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox lstConsole 
      Appearance      =   0  '2D
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.ListBox lstEvents 
      Appearance      =   0  '2D
      Height          =   1395
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.OptionButton optErrors 
      Appearance      =   0  '2D
      Caption         =   "Errors"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   -60
      Width           =   1095
   End
   Begin VB.OptionButton optEvents 
      Appearance      =   0  '2D
      Caption         =   "Events"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1200
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   -60
      Width           =   1095
   End
   Begin VB.OptionButton optMessages 
      Appearance      =   0  '2D
      Caption         =   "Messages"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   -60
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Menu menuMain 
      Caption         =   "menuMain"
      Visible         =   0   'False
      Begin VB.Menu mnuDock 
         Caption         =   "&Dock"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "frmLogViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LogBuffer As New Collection
Private ErrBuffer As New Collection
Private EvtBuffer As New Collection
Private LogPaused As Boolean

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Public Sub AddLogMessage(Message As String, Optional MsgType As cxeMsgType)
  Dim i As Long, Tmp As String
  Dim oConsole As ListBox
  Dim oBuffer As Collection
  
  On Error Resume Next
  
  If Application.ShowLogger And frmMain.Visible Then
    If Me.Visible = False Then Me.Visible = True
  End If
  
  If MsgType = mEvent Then
    Set oConsole = lstEvents
    Set oBuffer = EvtBuffer
    If Not LogPaused Then
      optEvents.Value = True
    End If
  ElseIf MsgType = mError Then
    Set oConsole = lstErrors
    Set oBuffer = ErrBuffer
    If Not LogPaused Then
      optErrors.Value = True
    End If
  Else
    Set oConsole = lstConsole
    Set oBuffer = LogBuffer
    If Not LogPaused Then
      optMessages.Value = True
    End If
  End If
  
  ReleaseMessages
  
  For i = 1 To DelimiterCount(Message, vbCrLf) + 1
    Tmp = GetToken(Message, vbCrLf, i)
    If Len(Trim(Tmp)) Then
      If LogPaused Then
        oBuffer.Add Tmp
      Else
        oConsole.AddItem Tmp
      End If
    End If
  Next i
  
  While oConsole.ListCount > 1000
    oConsole.RemoveItem 0
  Wend
  
  If oConsole.ListCount And Not LogPaused Then
    oConsole.TopIndex = oConsole.ListCount - 1
  End If
End Sub

Private Sub ReleaseMessages()
  If Not LogPaused Then
    While LogBuffer.Count
      lstConsole.AddItem LogBuffer.Item(1)
      LogBuffer.Remove 1
    Wend
    While ErrBuffer.Count
      lstErrors.AddItem ErrBuffer.Item(1)
      ErrBuffer.Remove 1
    Wend
    While EvtBuffer.Count
      lstEvents.AddItem EvtBuffer.Item(1)
      EvtBuffer.Remove 1
    Wend
    If lstConsole.ListCount Then
      lstConsole.TopIndex = lstConsole.ListCount - 1
    End If
    If lstErrors.ListCount Then
      lstErrors.TopIndex = lstErrors.ListCount - 1
    End If
    If lstEvents.ListCount Then
      lstEvents.TopIndex = lstEvents.ListCount - 1
    End If
  End If
End Sub

Private Sub lstConsole_DblClick()
  MsgBox lstConsole, vbOKOnly, "Long Line View"
End Sub

Private Sub lstConsole_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu menuMain
  End If
End Sub

Private Sub lstErrors_DblClick()
  MsgBox lstErrors, vbOKOnly, "Long Line View"
End Sub

Private Sub lstErrors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu menuMain
  End If
End Sub

Private Sub lstEvents_DblClick()
  MsgBox lstEvents, vbOKOnly, "Long Line View"
End Sub

Private Sub lstEvents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu menuMain
  End If
End Sub

Private Sub mnuClear_Click()
  If optMessages.Value Then
    lstConsole.Clear
  ElseIf optEvents.Value Then
    lstEvents.Clear
  ElseIf optErrors.Value Then
    lstErrors.Clear
  End If
End Sub

Private Sub mnuDock_Click()
  mnuDock.Checked = Not mnuDock.Checked
End Sub

Private Sub mnuPause_Click()
  LogPaused = Not LogPaused
  mnuPause.Checked = LogPaused
  ReleaseMessages
End Sub

Private Sub optErrors_Click()
  DoListOnTop 2
End Sub
Private Sub optEvents_Click()
  DoListOnTop 1
End Sub
Private Sub optMessages_Click()
  DoListOnTop 0
End Sub

Private Sub DoListOnTop(ByVal Index As Long)
  On Error Resume Next
  lstConsole.Visible = Index = 0
  lstEvents.Visible = Index = 1
  lstErrors.Visible = Index = 2
  Select Case Index
    Case 0: lstConsole.SetFocus
    Case 1: lstEvents.SetFocus
    Case 2: lstErrors.SetFocus
  End Select
End Sub

