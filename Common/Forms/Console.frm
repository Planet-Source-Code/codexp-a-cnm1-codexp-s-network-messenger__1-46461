VERSION 5.00
Begin VB.Form frmConsole 
   Caption         =   "Console"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Console.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.PictureBox picHolder 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   960
      ScaleHeight     =   795
      ScaleWidth      =   1875
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
      Begin VB.PictureBox picIcon 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   720
         Picture         =   "Console.frx":058A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   120
         Picture         =   "Console.frx":0B14
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   510
      End
   End
   Begin VB.ComboBox txtLine 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   0
      Style           =   1  'Einfaches Kombinationsfeld
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
   End
   Begin VB.PictureBox picCon 
      BackColor       =   &H00808080&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.Timer tmrControl 
         Interval        =   500
         Left            =   120
         Top             =   120
      End
      Begin VB.ListBox lstUsers 
         Appearance      =   0  '2D
         Height          =   2175
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox lstCon 
         Appearance      =   0  '2D
         Height          =   2175
         ItemData        =   "Console.frx":13DE
         Left            =   0
         List            =   "Console.frx":13E0
         TabIndex        =   2
         Top             =   0
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bPaused As Boolean
Public bChannel As Boolean
Public bConsole As Boolean
Public User As String
Public Chan As clsChannel

Private Sub Form_Load()
  tmrControl_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If bChannel Then
    If Not Chan Is Nothing Then
      If Len(Trim(Chan.Caption)) Then
        frmMain.SendCL ":" & Client.User & " PART " & Chan.Caption
      End If
    End If
  Else
    If UnloadMode = 0 Then
      Cancel = True
      Me.Hide
    End If
  End If
End Sub

Private Sub Form_Resize()
  Dim W As Single
  Dim H As Single
  
  If bChannel Then
    If Not lstUsers.Visible Then lstUsers.Visible = True
  End If
  
  H = Me.ScaleHeight
  W = Me.ScaleWidth
  If H < 1500 Then H = 1500
  If W < 3000 Then W = 3000
  picCon.Move 0, 0, W, H - txtLine.Height
  txtLine.Move 0, picCon.Height, W
  If lstUsers.Visible Then
    lstCon.Move 0, 0, picCon.ScaleWidth - lstUsers.Width - 60, picCon.ScaleHeight
    lstUsers.Move lstCon.Width + 30, 0, lstUsers.Width, picCon.ScaleHeight
  Else
    lstCon.Move 0, 0, picCon.ScaleWidth, picCon.ScaleHeight
  End If
End Sub

Public Sub AddMessage(ByVal Message As String)
  Dim Tmp As String
  Dim i As Long
  Dim n As Long
  
  n = DelimiterCount(Message, vbCrLf) + 1
  For i = 1 To n
    Tmp = GetToken(Message, vbCrLf, i)
    If Not (Len(Trim(Tmp)) = 0 And i = n) Then
      lstCon.AddItem Tmp
    End If
  Next i
  
  While lstCon.ListCount > 1000
    lstCon.RemoveItem 0
  Wend
  
  If lstCon.ListCount And Not bPaused Then
    lstCon.TopIndex = lstCon.ListCount - 1
  End If
End Sub

Private Sub lstCon_DblClick()
  MsgInformation lstCon
End Sub

Private Sub lstUsers_DblClick()
  OpenPrivate lstUsers
End Sub

Private Sub tmrControl_Timer()
  Dim Tmp As String
  If bChannel Then
    If Chan Is Nothing Then
      Tmp = "{Unknown Channel}"
    Else
      Tmp = Chan.Caption
    End If
  Else
    Tmp = User
  End If
  If Len(Trim(Tmp)) Then
    If Me.Caption <> Trim(Tmp) Then
      Me.Caption = Trim(Tmp)
    End If
  End If
End Sub

Private Sub txtLine_KeyPress(KeyAscii As Integer)
  Dim Tmp As String
  Select Case KeyAscii
    Case 10, 13
      KeyAscii = 0
      Tmp = RTrim(txtLine)
      If Trim(Tmp) = "" Then
        txtLine = ""
        Exit Sub
      End If
      If bConsole Then Tmp = PrependSlash(Tmp)
      If Left(Tmp, 1) = "/" Or Left(Tmp, 1) = "\" Then
        Tmp = Mid(Tmp, 2)
        ' Execute Local Command Line  '
        Select Case UCase(GetToken(Tmp, " ", 1))
          Case "CLEAR": lstCon.Clear
          Case Else
            frmMain.SendCL Tmp
        End Select
        CaptureLine
      Else
        If Client.LoggedIn > 0 And frmMain.WS.State = 7 And Len(Trim(User)) Then
          AddMessage "<" & Client.User & "> " & Tmp
          frmMain.SendCL ":" & Client.User & " PRIVMSG " & User & ":" & Tmp
          CaptureLine
        ElseIf Client.LoggedIn > 0 And frmMain.WS.State = 7 And bChannel Then
          If Len(Trim(Chan.Caption)) Then
            If Left(Chan.Caption, 1) = "#" Then Chan.Caption = Trim(Mid(Chan.Caption, 2))
            If Len(Chan.Caption) Then
              frmMain.SendCL ":" & Client.User & " PRIVMSG #" & Chan.Caption & ":" & Tmp
            End If
            CaptureLine
          End If
        Else
          Beep
        End If
      End If
  End Select
End Sub

Private Sub CaptureLine()
  Dim Tmp As String
  Tmp = RTrim(txtLine)
  txtLine = ""
  If Trim(Tmp) = "" Then Exit Sub
  txtLine.AddItem Tmp, 0
  While txtLine.ListCount > 10
    txtLine.RemoveItem txtLine.ListCount - 1
  Wend
End Sub

Public Sub AddUser(ByVal User As String, Optional Flags As Long)
  If Chan.Users.AddAs(User, Flags) = 0 Then
    lstUsers.AddItem "  " & User
  End If
End Sub

Public Sub RemoveUser(ByVal User As String)
  Dim i As Long
  Chan.Users.Remove User
  For i = 0 To lstUsers.ListCount - 1
    If UCase(Trim(lstUsers.List(i))) = UCase(Trim(User)) Then
      lstUsers.RemoveItem i
    End If
  Next i
End Sub

Public Sub ParseUserList(ByVal UserList As String)
  Dim UsrL As String
  Dim Usr As String
  Dim Flag As Long
  Dim i As Long
  
  lstUsers.Clear
  Set Chan.Users = New clsChanUsers
  UsrL = Trim(UserList)
  If UsrL = "" Then Exit Sub
  For i = 1 To DelimiterCount(UsrL, ",") + 1
    Usr = Trim(GetToken(GetToken(UsrL, ",", i), "|", 1))
    Flag = Val(GetToken(GetToken(UsrL, ",", i), "|", 2))
    If Len(Usr) Then
      AddUser Usr, Flag
    End If
  Next i
End Sub

