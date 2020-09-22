VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Dialog"
   ClientHeight    =   1830
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.PictureBox picInformation 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   120
      Picture         =   "Dialog.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   240
      Width           =   480
   End
   Begin VB.TextBox Message 
      Alignment       =   2  'Zentriert
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      Locked          =   -1  'True
      MousePointer    =   1  'Pfeil
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Dialog.frx":08CA
      Top             =   240
      Width           =   3975
   End
   Begin VB.CommandButton cmdButton3 
      Caption         =   "Retry"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox picQuestion 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   120
      Picture         =   "Dialog.frx":08D4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picCritical 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   120
      Picture         =   "Dialog.frx":119E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ReturnValue As Long

Private Sub cmdButton1_Click()
  ReturnValue = 1
  Me.Hide
End Sub

Private Sub cmdButton1_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 32, 13, 10
    Case Else
      KeyPress KeyAscii
  End Select
End Sub

Private Sub cmdButton2_Click()
  ReturnValue = 2
  Me.Hide
End Sub

Private Sub cmdButton2_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 32, 13, 10
    Case Else
      KeyPress KeyAscii
  End Select
End Sub

Private Sub cmdButton3_Click()
  ReturnValue = 3
  Me.Hide
End Sub

Public Sub ResetForm()
  ReturnValue = 0
  picInformation.Visible = True
  picQuestion.Visible = False
  picCritical.Visible = False
  cmdButton3.Visible = False
  cmdButton2.Visible = False
  cmdButton1.Visible = True
  cmdButton1.Default = True
  cmdButton1.Caption = "Ok"
  cmdButton2.Caption = "Cancel"
  cmdButton3.Caption = "Ignore"
  Me.Caption = "Dialog"
  Message = "Message"
End Sub

Private Sub cmdButton3_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 32, 13, 10
    Case Else
      KeyPress KeyAscii
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  KeyPress KeyAscii
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ReturnValue = 0
  If UnloadMode = 0 Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub Message_KeyPress(KeyAscii As Integer)
  KeyPress KeyAscii
End Sub

Private Sub picCritical_KeyPress(KeyAscii As Integer)
  KeyPress KeyAscii
End Sub

Private Sub picQuestion_KeyPress(KeyAscii As Integer)
  KeyPress KeyAscii
End Sub

Private Sub KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 27: Me.Hide
    Case Else
      If InStr(UCase(cmdButton1.Caption), "&" & UCase(Chr(KeyAscii))) Then
        cmdButton1.Value = True
      ElseIf InStr(UCase(cmdButton2.Caption), "&" & UCase(Chr(KeyAscii))) Then
        cmdButton2.Value = True
      ElseIf InStr(UCase(cmdButton3.Caption), "&" & UCase(Chr(KeyAscii))) Then
        cmdButton3.Value = True
      End If
  End Select
End Sub
