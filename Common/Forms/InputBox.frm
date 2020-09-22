VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Input"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "frmInputBox"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Visible         =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Input:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  txtInput = ""
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  Me.Hide
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    KeyAscii = 0
    cmdCancel_Click
    Exit Sub
  End If
End Sub
