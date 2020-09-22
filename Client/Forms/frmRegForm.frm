VERSION 5.00
Begin VB.Form frmRegForm 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "User Registration"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
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
   ScaleHeight     =   4560
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrControl 
      Interval        =   300
      Left            =   4560
      Top             =   240
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdNewUser 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&New User"
         Height          =   375
         Left            =   3360
         Style           =   1  'Grafisch
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Query User Data"
         Height          =   375
         Left            =   240
         Style           =   1  'Grafisch
         TabIndex        =   33
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmdRegister 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Register"
         Height          =   375
         Left            =   3360
         Style           =   1  'Grafisch
         TabIndex        =   18
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Clear Form"
         Height          =   375
         Left            =   1800
         Style           =   1  'Grafisch
         TabIndex        =   17
         Top             =   3720
         Width           =   1455
      End
      Begin VB.PictureBox picForm 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   0
         ScaleHeight     =   3135
         ScaleWidth      =   5055
         TabIndex        =   19
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtMSN 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            MaxLength       =   50
            TabIndex        =   13
            Text            =   "ICQN"
            Top             =   2160
            Width           =   1095
         End
         Begin VB.OptionButton optNobody 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            Caption         =   "Nobody"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3720
            TabIndex        =   14
            Top             =   2520
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtPassword 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1440
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   2
            Text            =   "Password"
            ToolTipText     =   " user password for login "
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtUserID 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   1
            Text            =   "User"
            ToolTipText     =   " user id for login "
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtUserName 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   9
            Text            =   "Name of User"
            ToolTipText     =   " name and surname "
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtAddress3 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "Address 3"
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txtAddress2 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "Address 2"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox txtAddress1 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Address 1"
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtEMail 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   10
            Text            =   "E-Mail Address"
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Telephone"
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtICQN 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   12
            Text            =   "ICQN"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtBDate 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   30
            TabIndex        =   7
            Text            =   "00/00/0000"
            Top             =   2640
            Width           =   1695
         End
         Begin VB.OptionButton optGenderM 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            Caption         =   "male"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3360
            TabIndex        =   15
            Top             =   2655
            Width           =   735
         End
         Begin VB.OptionButton optGenderF 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            Caption         =   "female"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4080
            TabIndex        =   16
            Top             =   2655
            Width           =   855
         End
         Begin VB.TextBox txtAge 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "99"
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox txtNick 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Nickname"
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MSN ID:"
            Height          =   195
            Index           =   0
            Left            =   3720
            TabIndex        =   32
            Top             =   1920
            Width           =   585
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   195
            Index           =   4
            Left            =   1440
            TabIndex        =   30
            Top             =   120
            Width           =   750
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User ID:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   29
            Top             =   120
            Width           =   600
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full Name:"
            Height          =   195
            Index           =   5
            Left            =   2640
            TabIndex        =   28
            Top             =   120
            Width           =   750
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail Adress:"
            Height          =   195
            Index           =   8
            Left            =   2640
            TabIndex        =   26
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Index           =   9
            Left            =   2640
            TabIndex        =   25
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICQ Number:"
            Height          =   195
            Index           =   10
            Left            =   2640
            TabIndex        =   24
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of birth:"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   23
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender:"
            Height          =   195
            Index           =   12
            Left            =   2640
            TabIndex        =   22
            Top             =   2640
            Width           =   585
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age:"
            Height          =   195
            Index           =   14
            Left            =   2040
            TabIndex        =   21
            Top             =   2400
            Width           =   345
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nickname:"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Form"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmRegForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bEnableRegButton As Boolean



Private Sub cmdClear_Click()
  ClearForm
End Sub


Private Sub cmdNewUser_Click()
  If Client.LoggedIn > 0 Then
    frmMain.SendCL ":" & Client.User & " LOGOUT"
  End If
End Sub


Private Sub cmdQuery_Click()
  If Trim(txtUserID) = "" Then
    MsgInformation "To show User Informations input an UserID first!"
  Else
    frmMain.SendCL "QUERY USER " & txtUserID & " INFO"
  End If
End Sub


Private Sub cmdRegister_Click()
  Dim Tmp As String
  Dim Tmr As Long
  
  If Client.Connected = 0 Then
    bDontAutoLogin = True
    frmMain.ConnectToServer , True
    Exit Sub
  End If
  
  cmdRegister.Enabled = False
  ' Send all Data to Server   '
  With frmMain
    Tmp = "REG FILL USERID:" & Trim(txtUserID) & vbCrLf & _
          "REG FILL PASSWORD:" & Trim(txtPassword) & vbCrLf & _
          "REG FILL NICKNAME:" & Trim(txtNick) & vbCrLf & _
          "REG FILL USERNAME:" & Trim(txtUserName) & vbCrLf & _
          "REG FILL ADDRESS1:" & Trim(txtAddress1) & vbCrLf & _
          "REG FILL ADDRESS2:" & Trim(txtAddress2) & vbCrLf & _
          "REG FILL ADDRESS3:" & Trim(txtAddress3) & vbCrLf & _
          "REG FILL EMAIL:" & Trim(txtEMail) & vbCrLf & _
          "REG FILL MSNID:" & Trim(txtMSN) & vbCrLf & _
          "REG FILL ICQN:" & Trim(txtICQN) & vbCrLf & _
          "REG FILL BDATE:" & Trim(txtBDate) & vbCrLf & _
          "REG FILL PHONE:" & Trim(txtPhone) & vbCrLf & _
          "REG FILL GENDER:" & GetGenderValue
    .SendCL Tmp
    If Client.LoggedIn Then
      .SendCL "REG UPDATE"
    Else
      .SendCL "REG REGISTER"
    End If
  End With
  
  bEnableRegButton = True
End Sub


Private Sub Form_Load()
  ClearForm
  txtUserID = Client.User
  txtPassword = Client.Password
  If Len(Trim(txtUserID)) And Client.Connected > 0 Then
    cmdQuery_Click
  End If
End Sub


Public Sub ClearForm()
  txtUserID = ""
  txtPassword = ""
  txtUserName = ""
  txtAddress1 = ""
  txtAddress2 = ""
  txtAddress3 = ""
  txtEMail = ""
  txtMSN = ""
  txtPhone = ""
  txtICQN = ""
  txtBDate = ""
  txtAge = ""
  txtNick = ""
  SetGenderValue 0
End Sub


Public Function GetGenderValue() As Integer
  Dim iRet As Integer
  If optGenderM.Value Then iRet = 1
  If optGenderF.Value Then iRet = 2
  GetGenderValue = iRet
End Function


Public Sub SetGenderValue(ByVal Value As Integer)
  optNobody.Value = True
  If Value = 1 Then optGenderM.Value = True
  If Value = 2 Then optGenderF.Value = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    Me.Hide
    Cancel = True
  End If
End Sub


Private Sub tmrControl_Timer()
  Static lERBT  As Long
  Dim Tmp       As String
  If bEnableRegButton Then
    If lERBT = 0 Then lERBT = TickSeconds + 5
    If lERBT < TickSeconds Then
      lERBT = 0
      cmdRegister.Enabled = True
      bEnableRegButton = False
    End If
  End If
  
  cmdNewUser.Visible = Client.LoggedIn > 0
  If Client.Connected > 0 Then
    If Client.LoggedIn > 0 Then
      Tmp = "&Update Data"
    Else
      Tmp = "&Register"
    End If
  Else
    Tmp = "&Connect"
  End If
  If cmdRegister.Caption <> Tmp Then cmdRegister.Caption = Tmp
End Sub


Private Sub txtBDate_Change()
  Dim lAge As Long
  On Error GoTo BDate_Change_Error
  lAge = YearsLeftAfter(TimeToLong(CDate(txtBDate)))
  Debug.Print lAge
  If lAge > 1 And lAge < 100 Then
    txtAge = lAge
  Else
    txtAge = ""
  End If
  
BDate_Change_Error:
  If Err Then txtAge = ""
End Sub


Private Sub txtPassword_Change()
  txtPassword.ToolTipText = " " & txtPassword & " "
End Sub
