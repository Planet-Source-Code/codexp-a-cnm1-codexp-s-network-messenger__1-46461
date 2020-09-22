VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserDB 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CNM User DataBase"
   ClientHeight    =   5280
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
   Icon            =   "frmUserDB.frx":0000
   LinkTopic       =   "CNMUDB"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4905
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.PictureBox picToolbar 
         Appearance      =   0  '2D
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   4785
         TabIndex        =   29
         Top             =   120
         Width           =   4815
         Begin MSComctlLib.Toolbar tbrToolbar 
            Height          =   330
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            Style           =   1
            ImageList       =   "imlTools"
            HotImageList    =   "imlToolsHot"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   16
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
                  Object.Width           =   30
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "first"
                  Description     =   "Goto first Recordset"
                  Object.ToolTipText     =   " Goto first Recordset "
                  ImageKey        =   "first"
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "previous"
                  Description     =   "Goto previous Recordset"
                  Object.ToolTipText     =   " Goto previous Recordset "
                  ImageKey        =   "previous"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "next"
                  Description     =   "Goto next Recordset"
                  Object.ToolTipText     =   " Goto next Recordset "
                  ImageKey        =   "next"
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "last"
                  Description     =   "Goto last Recordset"
                  Object.ToolTipText     =   " Goto last Recordset "
                  ImageKey        =   "last"
                  Object.Width           =   1e-4
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
                  Object.Width           =   30
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "new"
                  Description     =   "Create new Recordset"
                  Object.ToolTipText     =   " Create new Recordset "
                  ImageKey        =   "new"
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "erase"
                  Description     =   "Erase current Recordset"
                  Object.ToolTipText     =   " Erase current Recordset "
                  ImageKey        =   "erase"
               EndProperty
               BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   4
                  Object.Width           =   30
               EndProperty
               BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "requery"
                  Description     =   "Requery Data for this Recordset"
                  Object.ToolTipText     =   " Requery Data for this Recordset "
                  ImageKey        =   "requery"
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picForm 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   0
         ScaleHeight     =   4335
         ScaleWidth      =   5055
         TabIndex        =   2
         Top             =   480
         Width           =   5055
         Begin VB.TextBox txtMSN 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            MaxLength       =   50
            TabIndex        =   49
            Text            =   "ICQN"
            Top             =   2160
            Width           =   1095
         End
         Begin VB.PictureBox picFlags 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Kein
            ForeColor       =   &H00000000&
            Height          =   975
            Left            =   2640
            ScaleHeight     =   975
            ScaleWidth      =   2175
            TabIndex        =   43
            Top             =   2520
            Width           =   2175
            Begin VB.CheckBox chkUserFlag 
               Appearance      =   0  '2D
               BackColor       =   &H80000005&
               Caption         =   "Bot"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   1080
               TabIndex        =   21
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox chkUserFlag 
               Appearance      =   0  '2D
               BackColor       =   &H80000005&
               Caption         =   "Service"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   1080
               TabIndex        =   19
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chkUserFlag 
               Appearance      =   0  '2D
               BackColor       =   &H80000005&
               Caption         =   "Operator"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   17
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox chkUserFlag 
               Appearance      =   0  '2D
               BackColor       =   &H80000005&
               Caption         =   "Admin"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   16
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox chkUserFlag 
               Appearance      =   0  '2D
               BackColor       =   &H80000005&
               Caption         =   "Invisible"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   18
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chkUserFlag 
               Appearance      =   0  '2D
               BackColor       =   &H80000005&
               Caption         =   "Clan"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   20
               Top             =   720
               Width           =   615
            End
            Begin VB.Label lblCap 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User Flags:"
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   810
            End
            Begin VB.Label lblFlags 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   195
               Left            =   960
               TabIndex        =   44
               Top             =   0
               Width           =   90
            End
         End
         Begin VB.TextBox txtNick 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "Nickname"
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtAge 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "99"
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox txtRegDate 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "00/00/0000"
            ToolTipText     =   " date of registration "
            Top             =   3840
            Width           =   2175
         End
         Begin VB.OptionButton optGenderF 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            Caption         =   "female"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4080
            TabIndex        =   23
            Top             =   3615
            Width           =   855
         End
         Begin VB.OptionButton optGenderM 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            Caption         =   "male"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3360
            TabIndex        =   22
            Top             =   3615
            Width           =   735
         End
         Begin VB.TextBox txtBDate 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "00/00/0000"
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtICQN 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   15
            Text            =   "ICQN"
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtPhone 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   14
            Text            =   "Telephone"
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtEMail 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   13
            Text            =   "E-Mail Address"
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtAddress1 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "Address 1"
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox txtAddress2 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   7
            Text            =   "Address 2"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox txtAddress3 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "Address 3"
            Top             =   2040
            Width           =   2175
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Save"
            Height          =   375
            Left            =   3960
            Style           =   1  'Grafisch
            TabIndex        =   24
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtUserName 
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2640
            MaxLength       =   50
            TabIndex        =   12
            Text            =   "Name of User"
            ToolTipText     =   " name and surname "
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtUserID 
            Alignment       =   2  'Zentriert
            Appearance      =   0  '2D
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   3
            Text            =   "User"
            ToolTipText     =   " user id for login "
            Top             =   360
            Width           =   975
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
            TabIndex        =   4
            Text            =   "Password"
            ToolTipText     =   " user password for login "
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optNobody 
            Appearance      =   0  '2D
            BackColor       =   &H80000005&
            Caption         =   "Nobody"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3840
            TabIndex        =   39
            Top             =   2640
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MSN ID:"
            Height          =   195
            Index           =   16
            Left            =   3720
            TabIndex        =   50
            Top             =   1920
            Width           =   585
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   19
            Left            =   1320
            TabIndex        =   48
            Top             =   360
            Width           =   120
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   18
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   120
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "use 'now' for current date!"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   46
            Top             =   4095
            Width           =   1920
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nickname:"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   42
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age:"
            Height          =   195
            Index           =   14
            Left            =   2040
            TabIndex        =   41
            Top             =   2400
            Width           =   345
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of registration:"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   40
            Top             =   3600
            Width           =   1470
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender:"
            Height          =   195
            Index           =   12
            Left            =   2640
            TabIndex        =   38
            Top             =   3600
            Width           =   585
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date of birth:"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   37
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ICQ Number:"
            Height          =   195
            Index           =   10
            Left            =   2640
            TabIndex        =   36
            Top             =   1920
            Width           =   945
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone:"
            Height          =   195
            Index           =   9
            Left            =   2640
            TabIndex        =   35
            Top             =   1320
            Width           =   810
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail Adress:"
            Height          =   195
            Index           =   8
            Left            =   2640
            TabIndex        =   34
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full Name:"
            Height          =   195
            Index           =   5
            Left            =   2640
            TabIndex        =   32
            Top             =   120
            Width           =   750
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UserID:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   31
            Top             =   120
            Width           =   555
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
      End
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":0442
            Key             =   "first"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":0796
            Key             =   "last"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":0AEA
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":0E3E
            Key             =   "next"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":1192
            Key             =   "previous"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":14E6
            Key             =   "erase"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":183A
            Key             =   "requery"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrUDBControl 
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ImageList imlToolsHot 
      Left            =   1080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":1B8E
            Key             =   "first"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":1EE2
            Key             =   "last"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":2236
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":258A
            Key             =   "next"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":28DE
            Key             =   "previous"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":2C32
            Key             =   "erase"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserDB.frx":2F86
            Key             =   "requery"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE IS NOT OPENED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Top             =   1080
      Width           =   4275
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE IS NOT OPENED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   495
      TabIndex        =   27
      Top             =   1095
      Width           =   4275
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATABASE IS NOT OPENED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   465
      TabIndex        =   26
      Top             =   1065
      Width           =   4275
   End
End
Attribute VB_Name = "frmUserDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bNewRecordset As Boolean
Private bRSChanged    As Boolean

Private Sub EnableToolbar(ByVal bStatus As Boolean)
  Dim i As Long
  For i = 1 To tbrToolbar.Buttons.Count
    tbrToolbar.Buttons(i).Enabled = bStatus
  Next i
End Sub

Private Sub chkUserFlag_Click(Index As Integer)
  lblFlags = CalculateFlags
  bRSChanged = True
End Sub

Public Function CalculateFlags() As Long
  Dim lFlags As Long
  Dim i As Long
  For i = 0 To chkUserFlag.UBound
    lFlags = lFlags Or (2 ^ i * Abs(chkUserFlag(i).Value))
  Next i
  CalculateFlags = lFlags
End Function

Public Sub SetFlags(ByVal lFlags As Long)
  Dim i As Long
  For i = 0 To chkUserFlag.UBound
    chkUserFlag(i).Value = Abs((lFlags And (2 ^ i)) = (2 ^ i))
  Next i
End Sub

Private Sub cmdSave_Click()
  Save
End Sub

Private Sub Form_Load()
  UpdateForm
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
  On Local Error Resume Next
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
    frmMain.SetFocus
  End If
End Sub

' AskToSaveRS() - Returns True if Canceled  '
Public Function AskToSaveRS() As Boolean
  Dim lSave As Long
  Dim Tmp As String
  If bNewRecordset Or bRSChanged Then
    ' Ask to Save new/changed Recordset '
    Tmp = "Do you want to Save this Recordset?"
    lSave = MsgYesNoCancel(Tmp, "Save this Recordset?")
    If lSave < 2 Then
      AskToSaveRS = True
    Else
      If lSave = 3 Then
        Save ' Save Recordset  '
      End If
      ClearForm
    End If
  End If
End Function

Private Sub optGenderF_Click()
  bRSChanged = True
End Sub

Private Sub optGenderM_Click()
  bRSChanged = True
End Sub

Private Sub tbrToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim bRSNavigation As Boolean
  
  Select Case Button.Key
    Case "requery"
      UserDB.RSClients.Requery
      UpdateForm
    Case "erase"
      EraseCurrentRS
    Case Else
      bRSNavigation = True
  End Select
  If Not bRSNavigation Then Exit Sub
  
  If AskToSaveRS Then Exit Sub
  If bNewRecordset Or bRSChanged Then Exit Sub
  ' Recordset Navigation    '
  If UserDB.RSClients Is Nothing Then Exit Sub
  Select Case Button.Key
    Case "new"
      ClearForm
      txtRegDate = "now"
      bNewRecordset = True
    Case "previous"
      If UserDB.RSClients.RecordCount Then
        UserDB.RSClients.MovePrevious
      End If
      UpdateForm
    Case "next"
      If UserDB.RSClients.RecordCount Then
        UserDB.RSClients.MoveNext
      End If
      UpdateForm
    Case "first"
      If UserDB.RSClients.RecordCount Then
        UserDB.RSClients.MoveFirst
      End If
      UpdateForm
    Case "last"
      If UserDB.RSClients.RecordCount Then
        UserDB.RSClients.MoveLast
      End If
      UpdateForm
  End Select
End Sub

Public Function Save() As Long
  Dim Tmp As String
  Dim dBDateV As Double
  Dim dRegDate As Double
  
  On Error GoTo Save_Resume
  
  If UserDB.RSClients Is Nothing Then Exit Function
  If UserDB.RSClients.AbsolutePosition < 0 Then bNewRecordset = True
  
  If CheckUserID(txtUserID) Then
    Tmp = "UserID is invalid or empty!"
    Save = 1
  End If
  If CheckPassword(txtPassword) Then
    Tmp = "Password is invalid or empty!"
    Save = 2
  End If
  If bNewRecordset And UserDB.UserExists(txtUserID) Then
    Tmp = "User already exists in the DataBase!"
    Save = 3
  End If
  If UserDB.NickNameExists(txtNick, txtUserID) Then
    Tmp = "This Nickname is already used!"
    Save = 4
  End If
  If Len(Tmp) Then
    MsgInformation Tmp, "Error:", True
    Exit Function
  End If
  
  If Len(Trim(txtBDate)) Then
111 txtBDate = CDate(txtBDate)
112 dBDateV = TimeToLong(txtBDate)
    txtAge = YearsLeftAfter(dBDateV)
  End If
  If Len(Trim(txtRegDate)) Then
    If UCase(Trim(txtRegDate)) = "NOW" Then txtRegDate = Now
113 txtRegDate = CDate(txtRegDate)
114 dRegDate = TimeToLong(txtRegDate)
  End If
  
  If bNewRecordset Then
    UserDB.RSClients.AddNew
  Else
    UserDB.RSClients.Edit
  End If
  
  With UserDB.RSClients
    .Fields("User") = txtUserID
    .Fields("Password") = Trim(txtPassword)
    .Fields("Nickname") = Trim(txtNick)
    .Fields("Username") = Trim(txtUserName)
    .Fields("Address1") = Trim(txtAddress1)
    .Fields("Address2") = Trim(txtAddress2)
    .Fields("Address3") = Trim(txtAddress3)
    .Fields("EMail") = Trim(txtEMail)
    .Fields("ICQN") = Trim(txtICQN)
    .Fields("MSNID") = Trim(txtMSN)
    .Fields("Phone") = Trim(txtPhone)
    .Fields("Flags") = Val(CalculateFlags)
    .Fields("BDate") = dBDateV
    .Fields("Gender") = GetGenderValue
    .Fields("Registered") = dRegDate
    .Update
  End With
  
  If bNewRecordset Then UserDB.RSClients.MoveLast
  bNewRecordset = False
  bRSChanged = False

Save_Resume:
  If Err Then
    Tmp = Err.Description
    If Erl = 111 Or Erl = 113 Then
      Tmp = "Please enter a valid Date or Nothing!"
    End If
    MsgInformation Tmp, "Error:", True
  End If
  Save = Err
End Function

Private Sub EraseCurrentRS()
  Static bExit As Boolean
  Dim lRet As Long
  
  On Error Resume Next
  
  If bExit Or UserDB.RSClients Is Nothing Then Exit Sub
  bExit = True
  
  lRet = MsgQuestion("Do you realy want to erase current Recordset?", "Realy erase?")
  If lRet = 2 Then ' If Yes was pressed '
    If Not bNewRecordset Then
      If UserDB.RSClients.AbsolutePosition >= 0 Then
        UserDB.RSClients.Delete
      End If
    End If
    ClearForm
    If UserDB.RSClients.RecordCount Then UserDB.RSClients.MoveLast
    UpdateForm
  End If
  
  bExit = False
End Sub

Private Sub tmrUDBControl_Timer()
  Dim bState As Boolean
  Dim Tmp As String
  
  picContainer.Visible = UserDB.IsOpen
  picForm.Visible = Not UserDB.RSClients Is Nothing
  If UserDB.IsOpen Then
    If UserDB.RSClients Is Nothing Then
      Tmp = "Account Recordsets are not available!"
      If lblStat <> Tmp Then lblStat = Tmp
    Else
      If bNewRecordset Then
        Tmp = "New Recordset"
      Else
        Tmp = "Recordset " & UserDB.RSClients.AbsolutePosition + 1 & "/" & UserDB.RSClients.RecordCount
      End If
      If lblStat <> Tmp Then lblStat = Tmp
    End If
  Else
    If lblStat <> "" Then lblStat = ""
  End If
  cmdSave.Enabled = bNewRecordset Or bRSChanged
End Sub

Public Sub UpdateForm()
  On Error Resume Next
  
  If UserDB.RSClients Is Nothing Then Exit Sub
  
  If bNewRecordset Or bRSChanged Then
    If AskToSaveRS Then Exit Sub
  End If
  ClearForm
  
  If UserDB.RSClients.AbsolutePosition < 0 Then
    If UserDB.RSClients.RecordCount > 0 Then
      UserDB.RSClients.MoveFirst
    Else
      Exit Sub
    End If
  End If
  
  txtUserID = UserDB.RSClients("User").Value
  txtPassword = UserDB.RSClients("Password").Value
  txtNick = UserDB.RSClients("Nickname").Value
  txtUserName = UserDB.RSClients("Username").Value
  txtAddress1 = UserDB.RSClients("Address1").Value
  txtAddress2 = UserDB.RSClients("Address2").Value
  txtAddress3 = UserDB.RSClients("Address3").Value
  txtEMail = UserDB.RSClients("EMail").Value
  txtPhone = UserDB.RSClients("Phone").Value
  txtICQN = UserDB.RSClients("ICQN").Value
  txtMSN = UserDB.RSClients("MSNID").Value
  If UserDB.RSClients("Flags").Value > 0 Then
    Me.SetFlags UserDB.RSClients("Flags").Value
  End If
  If UserDB.RSClients("BDate").Value > 0 Then
    txtBDate = LongToTime(UserDB.RSClients("BDate").Value)
    txtAge = YearsLeftAfter(TimeToLong(txtBDate))
  End If
  If UserDB.RSClients("Gender").Value > 0 Then
    SetGenderValue UserDB.RSClients("Gender").Value
  End If
  If UserDB.RSClients("Registered").Value > 0 Then
    txtRegDate = LongToTime(UserDB.RSClients("Registered").Value)
  End If
  
  bRSChanged = False
End Sub

Private Sub ClearForm()
  txtUserID = ""
  txtPassword = ""
  txtUserName = ""
  txtAddress1 = ""
  txtAddress2 = ""
  txtAddress3 = ""
  txtEMail = ""
  txtPhone = ""
  txtICQN = ""
  txtMSN = ""
  txtBDate = ""
  txtRegDate = ""
  txtAge = ""
  txtNick = ""
  SetGenderValue 0
  SetFlags 0
  bNewRecordset = False
  bRSChanged = False
End Sub

Private Sub txtAddress1_Change()
  bRSChanged = True
End Sub

Private Sub txtAddress2_Change()
  bRSChanged = True
End Sub

Private Sub txtAddress3_Change()
  bRSChanged = True
End Sub

Private Sub txtBDate_Change()
  bRSChanged = True
End Sub

Private Sub txtEMail_Change()
  bRSChanged = True
End Sub

Private Sub txtICQN_Change()
  bRSChanged = True
End Sub

Private Sub txtMSN_Change()
  bRSChanged = True
End Sub

Private Sub txtNick_Change()
  bRSChanged = True
End Sub

Private Sub txtPassword_Change()
  txtPassword.ToolTipText = " " & txtPassword & " "
  bRSChanged = True
End Sub

Private Sub txtPhone_Change()
  bRSChanged = True
End Sub

Private Sub txtRegDate_Change()
  bRSChanged = True
End Sub

Private Sub txtUserID_Change()
  bRSChanged = True
End Sub

Private Sub txtUserName_Change()
  bRSChanged = True
End Sub

