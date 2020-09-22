VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4185
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2160
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7065
      Begin VB.PictureBox Picture1 
         Height          =   3615
         Left            =   0
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   3555
         ScaleWidth      =   1875
         TabIndex        =   9
         Top             =   0
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   6960
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Left            =   2160
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright : Cajetan D'souza"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Organization : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5760
         TabIndex        =   3
         Top             =   3270
         Width           =   1215
      End
      Begin VB.Label lblWarning 
         Caption         =   $"frmSplash.frx":1EA4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   2
         Top             =   3600
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version : v1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5130
         TabIndex        =   5
         Top             =   2700
         Width           =   1755
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform : Windows 95/98/Me/XP/2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2580
         TabIndex        =   6
         Top             =   2340
         Width           =   4275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Online Testing and Certification"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1230
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Licensed To : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   3000
         Width           =   5055
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "CajieSoft Infotech Ltd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   720
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If Not GetSetting("CajieSoft", "Registration", "User name") = "" Then
     lblLicenseTo.Caption = lblLicenseTo.Caption + GetSetting("CajieSoft", "Registration", "User name")
     lblCompany.Caption = lblCompany.Caption + GetSetting("CajieSoft", "Registration", "Organization")
   Else
     lblLicenseTo.Caption = lblLicenseTo.Caption + "<< Unlicensed Copy >>"
     lblCompany.Caption = lblCompany.Caption + "<< Unlicensed Copy >>"
   End If
End Sub

Private Sub Timer1_Timer()
  frmLogin.Show
  Unload Me
End Sub
