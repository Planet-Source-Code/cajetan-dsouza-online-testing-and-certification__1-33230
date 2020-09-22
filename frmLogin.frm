VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Login Screen"
   ClientHeight    =   5100
   ClientLeft      =   -4425
   ClientTop       =   -1410
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUserID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   2
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   900
      TabIndex        =   10
      Top             =   3840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Max             =   1000
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   5280
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login Screen"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "User ID :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   8
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "### *******Online Testing and Certification *******###"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   840
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00C00000&
      Height          =   4815
      Left            =   120
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim com As Command
Dim rst As Recordset
Private Sub Form_Load()
   Set conn = New Connection
   Set com = New Command
   Set rst = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With
   rst.Open "Login", conn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim NextTabIndex As Integer, i As Integer
   If KeyAscii = 13 Then
      If Screen.ActiveControl.TabIndex = _
      Count - 1 Then
         NextTabIndex = 0
      Else
         NextTabIndex = Screen.ActiveControl.TabIndex + 1
      End If
      For i = 0 To Count - 1
         If Me.Controls(i).TabIndex = NextTabIndex Then
            Me.Controls(i).SetFocus
            Exit For
         End If
      Next i
      KeyAscii = 0
   End If
End Sub

Private Sub cmdOK_Click()
Dim uname As String
Dim Pass As String
Dim flag As Boolean

Pass = txtPassword.Text

rst.MoveFirst
Do While Not rst.EOF
 If rst!vUserName = txtUserID Then
  If rst!vPassword = Pass Then
      rst!CurrentLogin = 1
      rst.Update
      For i = 0 To 1000
        ProgressBar1.Value = i
      Next
      frmMain.Show
      Unload Me
      Exit Do
   Else
      MsgBox "Invalid Password.", vbExclamation, "Authentication"
      frmLogin.SetFocus
      txtPassword.SetFocus
      SendKeys "{Home}+{End}"
      Exit Do
   End If
 End If
   rst.MoveNext
Loop
End Sub

