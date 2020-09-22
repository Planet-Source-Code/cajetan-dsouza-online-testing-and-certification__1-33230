VERSION 5.00
Begin VB.Form frmChangePassword 
   Caption         =   "Change Password"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5513
      TabIndex        =   6
      Text            =   "Administrator"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtOldPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   5513
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3165
      Width           =   2895
   End
   Begin VB.TextBox txtNewPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   5513
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3795
      Width           =   2895
   End
   Begin VB.TextBox txtVerifyPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   5513
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4440
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6473
      TabIndex        =   4
      Top             =   5700
      Width           =   2415
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "C&hange"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3113
      TabIndex        =   3
      Top             =   5700
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4110
      TabIndex        =   10
      Top             =   2640
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Old Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3840
      TabIndex        =   9
      Top             =   3285
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "New Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3750
      TabIndex        =   8
      Top             =   3915
      Width           =   1665
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Verify Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   660
      X2              =   11340
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4140
      TabIndex        =   5
      Top             =   1080
      Width           =   3720
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim com As Command
Dim rst As Recordset

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

Private Sub Form_Load()
   Set conn = New Connection
   Set com = New Command
   Set rst = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   com.ActiveConnection = conn
   rst.Open "login", conn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmdCancel_Click()
  frmMain.SetFocus
  Unload Me
End Sub

Private Sub cmdChange_Click()
  rst.MoveFirst
  Do While Not rst.EOF
    If rst!vUserName = txtUserName.Text Then
      If rst!vPassword = txtOldPassword.Text Then
        If txtNewPassword.Text = txtVerifyPassword.Text Then
           rst("vPassword") = txtNewPassword.Text
           MsgBox "Password successfully changed", vbInformation
           frmMain.SetFocus
           Unload Me
        Else
           MsgBox "Passwords do not match", vbExclamation
           txtVerifyPassword.SetFocus
           SendKeys "{Home}+{End}"
        End If
     Else
       MsgBox "Invalid Password.", vbExclamation
       txtOldPassword.SetFocus
       SendKeys "{Home}+{End}"
     End If
    End If
    rst.MoveNext
Loop
End Sub

