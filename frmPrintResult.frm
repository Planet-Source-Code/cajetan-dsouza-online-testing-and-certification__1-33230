VERSION 5.00
Begin VB.Form frmPrintResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Result"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbTest 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3563
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   1523
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtRegNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Test name :"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1500
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Registraton No.:"
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
      Left            =   420
      TabIndex        =   1
      Top             =   720
      Width           =   1845
   End
End
Attribute VB_Name = "frmPrintResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim rst As Recordset
Dim rst1 As Recordset
Dim rst2 As Recordset
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Set conn = New Connection
   Set rst = New Recordset
   Set rst1 = New Recordset
   Set rst2 = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   rst.Open "StudentDetails", conn, adOpenDynamic, adLockOptimistic
   
   
   ''''Populating the Combo box with Test names'''''''''''
   
   rst1.Open "TestDetails", conn, adOpenDynamic, adLockOptimistic
   
   rst1.MoveFirst
   Do While Not rst1.EOF
     cmbTest.AddItem (rst1!vTestName)
     rst1.MoveNext
   Loop
   cmbTest.ListIndex = 0
End Sub

Private Sub cmdOK_Click()
    Dim flag As Boolean
    flag = False
    rst.MoveFirst
    Do While Not rst.EOF
      If UCase(txtRegNo.Text) = UCase(rst!vRegistrationNo) And cmbTest.Text = rst!vTestName Then
        flag = True
        Exit Do
      End If
      rst.MoveNext
    Loop
 If flag = True Then
    frmPrint.Show
    Unload Me
 Else
    MsgBox "The registration number does not exist.", vbExclamation, "Error"
 End If
End Sub
