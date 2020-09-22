VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGenerateTest 
   Caption         =   "Generate test"
   ClientHeight    =   8490
   ClientLeft      =   -4365
   ClientTop       =   -990
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtAns 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   -840
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   6720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
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
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   4920
      Width           =   2055
   End
   Begin VB.ComboBox cmbTest 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3840
      Width           =   4335
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox txtRegNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3840
      TabIndex        =   0
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   6120
      Width           =   3975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Test Name :"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Student Name :"
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
      Left            =   2205
      TabIndex        =   7
      Top             =   3060
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Registration Number :"
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
      Left            =   1530
      TabIndex        =   6
      Top             =   2280
      Width           =   2265
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   11280
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Generate Test"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4350
      TabIndex        =   5
      Top             =   600
      Width           =   3315
   End
End
Attribute VB_Name = "frmGenerateTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim com As Command
Dim rst As Recordset
Dim rst1 As Recordset
Dim rst2 As Recordset
Private Sub cmdCancel_Click()
  frmMain.SetFocus
  Unload Me
End Sub

Private Sub Form_Load()
   frmDetail.Show
   frmDetail.Visible = False
   Set conn = New Connection
   Set com = New Command
   Set rst = New Recordset
   Set rst1 = New Recordset
   Set rst2 = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   com.ActiveConnection = conn
   rst.Open "StudentDetails", conn, adOpenDynamic, adLockOptimistic
   
   
   ''''Populating the Combo box with Test names'''''''''''
   
   rst1.Open "TestDetails", conn, adOpenDynamic, adLockOptimistic
   
   rst1.MoveFirst
   Do While Not rst1.EOF
     cmbTest.AddItem (rst1!vTestName)
     rst1.MoveNext
   Loop
   cmbTest.ListIndex = 0
   
   '''''
End Sub

Private Sub cmdGenerate_Click()
Dim flag As Boolean
Dim obj As clsGenerateTest
Set obj = New clsGenerateTest

flag = False
 
 If txtRegNo.Text = "" Then
   txtRegNo.SetFocus
   MsgBox "Enter the Student Registration Number.", vbInformation, "Invalid Data"
   SendKeys "{Home}+{End}"
 ElseIf txtName.Text = "" Then
   txtName.SetFocus
   MsgBox "Enter the Student Name.", vbInformation, "Invalid Data"
   SendKeys "{Home}+{End}"
 ElseIf cmbTest.Text = "" Then
   cmbTest.SetFocus
   MsgBox "Select the Test Name from the list.", vbInformation, "Invalid Data"
   SendKeys "{Home}+{End}"
 Else
    If Not rst.RecordCount < 0 Then
      rst.MoveFirst
    End If
    
    Do While Not rst.EOF
      If txtRegNo.Text = rst!vRegistrationNo And cmbTest.Text = rst!vTestName Then
        flag = True
        Exit Do
      End If
      rst.MoveNext
    Loop
     If flag = False Then
           rst.AddNew
              rst("vRegistrationNo") = txtRegNo.Text
              rst("vStudentName") = txtName.Text
              rst("vTestName") = cmbTest.Text
           rst.Update
           obj.Generate (Trim(cmbTest.Text))
           If txtAns = True Then
               'frmOnlineTesting.Show
               frmDetail.txtRegNo = txtRegNo
               For i = 0 To 999
                 PB.Value = PB.Value + 1
               Next
               Label5.Caption = "Test Generated....."
               frmSampleTest.Show
               
           Else
               txtRegNo.SetFocus
           End If
     Else
       cmbTest.SetFocus
       MsgBox "Registration Number already exist.Student has appeared the test before.", vbInformation, "Invalid Data"
       SendKeys "{Home}+{End}"
    End If
 End If
End Sub
