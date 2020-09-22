VERSION 5.00
Begin VB.Form frmDeleteTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Test"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   1935
   ClientWidth     =   7350
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2441.739
   ScaleMode       =   0  'User
   ScaleWidth      =   7350
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
      Left            =   3960
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Delete"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cmbTestName 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmDeleteTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim com As Command
Dim rst As Recordset
Dim rst1 As Recordset

Private Sub Form_Load()
   Set conn = New Connection
   Set com = New Command
   Set rst = New Recordset
   Set rst1 = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   rst.Open "question", conn, adOpenDynamic, adLockOptimistic
   rst1.Open "Testdetails", conn, adOpenDynamic, adLockOptimistic
   rst1.MoveFirst
   Do While Not rst1.EOF
     cmbTestName.AddItem (rst1!vTestName)
     rst1.MoveNext
   Loop
   frmDeleteTest.Height = 2895
   frmDeleteTest.Width = 7410
End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  a = MsgBox("Do you want to Delete this Test", vbYesNoCancel, "Confirmation")
  If a = 6 Then
    conn.Execute "Delete  from TestDetails where vTestName='" & cmbTestName.Text & "'"
    conn.Execute "Delete  from Question where vTestName='" & cmbTestName.Text & "'"
  End If
  cmbTestName.RemoveItem (cmbTestName.ListIndex)
End Sub
