VERSION 5.00
Begin VB.Form frmNewTest1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Test"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   ControlBox      =   0   'False
   Icon            =   "frmNewTest1.frx":0000
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTime 
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
      Left            =   4200
      TabIndex        =   2
      Top             =   1455
      Width           =   1815
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
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
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
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtTotalMarks 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtTestName 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Time (Min)"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Marks :"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1365
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1275
   End
End
Attribute VB_Name = "frmNewTest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim rst As Recordset

Private Sub Form_Load()
   Set conn = New Connection
   Set rst = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   rst.Open "TestDetails", conn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub cmdCancel_Click()
  frmMain.SetFocus
  Unload Me
End Sub

Private Sub cmdOK_Click()
 If IsNumeric(txtTotalMarks) And IsNumeric(txtTime) Then
    If Not txtTotalMarks = "" And Not txtTime = "" Then
      If Not CInt(txtTotalMarks) < 1 Then
         rst.AddNew
          rst!vTestName = txtTestName
          rst!iTotalMarks = txtTotalMarks
          rst!iTimeinMinutes = txtTime
         rst.Update
         frmNewTest.Show
       Else
         MsgBox "Total marks are zero. Enter marks greater than zero", vbInformation, "Invalid Data"
         txtTotalMarks.SetFocus
       End If
     Else
       MsgBox "Total marks are zero. Enter marks greater than zero", vbInformation, "Invalid Data"
       txtTotalMarks.SetFocus
     End If
  Else
     MsgBox "Invalid Value", vbInformation, "Invalid Data"
  End If
End Sub
