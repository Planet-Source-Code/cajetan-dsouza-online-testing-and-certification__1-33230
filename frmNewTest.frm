VERSION 5.00
Begin VB.Form frmNewTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Test"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   9240
      TabIndex        =   24
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtTotalPositive 
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
      Left            =   10800
      Locked          =   -1  'True
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Text            =   "0"
      ToolTipText     =   "Total number of Positive Marks present for this Test"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtTotalNegative 
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
      Left            =   10800
      TabIndex        =   20
      Text            =   "0"
      ToolTipText     =   "Total number of Positive Marks present for this Test"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtPositive 
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
      Left            =   10800
      TabIndex        =   17
      Top             =   5520
      WhatsThisHelpID =   1
      Width           =   615
   End
   Begin VB.TextBox txtNegative 
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
      Left            =   10800
      TabIndex        =   16
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtIndex 
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Left            =   9240
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtAnswer 
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
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   6480
      Width           =   6495
   End
   Begin VB.TextBox txtAnswer 
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
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   6000
      Width           =   6495
   End
   Begin VB.TextBox txtAnswer 
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
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   5520
      Width           =   6495
   End
   Begin VB.OptionButton optAnswer 
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   8
      Top             =   6600
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   7
      Top             =   6120
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   5640
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   9
      Top             =   7080
      Width           =   255
   End
   Begin VB.TextBox txtAnswer 
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
      Index           =   4
      Left            =   1560
      TabIndex        =   5
      Top             =   6960
      Width           =   6495
   End
   Begin VB.TextBox txtQuestion 
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
      Height          =   4065
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   960
      Width           =   6495
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
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
      Height          =   255
      Left            =   9480
      TabIndex        =   12
      Top             =   -240
      Width           =   1695
   End
   Begin VB.TextBox txtTestName 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Total Positive Marks :"
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
      Left            =   8505
      TabIndex        =   23
      Top             =   3240
      Width           =   2265
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Negative Marks :"
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
      Left            =   8400
      TabIndex        =   22
      Top             =   3720
      Width           =   2370
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Positive Marks :"
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
      Left            =   9105
      TabIndex        =   19
      Top             =   5640
      Width           =   1665
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Negative Marks :"
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
      Left            =   9000
      TabIndex        =   18
      Top             =   6120
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Answers :"
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
      TabIndex        =   14
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Question :"
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
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1095
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
      TabIndex        =   11
      Top             =   480
      Width           =   1275
   End
End
Attribute VB_Name = "frmNewTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim rst As Recordset

Private Sub Form_Load()
   txtDate.Text = Date
   txtTestName.Text = frmNewTest1.txtTestName.Text
   Unload frmNewTest1
   Set conn = New Connection
   Set rst = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   rst.Open "Question", conn, adOpenDynamic, adLockOptimistic
   rst.MoveFirst
   Do While Not rst.EOF
     If rst!vTestName = txtTestName Then
       txtTotalPositive = CInt(txtTotalPositive.Text) + rst!iPositiveMarks
     End If
     rst.MoveNext
   Loop
   txtIndex.Text = 1
End Sub

Private Sub cmdAdd_Click()
  Dim Stat As Boolean
    
  For Each Control In frmNewTest
    If TypeOf Control Is TextBox Then
      If Control.Text = "" Then
        MsgBox "Incomplete Data. Enter the data correctly", vbExclamation, "Incomplete Data"
        txtQuestion.SetFocus
        Stat = False
        Exit For
      Else
        Stat = True
      End If
    End If
  Next

  If Not IsNumeric(txtPositive) Or Not IsNumeric(txtNegative) Then
     MsgBox "Invalid data . Enter Numeric value", vbInformation, "Invalid Data"
     txtPositive.SetFocus
     Exit Sub
  End If
  
  If txtPositive = 0 Then
    MsgBox "Positive marks cannot be Zero. Reenter positive marks", vbInformation, "Invalid Data"
    txtPositive.SetFocus
    Stat = False
  End If
  
  If Stat = True Then
   rst.MoveLast
   No = rst!iquestionno
   
    rst.AddNew
      rst("vTestName") = txtTestName.Text
      rst("iQuestionNo") = No + 1
      rst("vQuestion") = txtQuestion.Text
      rst("vOption1") = txtAnswer(1).Text
      rst("vOption2") = txtAnswer(2).Text
      rst("vOption3") = txtAnswer(3).Text
      rst("vOption4") = txtAnswer(4).Text
      rst("vAnswer") = txtAnswer(txtIndex.Text).Text
      rst("iPositiveMarks") = txtPositive.Text
      rst("iNegativeMarks") = txtNegative.Text
    rst.Update
  txtTotalPositive = CInt(txtTotalPositive) + CInt(txtPositive)
  txtTotalNegative = CInt(txtTotalNegative) + CInt(txtNegative)
  
  txtQuestion = ""
  For i = 1 To 4
    txtAnswer(i) = ""
  Next
  optAnswer(1).Value = True
  txtPositive = ""
  txtNegative.Text = 0
  End If
End Sub

Private Sub cmdCancel_Click()
  frmMain.SetFocus
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

Private Sub optAnswer_GotFocus(Index As Integer)
 txtIndex.Text = Index
End Sub

