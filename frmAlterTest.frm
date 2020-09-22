VERSION 5.00
Begin VB.Form frmAlterTest 
   Caption         =   "Alter Test"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   120
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      DataField       =   "vQuestion"
      DataSource      =   "Adodc1"
      Height          =   2775
      Left            =   1650
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1560
      Width           =   6495
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   9000
      TabIndex        =   17
      Top             =   2280
      Width           =   2415
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
      Left            =   8970
      TabIndex        =   16
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      DataField       =   "vOption3"
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
      Left            =   1650
      TabIndex        =   4
      Top             =   5520
      Width           =   6495
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      DataField       =   "vOption2"
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
      Left            =   1650
      TabIndex        =   3
      Top             =   5040
      Width           =   6495
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      DataField       =   "vOption1"
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
      Left            =   1650
      TabIndex        =   2
      Top             =   4560
      Width           =   6495
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1170
      TabIndex        =   8
      Top             =   5640
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1170
      TabIndex        =   7
      Top             =   5160
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1170
      TabIndex        =   6
      Top             =   4680
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   1170
      TabIndex        =   9
      Top             =   6120
      Width           =   255
   End
   Begin VB.TextBox txtAnswer 
      Appearance      =   0  'Flat
      DataField       =   "vOption4"
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
      Left            =   1650
      TabIndex        =   5
      Top             =   6000
      Width           =   6495
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   9090
      TabIndex        =   19
      Top             =   360
      Width           =   2415
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   1500
      TabIndex        =   18
      Top             =   6720
      Width           =   8415
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Previous"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtNegativeMarks 
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
      Left            =   10440
      TabIndex        =   11
      Top             =   5040
      WhatsThisHelpID =   1
      Width           =   615
   End
   Begin VB.TextBox txtPositiveMarks 
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
      Left            =   10440
      TabIndex        =   10
      Top             =   4560
      WhatsThisHelpID =   1
      Width           =   615
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
      TabIndex        =   27
      Top             =   4560
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
      Left            =   330
      TabIndex        =   26
      Top             =   1560
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
      Left            =   210
      TabIndex        =   25
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Alter Test"
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
      Left            =   5295
      TabIndex        =   24
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   855
      X2              =   11535
      Y1              =   720
      Y2              =   720
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
      Index           =   0
      Left            =   16905
      TabIndex        =   23
      Top             =   15360
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
      Left            =   16800
      TabIndex        =   22
      Top             =   15840
      Width           =   1770
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
      Index           =   1
      Left            =   8745
      TabIndex        =   21
      Top             =   4680
      Width           =   1665
   End
   Begin VB.Label Label7 
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
      Left            =   8640
      TabIndex        =   20
      Top             =   5160
      Width           =   1770
   End
End
Attribute VB_Name = "frmAlterTest"
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
   txtDate = CStr(Date) + "      " + CStr(Time)
End Sub
Private Sub cmbTestName_Click()
  rst.Close
  rst.Open "Select * from Question where vTestName='" & cmbTestName.Text & "'", conn, adOpenDynamic, adLockOptimistic
  If Not rst.EOF Then
    rst.MoveFirst
    Call Fill
  End If
  cmdNext.Enabled = True
  cmdFirst.Enabled = True
  cmdPrevious.Enabled = True
  cmdLast.Enabled = True
End Sub

Private Sub cmdCancel_Click()
  frmMain.SetFocus
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  Call SaveInfo
  rst.MoveFirst
  Call Fill
End Sub

Private Sub cmdLast_Click()
  Call SaveInfo
  rst.MoveLast
  Call Fill
End Sub

Private Sub cmdNext_Click()
  Call SaveInfo
  If rst.EOF Then
    rst.MoveLast
  Else
    rst.MoveNext
    If rst.EOF Then
      MsgBox "End of the Test", vbInformation, "Information"
    End If
  End If
  Call Fill
End Sub

Private Sub cmdOK_Click()
  Call SaveInfo
  frmMain.SetFocus
  Unload Me
End Sub

Private Sub cmdPrevious_Click()
  Call SaveInfo
  If rst.BOF Then
    rst.MoveFirst
    MsgBox "Beginning of then test", vbInformation, "Information"
  Else
    rst.MovePrevious
    If rst.BOF Then
      MsgBox "Beginning of then test", , "Information"
    End If
  End If
  Call Fill
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

Function Fill()
    If rst.EOF Then
      rst.MoveLast
    ElseIf rst.BOF Then
      rst.MoveFirst
    End If
      
      txtQuestion = rst!vQuestion
      txtAnswer(1) = rst!vOption1
      txtAnswer(2) = rst!vOption2
      txtAnswer(3) = rst!vOption3
      txtAnswer(4) = rst!vOption4
      txtPositiveMarks = rst!iPositiveMarks
      txtNegativeMarks = rst!iNegativeMarks
      
      If rst!vAnswer = rst!vOption1 Then
        optAnswer(1).Value = True
      ElseIf rst!vAnswer = rst!vOption2 Then
        optAnswer(2).Value = True
      ElseIf rst!vAnswer = rst!vOption3 Then
        optAnswer(3).Value = True
      ElseIf rst!vAnswer = rst!vOption4 Then
        optAnswer(4).Value = True
      End If
End Function


Function SaveInfo()
  rst!vQuestion = txtQuestion.Text
  rst!vOption1 = txtAnswer(1).Text
  rst!vOption2 = txtAnswer(2).Text
  rst!vOption3 = txtAnswer(3).Text
  rst!vOption4 = txtAnswer(4).Text
  rst!iPositiveMarks = txtPositiveMarks.Text
  
  
  If optAnswer(1).Value = True Then
    rst!vAnswer = txtAnswer(1).Text
  ElseIf optAnswer(2).Value = True Then
    rst!vAnswer = txtAnswer(2).Text
  ElseIf optAnswer(3).Value = True Then
    rst!vAnswer = txtAnswer(3).Text
  ElseIf optAnswer(4).Value = True Then
    rst!vAnswer = txtAnswer(4).Text
  End If
  rst!iNegativeMarks = txtNegativeMarks.Text
  rst.Update
End Function

Private Sub Timer1_Timer()
  txtDate = CStr(Date) + "      " + CStr(Time)
End Sub
