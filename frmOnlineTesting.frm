VERSION 5.00
Begin VB.Form frmOnlineTesting 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   8595
   ClientLeft      =   -4365
   ClientTop       =   -1380
   ClientWidth     =   11880
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmOnlineTesting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9000
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtIndex 
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   22
      Top             =   6580
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   21
      Top             =   6020
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   20
      Top             =   7200
      Width           =   255
   End
   Begin VB.OptionButton optAnswer 
      Appearance      =   0  'Flat
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   19
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   18
      Top             =   1080
      Width           =   10455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Question"
      Height          =   615
      Left            =   6193
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtQuestionNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtTotalQuestion 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   255
         Width           =   180
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Marks"
      Height          =   615
      Left            =   3033
      TabIndex        =   5
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtPositive 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtNegative 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Positive : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   250
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Negative :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.CheckBox chkBookmark 
      Appearance      =   0  'Flat
      Caption         =   "Bookmark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   953
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
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
      Left            =   8213
      TabIndex        =   3
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdCalculator 
      Caption         =   "Calculator"
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
      Left            =   6173
      TabIndex        =   2
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton cmdReview 
      Caption         =   "Review"
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
      Left            =   4013
      TabIndex        =   1
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
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
      Left            =   1853
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7320
      Top             =   -2400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time Remaining"
      Height          =   615
      Left            =   8153
      TabIndex        =   14
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtMin 
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
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtSec 
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
         Height          =   240
         Left            =   1560
         TabIndex        =   28
         Text            =   "0"
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Minutes  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   1373
      TabIndex        =   17
      Top             =   7920
      Width           =   9135
   End
   Begin VB.Label txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   27
      Top             =   6000
      Width           =   7815
   End
   Begin VB.Label txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   26
      Top             =   6600
      Width           =   7815
   End
   Begin VB.Label txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   25
      Top             =   7200
      Width           =   7815
   End
   Begin VB.Label txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   24
      Top             =   5280
      Width           =   7815
   End
   Begin VB.Line Line3 
      X1              =   606.061
      X2              =   11272.73
      Y1              =   879.581
      Y2              =   879.581
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11151.51
      Y1              =   -1382.199
      Y2              =   -1382.199
   End
   Begin VB.Line Line1 
      X1              =   606.061
      X2              =   11272.73
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmOnlineTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim QPaper As Recordset
Dim TestDetails As Recordset
Dim QStatus As Recordset
Dim Marks As Integer
Dim Caller As String
Dim View As String

Private Sub Form_Load()
   Set conn = New Connection
   Set QPaper = New Recordset
   Set QStatus = New Recordset
   Set TestDetails = New Recordset

   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   QPaper.Open "Select iQuestionNo,vQuestion,vOption1,vOption2,vOption3,vOption4,vAnswer,iPositive,iNegative from QuestionPaper order by iQuestionNo", conn, adOpenDynamic, adLockOptimistic
   QStatus.Open "select iQuestionNo,vStatus,vAnswer from QuestionStatus order by iQuestionNo", conn, adOpenDynamic, adLockOptimistic
   TestDetails.Open "select * from TestDetails where vTestname='" & frmGenerateTest.cmbTest & "'", conn, adOpenDynamic, adLockOptimistic
   

   QPaper.MoveLast
   txtTotalQuestion.Text = QPaper!iquestionno
   txtQuestionNo.Text = 1

   QPaper.MoveFirst
   QStatus.MoveFirst

   Call Fill
   txtMin = TestDetails!iTimeinMinutes
   If QStatus!vStatus = "B" Then
      chkBookmark.Value = 1
      ind = 0

      QStatus.MoveFirst
      QStatus.MoveFirst
      ind = QStatus!vAnswer

      If Not ind = "" Then
          optAnswer(ind).Value = 1
          txtAnswer(ind).ForeColor = &HFF&
      End If
   ElseIf QStatus!vStatus = "S" Then
      ind = QStatus!vAnswer
      optAnswer(ind) = True
      txtAnswer(ind).ForeColor = &HFF&
   End If
   
End Sub

Private Sub chkBookmark_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  QStatus.MoveFirst
  QStatus.Move (txtQuestionNo)
  If QStatus.EOF Then
    QStatus.MoveLast
  End If

  QStatus.MovePrevious
  If Not QStatus!vAnswer = "" Then
    txtIndex.Text = "S"
  Else
    txtIndex.Text = "NS"
  End If
  Caller = "EnterData"
  Call EnterData
  Caller = ""
End Sub

Private Sub cmdNext_Click()
    chkBookmark.Value = 0

    For i = 1 To 4
       optAnswer(i).Value = False
       txtAnswer(i).ForeColor = &H0&
    Next
    QPaper.MoveFirst
    QStatus.MoveFirst

    If Not QPaper.EOF Then
        QPaper.Move (txtQuestionNo.Text)
        QStatus.Move (txtQuestionNo.Text)
        Call Fill

        If QStatus!vStatus = "B" Then
          chkBookmark.Value = 1
          ind = 0

          QStatus.MoveFirst
          QStatus.Move (txtQuestionNo.Text - 1)
          ind = QStatus!vAnswer

          If Not ind = "" Then
            optAnswer(ind).Value = 1
            txtAnswer(ind).ForeColor = &HFF&
          End If
        ElseIf QStatus!vStatus = "S" Then
          ind = QStatus!vAnswer
          optAnswer(ind) = True
          txtAnswer(ind).ForeColor = &HFF&
        End If
     Else
         MsgBox "You are on the Last Question.Next", vbInformation, "Message"
     End If
End Sub

Private Sub cmdPrevious_Click()
    txtQuestionNo.Text = txtQuestionNo.Text - 1
    chkBookmark.Value = 0
    QPaper.MoveFirst
    QStatus.MoveFirst

    For i = 1 To 4
       optAnswer(i).Value = False
       txtAnswer(i).ForeColor = &H0&
    Next
    If Not QPaper.BOF Then
        txtQuestionNo.Text = txtQuestionNo.Text - 1
        QPaper.Move (txtQuestionNo.Text)
        QStatus.Move (txtQuestionNo.Text)
        Call Fill

        If QStatus!vStatus = "B" Then
              chkBookmark.Value = 1
              ind = 0
              QStatus.MoveFirst

              QStatus.Move (txtQuestionNo.Text - 1)
              ind = QStatus!vAnswer
              If Not ind = "" Then
                optAnswer(QStatus!vAnswer).Value = 1
                txtAnswer(QStatus!vAnswer).ForeColor = &HFF&
              End If
        ElseIf QStatus!vStatus = "S" Then
            ind = QStatus!vAnswer
            optAnswer(ind) = True
            txtAnswer(ind).ForeColor = &HFF&
        End If
    Else
        MsgBox "You are on the First Question.Next", vbInformation, "Message"
    End If
End Sub

Private Sub optAnswer_LostFocus(Index As Integer)
   txtIndex.Text = Index
   QStatus.MoveFirst

   Do While Not QStatus.EOF
        If QStatus!iquestionno = txtQuestionNo.Text Then
          QStatus!vAnswer = Index
          QStatus.Update
        End If
        QStatus.MoveNext
   Loop
End Sub

Private Sub optAnswer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   For i = 1 To 4                  'Changing 4 radio button to black
     txtAnswer(i).ForeColor = Black
   Next
   txtAnswer(Index).ForeColor = &HFF&  'Changing color to red of selected radio button
   QStatus.MoveFirst
   Do While Not QStatus.EOF
        If QStatus!iquestionno = txtQuestionNo.Text Then
          QStatus!vAnswer = Index
          QStatus.Update
        End If
        QStatus.MoveNext
   Loop
   txtIndex.Text = "S"
   Caller = "EnterData"
   Call EnterData
   Caller = ""
End Sub

Private Sub txtMin_Change()
  If txtMin.Text < "0" And txtSec.Text = "0" Then
        txtMin.Text = "0"
        MsgBox "Test is ended", vbInformation, "Message"
        frmOnlineTesting.Enabled = False
        frmOnlineTesting.Visible = False
        Timer1.Enabled = False
        frmReview.Timer1.Enabled = False
        For Each Control In frmReview
          If TypeOf Control Is CommandButton Then
            Control.Enabled = False
          End If
        Next
        frmReview.cmdEndTest.Enabled = True
        frmReview.cmdView.Enabled = False
        frmReview.Visible = True
  End If
End Sub

Private Sub Timer1_Timer()
 If txtSec.Text = 0 Then
   txtMin.Text = txtMin.Text - 1
   txtSec.Text = 59
 Else
   txtSec.Text = txtSec.Text - 1
 End If
End Sub

Private Sub cmdCalculator_Click()
      Call Shell("C:\Windows\Calc.exe", vbNormalFocus)
End Sub

Private Sub cmdReview_Click()
  frmReview.Visible = True
  frmOnlineTesting.Visible = False
End Sub

Function Fill()
    If QPaper.EOF Then
       MsgBox "You are on the Last Question..", vbInformation, "Message"
       QPaper.MoveLast
       QStatus.MoveLast
    ElseIf QPaper.BOF Then
       txtQuestionNo.Text = 1
       MsgBox "You are on the first Question", vbInformation, "Message"
       QPaper.MoveFirst
       QStatus.MoveFirst
    Else                               'Fill up the questions
      'QPaper.MoveFirst
      'QPaper.Move (txtQuestionNo.Text)
      txtQuestion.Text = QPaper!vQuestion
      txtAnswer(1).Caption = QPaper!vOption1
      txtAnswer(2).Caption = QPaper!vOption2
      txtAnswer(3).Caption = QPaper!vOption3
      txtAnswer(4).Caption = QPaper!vOption4
      txtPositive.Text = QPaper!iPositive
      txtNegative.Text = QPaper!iNegative
      txtQuestionNo.Text = QPaper!iquestionno
    End If
End Function

Function EnterData()
   QStatus.MoveFirst

   For i = 1 To 4
     If optAnswer(i) = 1 Then
       txtIndex.Text = "S"
       Exit For
     End If
   Next

   Do While Not QStatus.EOF
     If QStatus!iquestionno = txtQuestionNo.Text Then
        If Caller = "EnterData" Then
           If chkBookmark.Value = 1 Then     'If Answer is not selected
             For i = 1 To 4
               If optAnswer(i) = 1 Then
                 QStatus!vAnswer = i
                 Exit For
               End If
             Next
              QStatus!vStatus = "B"
              QStatus.Update
           ElseIf txtIndex.Text = "S" Then
              QStatus!vStatus = "S"
              QStatus.Update
           Else
              QStatus!vStatus = "NS"
              QStatus.Update
           End If
        End If
     End If

     If Not QStatus.EOF Then
       QStatus.MoveNext
     Else
        Exit Do
     End If
   Loop
End Function
