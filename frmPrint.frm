VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPrint 
   BackColor       =   &H80000009&
   Caption         =   "Result"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   11320
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOK 
      Caption         =   "Close"
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
      Left            =   6105
      TabIndex        =   1
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   4065
      TabIndex        =   0
      Top             =   6960
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar PBar2 
      Height          =   375
      Left            =   3353
      TabIndex        =   2
      Top             =   3233
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   375
      Left            =   3353
      TabIndex        =   3
      Top             =   2513
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin VB.Shape Shape2 
      Height          =   2415
      Left            =   2400
      Top             =   4320
      Width           =   6855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The total negative marks obtained by candidate"
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
      Index           =   2
      Left            =   2520
      TabIndex        =   19
      Top             =   5160
      Width           =   4965
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The total positive marks obtained by candidate"
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
      Left            =   2520
      TabIndex        =   18
      Top             =   4440
      Width           =   4890
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   833
      Top             =   2153
      Width           =   10335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Marks Required"
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
      Left            =   1193
      TabIndex        =   17
      Top             =   2633
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Marks Obtained"
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
      Left            =   1193
      TabIndex        =   16
      Top             =   3353
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Result Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4688
      TabIndex        =   15
      Top             =   233
      Width           =   2625
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "50%"
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
      Left            =   6713
      TabIndex        =   14
      Top             =   2273
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "0%"
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
      Left            =   3353
      TabIndex        =   13
      Top             =   2993
      Width           =   330
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
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
      Left            =   1073
      TabIndex        =   12
      Top             =   1073
      Width           =   75
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
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
      Left            =   7800
      TabIndex        =   11
      Top             =   4560
      Width           =   75
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
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
      Left            =   7800
      TabIndex        =   10
      Top             =   5160
      Width           =   75
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   9240
      Y1              =   6560
      Y2              =   6560
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   7560
      Y1              =   5760
      Y2              =   8960
   End
   Begin VB.Line Line3 
      X1              =   2400
      X2              =   9240
      Y1              =   7360
      Y2              =   7360
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   9240
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The total marks obtained by candidate"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   5760
      Width           =   4005
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Total Test marks"
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
      Top             =   6360
      Width           =   1770
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
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
      Left            =   7800
      TabIndex        =   7
      Top             =   5760
      Width           =   75
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
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
      Left            =   7800
      TabIndex        =   6
      Top             =   6360
      Width           =   75
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   833
      Top             =   953
      Width           =   4815
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   5873
      Top             =   953
      Width           =   5295
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Date appeared :"
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
      Left            =   5993
      TabIndex        =   5
      Top             =   1073
      Width           =   1710
   End
   Begin VB.Shape Shape5 
      Height          =   495
      Left            =   833
      Top             =   1553
      Width           =   10335
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
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
      Left            =   1080
      TabIndex        =   4
      Top             =   1680
      Width           =   75
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim com As Command
Dim rst As Recordset

Private Sub cmdOK_Click()
  Unload Me
  frmMain.Show
End Sub

Private Sub cmdPrint_Click()
    For Each Control In frmPrint
     If TypeOf Control Is Command Then
       Control.Visible = False
     End If
    Next
    On Error GoTo Handler
    PrintForm
    
Handler:
    MsgBox "Printer Error : Printer does not exist or not connected.", vbCritical, "Error"
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
   rst.Open "Select * from StudentDetails where vRegistrationNo='" & frmPrintResult.txtRegNo & "'", conn, adOpenDynamic, adLockOptimistic
   
   PBar1.Value = 50
   PBar2.Value = rst!iMarksScored
   Label5.Caption = CStr(rst!iMarksScored) + "%"
   Label5.Move (3240 + rst!iMarksScored * 68.4)
   Label8.Caption = "Registration No is " + rst!vRegistrationNo
   Label9.Caption = rst!iRightAnswer
   Label10.Caption = rst!iWrongAnswer
   Label13.Caption = Round(rst!iMarksScored / 100 * rst!iTotalMarks)
   Label14.Caption = rst!iTotalMarks
   Label15.Caption = "Date appeared : " + CStr(rst!dDateAppeared)
   Label16.Caption = "Student Name : " + rst!vStudentName
End Sub


