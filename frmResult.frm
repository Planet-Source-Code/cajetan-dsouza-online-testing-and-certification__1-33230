VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmResult 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   -4425
   ClientTop       =   -1440
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar PBar2 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
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
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      OLEDropMode     =   1
      Scrolling       =   1
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
      TabIndex        =   17
      Top             =   1800
      Width           =   75
   End
   Begin VB.Shape Shape5 
      Height          =   495
      Left            =   840
      Top             =   1680
      Width           =   10335
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
      Left            =   6000
      TabIndex        =   16
      Top             =   1200
      Width           =   1710
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   5880
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   840
      Top             =   1080
      Width           =   4815
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
      TabIndex        =   15
      Top             =   7320
      Width           =   75
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
      TabIndex        =   14
      Top             =   6720
      Width           =   75
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
      Left            =   2760
      TabIndex        =   13
      Top             =   7320
      Width           =   1770
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
      Left            =   2760
      TabIndex        =   12
      Top             =   6720
      Width           =   4005
   End
   Begin VB.Line Line4 
      X1              =   2520
      X2              =   9240
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   9240
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   7560
      Y1              =   5280
      Y2              =   7680
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   9240
      Y1              =   5880
      Y2              =   5880
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
      TabIndex        =   11
      Top             =   6120
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
      TabIndex        =   10
      Top             =   5520
      Width           =   75
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
      Left            =   1080
      TabIndex        =   9
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The number of Wrong Answer is"
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
      Left            =   2760
      TabIndex        =   8
      Top             =   6060
      Width           =   75
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "The number of Right Answer is "
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
      Left            =   2760
      TabIndex        =   7
      Top             =   5460
      Width           =   75
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
      Left            =   3360
      TabIndex        =   6
      Top             =   3120
      Width           =   330
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
      Left            =   6720
      TabIndex        =   5
      Top             =   2400
      Width           =   450
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
      Left            =   4695
      TabIndex        =   4
      Top             =   360
      Width           =   2625
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
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   1650
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   840
      Top             =   2280
      Width           =   10335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000002&
      FillColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   2520
      Top             =   5280
      Width           =   6735
   End
End
Attribute VB_Name = "frmResult"
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

   com.ActiveConnection = conn
   rst.Open "Select * from StudentDetails where vRegistrationNo='" & frmDetail.txtRegNo & "'", conn, adOpenDynamic, adLockOptimistic
   
   PBar1.Value = 50
   PBar2.Value = frmReview.txtResult
   Label5.Caption = frmReview.txtResult + "%"
   Label5.Move (3240 + frmReview.txtResult * 68.4)
   Label8.Caption = "Registration No is " + frmDetail.txtRegNo
   Label9.Caption = frmReview.txtRight
   Label10.Caption = frmReview.txtWrong
   Label13.Caption = frmReview.txtResult / 100 * frmReview.txtTotal
   Label14.Caption = frmReview.txtTotal
   Label15.Caption = "Date appeared : " + CStr(frmDetail.txtDate)
   Label16.Caption = "Student Name : " + rst!vStudentName
   
   rst!iTotalMarks = CInt(frmReview.txtRight) + CInt(frmReview.txtWrong)
   rst!iMarksScored = frmReview.txtResult
   rst!iRightAnswer = frmReview.txtRight
   rst!iWrongAnswer = frmReview.txtWrong
   rst!dDateAppeared = frmDetail.txtDate
   rst.Update
   Unload frmOnlineTesting
   MsgBox "Press the 'P' key to PRINT the Report and 'Esc'Key to EXIT", vbInformation, Help
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
     Case 80
        On Error GoTo Handler
            PrintForm
Handler:          MsgBox "Printer Error : Printer does not exist or not connected.", vbCritical, "Error"
    Case 27
        End
   End Select
       
End Sub
