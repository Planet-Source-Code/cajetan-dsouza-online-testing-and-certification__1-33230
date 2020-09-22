VERSION 5.00
Begin VB.Form frmSampleTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terms & Rules"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSampleTest 
      Caption         =   "Sample Test"
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
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Take a sample test for practicing."
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdStartTest 
      Caption         =   "Start Test"
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
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Start test without taking sample test."
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmSampleTest.frx":0000
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rules :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmSampleTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSampleTest_Click()
  frmMain.Visible = False
  frmSampleOLT.Show
  Unload Me
End Sub

Private Sub cmdStartTest_Click()
  frmMain.Visible = False
  frmDetail.txtDate = Now
  frmOnlineTesting.Show
  Unload Me
End Sub
