VERSION 5.00
Begin VB.Form frmSampleReview 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEndTest 
      Caption         =   "End Test"
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
      Left            =   9360
      TabIndex        =   144
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time Remaining"
      Height          =   615
      Left            =   8400
      TabIndex        =   136
      Top             =   480
      Width           =   2775
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
         TabIndex        =   138
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
         TabIndex        =   137
         Top             =   255
         Width           =   750
      End
   End
   Begin VB.TextBox txtSec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Left            =   9960
      TabIndex        =   135
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txtMin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Left            =   8640
      TabIndex        =   134
      Top             =   720
      Width           =   255
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   3960
      Top             =   0
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
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
      Left            =   5400
      TabIndex        =   133
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "11"
      Height          =   375
      Index           =   10
      Left            =   1791
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "10"
      Height          =   375
      Index           =   9
      Left            =   1795
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   375
      Index           =   8
      Left            =   1791
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   1791
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   1795
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   1791
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   1795
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   1795
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   1795
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   1791
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "12"
      Height          =   375
      Index           =   11
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "13"
      Height          =   375
      Index           =   12
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "14"
      Height          =   375
      Index           =   13
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "15"
      Height          =   375
      Index           =   14
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "16"
      Height          =   375
      Index           =   15
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "17"
      Height          =   375
      Index           =   16
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "18"
      Height          =   375
      Index           =   17
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "19"
      Height          =   375
      Index           =   18
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "20"
      Height          =   375
      Index           =   19
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "21"
      Height          =   375
      Index           =   20
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "22"
      Height          =   375
      Index           =   21
      Left            =   2515
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "23"
      Height          =   375
      Index           =   22
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "24"
      Height          =   375
      Index           =   23
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "25"
      Height          =   375
      Index           =   24
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "26"
      Height          =   375
      Index           =   25
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "27"
      Height          =   375
      Index           =   26
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "28"
      Height          =   375
      Index           =   27
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "29"
      Height          =   375
      Index           =   28
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "30"
      Height          =   375
      Index           =   29
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "31"
      Height          =   375
      Index           =   30
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "32"
      Height          =   375
      Index           =   31
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "33"
      Height          =   375
      Index           =   32
      Left            =   3235
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "34"
      Height          =   375
      Index           =   33
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "35"
      Height          =   375
      Index           =   34
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "36"
      Height          =   375
      Index           =   35
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "37"
      Height          =   375
      Index           =   36
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "38"
      Height          =   375
      Index           =   37
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "39"
      Height          =   375
      Index           =   38
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "40"
      Height          =   375
      Index           =   39
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "41"
      Height          =   375
      Index           =   40
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "42"
      Height          =   375
      Index           =   41
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "43"
      Height          =   375
      Index           =   42
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "44"
      Height          =   375
      Index           =   43
      Left            =   3955
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "45"
      Height          =   375
      Index           =   44
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "46"
      Height          =   375
      Index           =   45
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "47"
      Height          =   375
      Index           =   46
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "48"
      Height          =   375
      Index           =   47
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "49"
      Height          =   375
      Index           =   48
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "50"
      Height          =   375
      Index           =   49
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "51"
      Height          =   375
      Index           =   50
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "52"
      Height          =   375
      Index           =   51
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "53"
      Height          =   375
      Index           =   52
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "54"
      Height          =   375
      Index           =   53
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "55"
      Height          =   375
      Index           =   54
      Left            =   4675
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "56"
      Height          =   375
      Index           =   55
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "57"
      Height          =   375
      Index           =   56
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "58"
      Height          =   375
      Index           =   57
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "59"
      Height          =   375
      Index           =   58
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "60"
      Height          =   375
      Index           =   59
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "61"
      Height          =   375
      Index           =   60
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "62"
      Height          =   375
      Index           =   61
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "63"
      Height          =   375
      Index           =   62
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "64"
      Height          =   375
      Index           =   63
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "65"
      Height          =   375
      Index           =   64
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "66"
      Height          =   375
      Index           =   65
      Left            =   5395
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "67"
      Height          =   375
      Index           =   66
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "68"
      Height          =   375
      Index           =   67
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "69"
      Height          =   375
      Index           =   68
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "70"
      Height          =   375
      Index           =   69
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "71"
      Height          =   375
      Index           =   70
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "72"
      Height          =   375
      Index           =   71
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "73"
      Height          =   375
      Index           =   72
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "74"
      Height          =   375
      Index           =   73
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "75"
      Height          =   375
      Index           =   74
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "76"
      Height          =   375
      Index           =   75
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "77"
      Height          =   375
      Index           =   76
      Left            =   6115
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "78"
      Height          =   375
      Index           =   77
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "79"
      Height          =   375
      Index           =   78
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "80"
      Height          =   375
      Index           =   79
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "81"
      Height          =   375
      Index           =   80
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "82"
      Height          =   375
      Index           =   81
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "83"
      Height          =   375
      Index           =   82
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "84"
      Height          =   375
      Index           =   83
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "85"
      Height          =   375
      Index           =   84
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "86"
      Height          =   375
      Index           =   85
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "87"
      Height          =   375
      Index           =   86
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "88"
      Height          =   375
      Index           =   87
      Left            =   6835
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "89"
      Height          =   375
      Index           =   88
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "90"
      Height          =   375
      Index           =   89
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "91"
      Height          =   375
      Index           =   90
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "92"
      Height          =   375
      Index           =   91
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "93"
      Height          =   375
      Index           =   92
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "94"
      Height          =   375
      Index           =   93
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "95"
      Height          =   375
      Index           =   94
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "96"
      Height          =   375
      Index           =   95
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "97"
      Height          =   375
      Index           =   96
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "98"
      Height          =   375
      Index           =   97
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "99"
      Height          =   375
      Index           =   98
      Left            =   7555
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "100"
      Height          =   375
      Index           =   99
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "101"
      Height          =   375
      Index           =   100
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "102"
      Height          =   375
      Index           =   101
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "103"
      Height          =   375
      Index           =   102
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "104"
      Height          =   375
      Index           =   103
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "105"
      Height          =   375
      Index           =   104
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "106"
      Height          =   375
      Index           =   105
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "107"
      Height          =   375
      Index           =   106
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "108"
      Height          =   375
      Index           =   107
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "109"
      Height          =   375
      Index           =   108
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "110"
      Height          =   375
      Index           =   109
      Left            =   8275
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "111"
      Height          =   375
      Index           =   110
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "112"
      Height          =   375
      Index           =   111
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "113"
      Height          =   375
      Index           =   112
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "114"
      Height          =   375
      Index           =   113
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "115"
      Height          =   375
      Index           =   114
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "116"
      Height          =   375
      Index           =   115
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "117"
      Height          =   375
      Index           =   116
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "118"
      Height          =   375
      Index           =   117
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "119"
      Height          =   375
      Index           =   118
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "120"
      Height          =   375
      Index           =   119
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "121"
      Height          =   375
      Index           =   120
      Left            =   8995
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "122"
      Height          =   375
      Index           =   121
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "123"
      Height          =   375
      Index           =   122
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "124"
      Height          =   375
      Index           =   123
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "125"
      Height          =   375
      Index           =   124
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "126"
      Height          =   375
      Index           =   125
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "127"
      Height          =   375
      Index           =   126
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "128"
      Height          =   375
      Index           =   127
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "129"
      Height          =   375
      Index           =   128
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "130"
      Height          =   375
      Index           =   129
      Left            =   9715
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   1791
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtResult 
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtRight 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtWrong 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   833
      X2              =   11153
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line 
      X1              =   0
      X2              =   10320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   1080
      X2              =   11160
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Question Review"
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
      Left            =   4313
      TabIndex        =   143
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "** Click on the question number button to jump to that respective question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2880
      TabIndex        =   142
      Top             =   7320
      Width           =   6315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "   Bookmark    "
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
      Left            =   1800
      TabIndex        =   141
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "   Not Solved    "
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
      Left            =   4417
      TabIndex        =   140
      Top             =   7800
      Width           =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      Caption         =   "   Solved    "
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
      Left            =   7080
      TabIndex        =   139
      Top             =   7800
      Width           =   1020
   End
End
Attribute VB_Name = "frmSampleReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim No As Integer
   Dim conn As Connection
   Dim QPaper As Recordset
   Dim QStatus As Recordset
   Dim com As Command
   Dim Com1 As Command
   Dim Sum As Integer


Private Sub cmdEndTest_Click()
a = MsgBox("Do you want to end the Sample Test", vbYesNoCancel, "Confirmation")
If a = 6 Then
  Unload frmSampleReview
  Unload frmSampleOLT
  frmDetail.txtDate = Now
  frmOnlineTesting.Show
End If
End Sub

Private Sub Form_Load()
   Set conn = New Connection
   Set QStatus = New Recordset
   Set QPaper = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With
   
   
   QStatus.Open "SampleStatus", conn, adOpenDynamic, adLockOptimistic
   QPaper.Open "SamplePaper", conn, adOpenDynamic, adLockOptimistic
    
   If Not QStatus.RecordCount < 0 Then
        QStatus.MoveFirst
   End If
   
   Do While Not QStatus.EOF
     Command1(QStatus!iquestionno - 1).Visible = True
     If QStatus!vStatus = "S" Then
         Command1(QStatus!iquestionno - 1).BackColor = &HC0C000   'Red
     ElseIf QStatus!vStatus = "B" Then
         Command1(QStatus!iquestionno - 1).BackColor = &HC0FFC0   'Green
     ElseIf QStatus!vStatus = "NS" Then
         Command1(QStatus!iquestionno - 1).BackColor = &HC0C0FF         'Dark green
     End If
     QStatus.MoveNext
   Loop
   
   txtMin.Text = frmSampleOLT.Text1.Text
   txtSec.Text = frmSampleOLT.Text2.Text
End Sub
   
Private Sub cmdView_Click()
  frmSampleReview.Visible = False
  frmSampleOLT.Visible = True
End Sub

Private Sub Command1_Click(Index As Integer)
   frmSampleOLT.Visible = True
   For i = 1 To 4
     frmSampleOLT.txtanswer1(i).ForeColor = &H0&
     frmSampleOLT.optAnswer(i).Value = False
   Next
   
   frmSampleOLT.chkBookmark.Value = 0
   frmSampleOLT.txtQuestionNo = Command1(Index).Caption
   QPaper.MoveFirst
   QStatus.MoveFirst
   QPaper.Move (Command1(Index).Caption - 1)
   QStatus.Move (Command1(Index).Caption - 1)
   
   Call Fill
   If QStatus!vStatus = "B" Then
      frmSampleOLT.chkBookmark.Value = 1
      ind = 0
    '  QStatus.MoveFirst
    '  QStatus.Move (Command1(Index).Caption - 1)
      
      ind = QStatus!vAnswer
      If Not ind = "" Then
        frmSampleOLT.optAnswer(QStatus!vAnswer).Value = 1
        frmSampleOLT.txtanswer1(QStatus!vAnswer).ForeColor = &HFF&
      End If
   ElseIf QStatus!vStatus = "S" Then
     QStatus.MoveFirst
     QStatus.Move (Command1(Index).Caption - 1)

     ind = QStatus!vAnswer
     frmSampleOLT.optAnswer(ind) = True
     frmSampleOLT.txtanswer1(ind).ForeColor = &HFF&
   End If
    '''''''''''''''''''''''''''''''''''''''''''''''''
   Unload Me
End Sub

Private Sub Timer_Timer()
  txtMin.Text = frmSampleOLT.Text1.Text
  txtSec.Text = frmSampleOLT.Text2.Text
End Sub
Function Fill()
   frmSampleOLT.txtQuestion.Text = QPaper!vQuestion
   frmSampleOLT.txtanswer1(1).Caption = QPaper!vOption1
   frmSampleOLT.txtanswer1(2).Caption = QPaper!vOption2
   frmSampleOLT.txtanswer1(3).Caption = QPaper!vOption3
   frmSampleOLT.txtanswer1(4).Caption = QPaper!vOption4
   frmSampleOLT.txtPositive.Text = QPaper!iPositive
   frmSampleOLT.txtNegative.Text = QPaper!iNegative
   frmSampleOLT.txtQuestionNo.Text = QPaper!iquestionno
End Function

Function Calculate()
  Dim Ans As String
  Dim AnsNo As Integer
  Dim cnt As Integer
  Dim cnt1 As Integer
  Dim Tot As Integer
    
  QPaper.MoveFirst
  QStatus.MoveFirst
  Do While Not QPaper.EOF And Not QStatus.EOF
      cnt1 = cnt1 + 1
      AnsNo = 0
      Tot = Tot + QPaper!iPositive
      If QStatus!vStatus = "S" Then
        AnsNo = QStatus!vAnswer
        Select Case AnsNo
          Case "1"
            Ans = QPaper!vOption1
          Case "2"
            Ans = QPaper!vOption2
          Case "3"
            Ans = QPaper!vOption3
          Case "4"
            Ans = QPaper!vOption4
        End Select
        If Ans = QPaper!vAnswer Then
           cnt = cnt + 1
           Sum = Sum + QPaper!iPositive
        Else
           Sum = Sum - QPaper!iNegative
        End If
     End If
     QStatus.MoveNext
     QPaper.MoveNext
  Loop
  txtResult.Text = Round(Sum * 100 / Tot)
  txtRight = cnt
  txtWrong = cnt1 - cnt
  frmResult.Show
End Function


