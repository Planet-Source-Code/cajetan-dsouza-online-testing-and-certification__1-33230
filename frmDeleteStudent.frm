VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDeleteStudent 
   Caption         =   "Delete Student Details"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10560
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1940
      Left            =   720
      TabIndex        =   8
      Top             =   3720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   4560
      TabIndex        =   7
      Top             =   6765
      Width           =   2055
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
      Left            =   3901
      TabIndex        =   6
      Top             =   2925
      Width           =   2895
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
      Left            =   3901
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2205
      Width           =   2895
   End
   Begin VB.OptionButton optDelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete Individual Student details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   1500
      Width           =   3495
   End
   Begin VB.OptionButton optDelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete all student for particular Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Value           =   -1  'True
      Width           =   3495
   End
   Begin VB.Label Label3 
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
      Left            =   2460
      TabIndex        =   4
      Top             =   2205
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Registration No :"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   1755
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   0
      X2              =   9308
      Y1              =   1358
      Y2              =   1358
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Delete Student(s) Details"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2010
      TabIndex        =   0
      Top             =   645
      Width           =   5250
   End
End
Attribute VB_Name = "frmDeleteStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim com As Command
Dim rst As Recordset
Dim rst1 As Recordset
Dim rst2 As Recordset
Dim ItemIndex%

Private Sub cmbTestName_Click()
    ListView1.ListItems.Clear
    
    SQL = "Select * from StudentDetails where vTestName= '" & cmbTestName & "'"
    Set rst = conn.Execute(SQL)
    
    If Not rst.EOF Then
        Dim a As Integer
        a = 1
        Do Until rst.EOF
            ListView1.ListItems.Add , , rst!vRegistrationNo
            ListView1.ListItems(a).ListSubItems.Add , , rst!vStudentName
            ListView1.ListItems(a).ListSubItems.Add , , rst!vTestName
            ListView1.ListItems(a).ListSubItems.Add , , rst!dDateAppeared
            ListView1.ListItems(a).ListSubItems.Add , , rst!iMarksScored
            
            a = a + 1
            rst.MoveNext
        Loop
'        ListView1.SetFocus
    End If
    
    rst.Close
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

   rst.Open "StudentDetails", conn, adOpenDynamic, adLockOptimistic
   ''''Populating the Combo box with Test names'''''''''''
   
   rst1.Open "TestDetails", conn, adOpenDynamic, adLockOptimistic
   
   rst1.MoveFirst
   Do While Not rst1.EOF
     cmbTestName.AddItem (rst1!vTestName)
     rst1.MoveNext
   Loop
   cmbTestName.ListIndex = 0
   '''''
   
   Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Reg No."
    Header.Width = ListView1.Width * 0.17   ' 17%

   Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Name"
    Header.Width = (ListView1.Width * 0.32)  ' 32%
    
   Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Test name"
    Header.Width = ListView1.Width * 0.17    ' 17%
    
   Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Date appeared"
    Header.Width = ListView1.Width * 0.17    ' 17%
    
   Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Marks obtained"
    Header.Width = ListView1.Width * 0.16    ' 16%

End Sub
Private Sub cmdCancel_Click()
  frmMain.Enabled = True
  Unload Me
  frmMain.SetFocus
End Sub

Private Sub optDelete_Click(Index As Integer)
   ListView1.ListItems.Clear
   If Index = 2 Then
      txtRegNo.Enabled = False
      Label2.Enabled = False
      cmbTestName_Click
   Else
      txtRegNo.Enabled = True
      Label2.Enabled = True
   End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ItemIndex = Item.Index
End Sub

Private Sub cmdDelete1_Click()
  If ListView1.ListItems.Count = 0 Then Exit Sub
   If ItemIndex <> 0 Then
        Dim Ask As String
        Ask = MsgBox("Are you sure that you want to delete", vbYesNo + vbInformation, "Delete record")
           
        If Ask = vbYes Then
          If optDelete(1).Value = True Then
              conn.Execute "Delete * from StudentDetails where vRegistrationNo='" & txtRegNo.Text & "'"
              ListView1.ListItems.Remove (ItemIndex)
           Else
              conn.Execute "Delete * from StudentDetails where vTestName ='" & cmbTestName & "'"
              cmbTestName_Click
              ListView1.ListItems.Clear
           End If
        End If
      End If
     ItemIndex = 0
End Sub

Private Sub txtRegNo_Change()
  ListView1.ListItems.Clear
  If Not txtRegNo = "" Then
     SQL = "Select * from StudentDetails where vRegistrationNo ='" & txtRegNo & "'"
     Set rst = conn.Execute(SQL)
     If Not rst.EOF Then
        ListView1.ListItems.Add , , rst!vRegistrationNo
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , rst!vStudentName
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , rst!vTestName
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , rst!dDateAppeared
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , rst!iMarksScored
     End If
  End If
End Sub
