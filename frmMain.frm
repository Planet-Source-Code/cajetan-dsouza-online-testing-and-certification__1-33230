VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0E0FF&
   Caption         =   "OLT and C"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7590
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test"
      Begin VB.Menu mnuNew 
         Caption         =   "New Test..."
      End
      Begin VB.Menu mnuAlter 
         Caption         =   "Alter Test..."
      End
      Begin VB.Menu mnuGenerate 
         Caption         =   "Generate Test"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Test..."
      End
      Begin VB.Menu hy1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Change Background Color..."
      End
      Begin VB.Menu hy2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResult 
         Caption         =   "Print Student Result"
      End
      Begin VB.Menu mnuDeleteStud 
         Caption         =   "Delete Student Details..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu mnuCreate 
         Caption         =   "Create User..."
      End
      Begin VB.Menu mnuChange 
         Caption         =   "Change Password..."
      End
      Begin VB.Menu mnuHy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteUser 
         Caption         =   "Delete User..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents..."
      End
      Begin VB.Menu hy3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As Connection
Dim rst As Recordset
Dim rst1 As Recordset

Private Sub MDIForm_Load()
   Set conn = New Connection
   Set rst = New Recordset
   Set rst1 = New Recordset
   
   With conn
     .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   rst.Open "FormOption", conn, adOpenDynamic, adLockOptimistic

'   '''''''/////////////////////////////////
  rst1.Open "Login", conn, adOpenDynamic, adLockOptimistic

  Do While Not rst1.EOF
    If rst1("CurrentLogin") = True Then
        MsgBox "Inside"
        If rst1!GenTest = False Then
           mnuGenerate.Enabled = False
        End If
        If rst1!CreateTest = False Then
           mnuNew.Enabled = False
        End If
        If rst1!ModifyTest = False Then
           mnuAlter.Enabled = False
        End If
        If rst1!DeleteTest = False Then
           mnuDelete.Enabled = False
        End If
        Exit Do
   End If
   rst1.MoveNext
 Loop
   frmMain.BackColor = rst!BackColor
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  rst1.Close
'  rst1.Open "Select * from Login where CurrentLogin = True", conn, adOpenDynamic, adLockOptimistic
  rst1.Open "Login", conn, adOpenDynamic, adLockOptimistic
  Do While Not rst1.EOF
    If rst1!CurrentLogin = True Then
       rst1!CurrentLogin = False
       rst1.Update
    End If
    rst1.MoveNext
  Loop
End Sub

Private Sub mnuAbout_Click()
   frmAbout.Show
End Sub

Private Sub mnuAlter_Click()
  frmAlterTest.Show
End Sub

Private Sub mnuChange_Click()
  frmChangePassword.Show
End Sub

Private Sub mnuColor_Click()
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   CommonDialog1.Flags = cdlCCRGBInit
   CommonDialog1.ShowColor

   frmMain.BackColor = CommonDialog1.Color
   rst!BackColor = CommonDialog1.Color
   rst.Update
   Exit Sub
ErrHandler:
   Exit Sub
End Sub

Private Sub mnuCreate_Click()
  frmCreateUser.Show
End Sub

Private Sub mnuDelete_Click()
  frmDeleteTest.Show
End Sub

Private Sub mnuDeleteStud_Click()
  frmDeleteStudent.Show
End Sub

Private Sub mnuExit_Click()
Dim Ans As Integer
Ans = MsgBox("Do you want to exit Online Testing and Certification.", vbYesNo, "Confirmation")
If Ans = 6 Then
  Unload Me
End If
End Sub

Private Sub mnuGenerate_Click()
  frmGenerateTest.Show
End Sub

Private Sub mnuNew_Click()
  frmNewTest1.Show
End Sub

Private Sub mnuResult_Click()
  frmPrintResult.Show
End Sub

