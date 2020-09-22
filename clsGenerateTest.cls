VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenerateTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim rs3 As Recordset
Dim rCount As Integer
Public a As Boolean

Function Generate(Test As String)
   Dim i As Integer
   Dim flag As Boolean
   Dim RemMarks As Integer
   RemMarks = 0
   rCount = 0

   Set conn = New Connection
   Set com = New Command
   Set rs = New Recordset
   Set rs1 = New Recordset
   Set rs2 = New Recordset
   Set rs3 = New Recordset

   With conn
       .Provider = "Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Information.mdb"
     .Open
   End With

   rs.Open "Select * from Question where vTestName='" & Test & "'", conn, adOpenDynamic, adLockOptimistic
   rs1.Open "Select * from TestDetails where vTestName='" & Test & "'", conn, adOpenDynamic, adLockOptimistic
   rs2.Open "select * from QuestionPaper", conn, adOpenDynamic, adLockOptimistic
   rs3.Open "QuestionStatus", conn, adOpenDynamic, adLockOptimistic


   If Not rs.EOF Then                   'Counting the no. of records
        Do While Not rs.EOF
            rCount = rCount + rs!iPositiveMarks
            rs.MoveNext
        Loop
   Else
       rCount = 0
   End If
   If rCount >= rs1!iTotalMarks Then          'If there are no records
            rs.MoveFirst
            Do While Not rs.EOF         'Finding highest marks
               If HighestMarks < rs!iPositiveMarks Then
                     HighestMarks = rs!iPositiveMarks
               End If
               rs.MoveNext
            Loop

            TotalMarks = rs1!iTotalMarks
            RemMarks = TotalMarks
            div = TotalMarks / HighestMarks   'Dividing marks

            For i = 1 To HighestMarks
                 flag = False
                 rs.Close
                 rs.Open "Select * from Question where vTestName='" & Test & "'", conn, adOpenDynamic, adLockOptimistic
                 rs.MoveFirst

                 Do While Not rs.EOF
                     If i = rs!iPositiveMarks Then   'If that marks question exist
                        flag = True
                        Exit Do
                      End If
                     rs.MoveNext
                 Loop

                 If flag = True Then
                     flag = False
                     NoOfQues = div / i               'Dividing ques for each mark
                     rs.Close
                     rs.Open "select * from Question where vTestName='" & Test & "' and iPositiveMarks= '" & i & "' ", conn, adOpenDynamic, adLockOptimistic
                     If Not rs2.EOF Then          'Finding the next Ques no.
                         rs2.MoveLast
                         cnt = rs2!iquestionno + 1
                     Else
                         cnt = 1                  'If no records the ques no=1
                     End If

                     rCount = 0
                     rs.MoveFirst
                     Do While Not rs.EOF
                       rCount = rCount + 1
                       rs.MoveNext
                     Loop

                     rs.MoveFirst
                     If rCount < Round(NoOfQues) Then
                         NoOfQues = rCount
                     End If

                     For j = 1 To Round(NoOfQues)
                         rs2.AddNew
                             rs2!iquestionno = cnt
                             rs2!vQuestion = rs!vQuestion
                             rs2!vOption1 = rs!vOption1
                             rs2!vOption2 = rs!vOption2
                             rs2!vOption3 = rs!vOption3
                             rs2!vOption4 = rs!vOption4
                             rs2!vAnswer = rs!vAnswer
                             rs2!iPositive = rs!iPositiveMarks
                             rs2!iNegative = rs!iNegativeMarks
                         rs2.Update
                         rs3.AddNew
                             rs3!iquestionno = cnt
                             rs3!vStatus = "NS"
                         rs3.Update
                         cnt = cnt + 1
                         rs.MoveNext
                         RemMarks = RemMarks - i
                     Next j
                 End If
             Next
            rs.Close
            rs.Open "Select * from Question where vTestName='" & Test & "'", conn, adOpenDynamic, adLockOptimistic
            rs.MoveLast
            Do While Not rs.BOF
                    If RemMarks = rs!iPositiveMarks Then   'If that marks question exist
                       rs2.AddNew
                             rs2!iquestionno = cnt
                             rs2!vQuestion = rs!vQuestion
                             rs2!vOption1 = rs!vOption1
                             rs2!vOption2 = rs!vOption2
                             rs2!vOption3 = rs!vOption3
                             rs2!vOption4 = rs!vOption4
                             rs2!vAnswer = rs!vAnswer
                             rs2!iPositive = rs!iPositiveMarks
                             rs2!iNegative = rs!iNegativeMarks
                         rs2.Update
                         rs3.AddNew
                             rs3!iquestionno = cnt
                             rs3!vStatus = "NS"
                         rs3.Update
                        Exit Do
                      End If
                     rs.MovePrevious
             Loop
             frmGenerateTest.txtAns = True
         Else
            MsgBox "There is no question paper for this test or no sufficent questions.", vbInformation, "Generation Error"
            frmGenerateTest.txtAns = False
         End If
End Function


