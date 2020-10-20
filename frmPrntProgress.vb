Imports ADODB

Public Class frmPrntProgress

Private Sub frmPrntProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  PrepareData()
End Sub

Private Sub PrepareData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim sql As String
  Dim sqlCmd As String

  Dim strArr() As String
  Dim SearchWithinThis As String
  Dim newSize As String

  Dim dbRecord As Double = 1

      Try

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      sql = "SELECT * FROM view_mold " _
                     & " WHERE ([group] ='A'" _
                     & " OR [group] ='B' OR [group] ='C' )" _
                     & " ORDER BY eqp_id"

      With Rsd

           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .CursorType = ADODB.CursorLocationEnum.adUseClient
           .Open(sql, Conn, , , )

           If .RecordCount <> 0 Then

              ProgressBar1.Minimum = 0
              ProgressBar1.Maximum = .RecordCount

               '----------------------- ล้างข้อมูลใน tmp_eqptrn_newsize ------------------------------

                 sqlCmd = "DELETE FROM print_view_allmold " _
                              & "WHERE user_id= '" & frmMainPro.lblLogin.Text.ToString.Trim & "'"

                 Conn.Execute(sqlCmd)

               ' ---------- วนลูปจัดเรียง size ใหม่ --------------

               Do While Not .EOF

                  SearchWithinThis = .Fields("size_id").Value.ToString.Trim
                  If SearchWithinThis.IndexOf("-") <> -1 Then          'หาก size ต้นฉบับไม่มี size รว่ม (x-xx)
                     strArr = SearchWithinThis.Split("-")              'อ่านค่า size เก็บไว้ในตัวเเปร
                     newSize = strArr(0)
                  Else
                       newSize = .Fields("size_id").Value.ToString.Trim
                  End If

                  '----------------------- Insert ข้อมูลลงในตารางใหม่หลังเรียง size ใหม่ ----------------------

                   sqlCmd = "INSERT INTO print_view_allmold " _
                           & "(user_id,eqp_id,eqp_name,desc_thai,size_id,size_desc " _
                           & ",set_qty,dimns,price,sup_name,[group],tmp_size)" _
                           & "VALUES( " _
                           & "'" & frmMainPro.lblLogin.Text.Trim & "'" _
                           & ",'" & .Fields("eqp_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("eqp_name").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("desc_thai").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_desc").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("set_qty").Value) _
                           & ",'" & .Fields("dimns").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("price").Value) _
                           & ",'" & .Fields("sup_name").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("group").Value.ToString.Trim & "'" _
                           & "," & newSize _
                           & ")"

                        Conn.Execute(sqlCmd)
                        ProgressBar1.Value = dbRecord

                  .MoveNext()

                            dbRecord = dbRecord + 1
                            Application.DoEvents()

               Loop


           End If

          .ActiveConnection = Nothing
          .Close()
      End With

   Conn.Close()

      Catch ex As Exception
            MsgBox(ex.Message)
      End Try

   Me.Close()

End Sub

End Class