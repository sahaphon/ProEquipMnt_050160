Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmPMoldinj

  Dim da As New System.Data.OleDb.OleDbDataAdapter
  Dim ds As New DataSet
  Dim dsTn As New DataSet

Private Sub frmPMoldinj_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearTmpTable(0)  'ล้างข้อมูตาราง tmp_eqptrn_newsize
  ClearTmpTable(1)  'ล้างข้อมูตาราง pre_tmp_eqptrn_newsize
End Sub

Private Sub frmPMoldinj_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
 Dim strCode As String
     strCode = frmMoldInj.dgvShoe.Rows(frmMoldInj.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString

     ClearTmpTableUser("tmp_eqptrn")
     ClearTmpTable(0)                 'ล้างข้อมูตาราง tmp_eqptrn_newsize
     ClearTmpTable(1)  'ล้างข้อมูตาราง pre_tmp_eqptrn_newsize

     chkAll.Checked = False
     InputEqpDataPrint()
     InputSize(strCode)               'แสดง size ที่มีอยู่
     ShowSizeSelect()                 'เเสดง Size ที่เลือก
     cmbOptPrint.Text = frmMoldInj.dgvShoe.Rows(frmMoldInj.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim()

End Sub

Private Sub cmbOptPrint_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOptPrint.SelectedIndexChanged
  Dim strCode As String
      strCode = cmbOptPrint.Text.Trim
      ClearTmpTableUser("tmp_eqptrn")

      InputSize(strCode)  'แสดง size ที่มีอยู่
      ShowSizeSelect()  'เเสดง Size ที่เลือก
End Sub

Private Sub chkAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAll.CheckedChanged
 Dim txtEqpid As String

     If chkAll.Checked = True Then
        txtEqpid = cmbOptPrint.Text.Trim

        DeleteAllSize(txtEqpid) 'กรณีเลือกทุก size ให้ลบ size ที่เลือกไว้ก่อนหน้าทิ้ง
        InsertAllSize(txtEqpid)
     End If

End Sub

Private Sub InsertAllSize(ByVal strID As String)

  Dim conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlCmd As String

      With conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

          strSqlCmd = "INSERT INTO pre_tmp_eqptrn_newsize" _
                                  & " SELECT *" _
                                  & " FROM tmp_eqptrn_newsize " _
                                  & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                                  & " AND eqp_id = '" & strID & "'" _
                                  & " ORDER BY eqp_id "

          conn.Execute(strSqlCmd)
          ShowSizeSelect()            'แสดงข้อมูลใน gridview resultsize

  conn.Close()
  conn = Nothing

End Sub

Private Sub dgvSelectSize_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSelectSize.CellClick

 Dim strSqlSize As String = ""
 Dim strGpsize As String = ""
 Dim result As String = ""
 Dim strDocId As String = cmbOptPrint.Text.ToUpper.Trim

     With dgvSelectSize

           If dgvSelectSize.Rows.Count <> 0 Then

              Select Case .CurrentCell.ColumnIndex

                     Case Is = 2
                          chkAll.Checked = False 'ยกเลิการ check 
                          strSqlSize = "size_id = '" & dgvSelectSize.Rows(dgvSelectSize.CurrentRow.Index).Cells(0).Value.ToString & "'"
                          result = strSqlSize & " AND " & "size_desc = '" & dgvSelectSize.Rows(dgvSelectSize.CurrentRow.Index).Cells(1).Value.ToString & "'"

                          DeleteSize(result, strDocId)

               End Select

            End If

     End With

End Sub

Private Sub DeleteSize(ByVal strSize As String, ByVal strEqpid As String)
 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

     With Conn

         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

     End With

     With dgvSelectSize

        If .Rows.Count > 0 Then

                        Conn.BeginTrans()

                                '------------------------------------ลบตาราง eqpmst--------------------------------------------

                                strSqlCmd = "DELETE FROM pre_tmp_eqptrn_newsize" _
                                                      & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                                                      & " AND eqp_id = '" & strEqpid & "'" _
                                                      & " AND " & strSize

                                Conn.Execute(strSqlCmd)

                         Conn.CommitTrans()

                 .Rows.RemoveAt(.CurrentRow.Index)        'ลบเเถวปัจจุบัน
                End If

      .Focus()
    End With

Conn.Close()
Conn = Nothing

End Sub

Private Sub dgvSize_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSize.CellClick

 Dim strSqlSize As String = ""
 Dim strGpsize As String = ""
 Dim result As String = ""
 Dim strDocId As String = cmbOptPrint.Text.ToUpper.Trim      'เก็บ eqp_id

     With dgvSize

          Select Case .CurrentCell.ColumnIndex

                 Case Is = 0

                     If dgvSize.Rows.Count <> 0 Then
                        strSqlSize = "size_id = '" & dgvSize.Rows(e.RowIndex).Cells(1).Value.ToString & "'"
                        result = strSqlSize & " AND " & "size_desc = '" & dgvSize.Rows(e.RowIndex).Cells(2).Value.ToString & "'"

                        If ChkPrnSizExist(result) Then
                           prePrintSubreport(result, strDocId)     'คิวรี่ข้อมูลมาใส่ tmp_eqptrn ออกรายงานใน subreport

                        Else
                           MsgBox("ข้อมูลซ้ำ โปรดเลือก Size อื่่น")
                        End If

                     End If

          End Select

     End With

End Sub

Private Sub prePrintSubreport(ByVal strQuery As String, ByVal strEqpid As String)

 Dim conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

     With conn
          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

          strSqlCmd = "INSERT INTO pre_tmp_eqptrn_newsize " _
                           & " SELECT *" _
                           & " FROM tmp_eqptrn_newsize " _
                           & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                           & " AND eqp_id = '" & strEqpid & "'" _
                           & " AND " & strQuery _
                           & " ORDER BY tmp_newsize "

          conn.Execute(strSqlCmd)
          ShowSizeSelect()            'แสดงข้อมูลใน gridview resultsize

  conn.Close()
  conn = Nothing

End Sub

Private Function ChkPrnSizExist(ByVal txtSizeid As String) As Boolean

 Dim Conn As New ADODB.Connection
 Dim strSqlSelcSelc As String
 Dim Rsd As New ADODB.Recordset

     With Conn

         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

     End With

       strSqlSelcSelc = "SELECT * FROM pre_tmp_eqptrn_newsize (NOLOCK)" _
                                 & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                                 & " AND " & txtSizeid

       With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(strSqlSelcSelc, Conn, , , )

            If .RecordCount > 0 Then
                Return False
            Else
                Return True
            End If

       .ActiveConnection = Nothing
       .Close()

       End With

 Conn.Close()
 Conn = Nothing

End Function

Private Sub btnPrntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntCancel.Click
  ClearTmpTableUser("tmp_eqptrn")
  Me.Close()
End Sub

Private Sub DeleteAllSize(ByVal eqpid As String)
  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

     With Conn

         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

     End With

     With dgvSelectSize

        If .Rows.Count > 0 Then

                        Conn.BeginTrans()

                                '------------------------------------ลบตาราง pre_tmp_eqptrn_newsize--------------------------------------------

                                strSqlCmd = "DELETE FROM pre_tmp_eqptrn_newsize" _
                                                      & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                                      & " AND eqp_id = '" & eqpid & "'"


                                Conn.Execute(strSqlCmd)

                         Conn.CommitTrans()
                         ShowSizeSelect()  'แสดงข้อมูลใน gridview resultsize()

                 '.Rows.RemoveAt(.CurrentRow.Index)        'ลบเเถวปัจจุบัน
                End If

      .Focus()
    End With

Conn.Close()
Conn = Nothing
End Sub

Private Sub InputEqpDataPrint()

Dim Conn As New ADODB.Connection
Dim RsdPnt As New ADODB.Recordset

Dim strSqlSelc As String

        With Conn

             If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .Open()

        End With

        strSqlSelc = "SELECT eqp_id FROM v_moldinj_hd (NOLOCK)" _
                             & " WHERE ([group] ='A'" _
                             & " OR [group] ='B' OR [group] ='C' )" _
                             & " ORDER BY eqp_id"

       With RsdPnt

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlSelc, Conn, , , )

               If .RecordCount <> 0 Then

                        ds.Clear()
                        da.Fill(ds, RsdPnt, "eqpid")
                        cmbOptPrint.DataSource = ds.Tables("eqpid").DefaultView
                        cmbOptPrint.DisplayMember = "eqp_id"
                        cmbOptPrint.ValueMember = "eqp_id"

                End If

               .ActiveConnection = Nothing
             ' .Close()
     End With

     RsdPnt = Nothing

End Sub

Private Sub InputSize(ByVal code As String)     'ให้ระบุ size ก่อนพิมพ์

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim SqlSelc As String
 Dim sqlCmd As String

  Dim strArr() As String
  Dim SearchWithinThis As String
  Dim newSize As String

  Dim prDate As String
  Dim RecvDate As String
  Dim FcDate As String

  Dim completed As Boolean
  Dim sngMouthMold As Single

  Dim dbRecordProg As Double = 1

     With Conn
          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

        SqlSelc = "SELECT * FROM eqptrn (NOLOCK)" _
                                & "WHERE eqp_id ='" & code & "'" _
                                & "ORDER BY size_id "

       With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(SqlSelc, Conn, , , )

           If .RecordCount <> 0 Then

              '----------------------- ล้างข้อมูลใน tmp_eqptrn_newsize ----------------------

                 sqlCmd = "DELETE FROM tmp_eqptrn_newsize " _
                                  & "WHERE user_id= '" & frmMainPro.lblLogin.Text.ToString.Trim & "'"

                 Conn.Execute(sqlCmd)

              '-------------------- วนลูปใส่ค่าลงตาราง tmp_eqptrn_newsize --------------------

                 Do While Not .EOF

                      SearchWithinThis = .Fields("size_id").Value.ToString.Trim
                      If SearchWithinThis.IndexOf("-") <> -1 Then          'หาก size ต้นฉบับไม่มี size รว่ม (x-xx)
                         strArr = SearchWithinThis.Split("-")              'อ่านค่า size เก็บไว้ในตัวเเปร
                         newSize = strArr(0)
                      Else
                            newSize = .Fields("size_id").Value.ToString.Trim
                      End If

                      If .Fields("pr_date").Value.ToString.Trim <> "" Then
                         prDate = Mid(.Fields("pr_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("pr_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("pr_date").Value.ToString.Trim, 1, 2)
                         prDate = "'" & SaveChangeEngYear(prDate) & "'"

                     Else
                           prDate = "NULL"
                     End If

                     If .Fields("recv_date").Value.ToString.Trim <> "" Then
                         RecvDate = Mid(.Fields("recv_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("recv_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("recv_date").Value.ToString.Trim, 1, 2)
                         RecvDate = "'" & SaveChangeEngYear(RecvDate) & "'"
                     Else
                          RecvDate = "NULL"
                     End If

                     If .Fields("fc_date").Value.ToString.Trim <> "" Then
                         FcDate = Mid(.Fields("fc_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("fc_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("fc_date").Value.ToString.Trim, 1, 2)
                         FcDate = "'" & SaveChangeEngYear(FcDate) & "'"
                     Else
                          FcDate = "NULL"
                     End If

                     If .Fields("mouth_long").Value.ToString.Trim <> "" Then
                         sngMouthMold = ChangFormat(.Fields("mouth_long").Value.ToString.Trim)
                     Else
                          sngMouthMold = 0.0

                     End If

                     '----------------------- Insert ข้อมูลลงในตารางใหม่หลังเรียง size ใหม่ ----------------------

                   sqlCmd = "INSERT INTO tmp_eqptrn_newsize " _
                           & "(user_id,[group],eqp_id,size_id,size_desc,size_qty,weight,dimns,backgup " _
                           & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty " _
                           & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,tmp_newsize,mouth_long) " _
                           & "VALUES( " _
                           & "'" & frmMainPro.lblLogin.Text.Trim & "'" _
                           & ",'" & .Fields("group").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("eqp_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_desc").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("size_qty").Value.ToString.Trim) _
                           & "," & ChangFormat(.Fields("weight").Value.ToString.Trim) _
                           & ",'" & .Fields("dimns").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("backgup").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("price").Value.ToString.Trim) _
                           & ",'" & .Fields("men_rmk").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("delvr_sta").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("sent_sta").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("set_qty").Value.ToString.Trim) _
                           & "," & prDate _
                           & ",'" & .Fields("pr_doc").Value.ToString.Trim & "'" _
                           & "," & RecvDate _
                           & "," & ChangFormat(.Fields("ord_rep").Value.ToString.Trim) _
                           & "," & ChangFormat(.Fields("ord_qty").Value.ToString.Trim) _
                           & "," & FcDate _
                           & ",'" & .Fields("impt_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("sup_name").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("lp_type").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_group").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("cut_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("mate_type").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("cut_detail").Value.ToString.Trim & "'" _
                           & "," & newSize _
                           & "," & sngMouthMold _
                           & ")"

                        Conn.Execute(sqlCmd)

                     .MoveNext()   'เป็นออบเจ็กต์สำหรับ การเลื่อน Record ไป 1 Record

                             Application.DoEvents()
                 Loop
                 completed = True

            End If

       .ActiveConnection = Nothing
       .Close()

       End With

          If completed Then
             DisplaySize_to_select()
          End If

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub DisplaySize_to_select()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim sqlSelc As String

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

       sqlSelc = "SELECT * FROM tmp_eqptrn_newsize (NOLOCK)" _
                         & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                         & " ORDER BY size_desc, tmp_newsize"

       With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(sqlSelc, Conn, , , )

           If .RecordCount <> 0 Then

              dgvSize.Rows.Clear()

                 Do While Not .EOF

                    dgvSize.Rows.Add( _
                                        "เลือก", _
                                        .Fields("size_id").Value.ToString.Trim, _
                                        .Fields("size_desc").Value.ToString.Trim _
                                    )

                     .MoveNext()   'เป็นออบเจ็กต์สำหรับ การเลื่อน Record ไป 1 Record
                 Loop

           End If

         .ActiveConnection = Nothing
         .Close()

      End With

  Conn.Close()
End Sub

Private Sub ShowSizeSelect()     'แสดงข้อมูล Size ที่เลือกใน Data gridview

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlSelcSelc As String

      With Conn

           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()

      End With

       strSqlSelcSelc = "SELECT * FROM pre_tmp_eqptrn_newsize (NOLOCK)" _
                                   & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                   & " ORDER BY size_desc, tmp_newsize"

       With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(strSqlSelcSelc, Conn, , , )

           If .RecordCount <> 0 Then

               dgvSelectSize.Rows.Clear()

                 Do While Not .EOF

                    dgvSelectSize.Rows.Add( _
                                           .Fields("size_id").Value.ToString.Trim, _
                                           .Fields("size_desc").Value.ToString.Trim, _
                                           "ลบ" _
                                            )

                   .MoveNext()   'เป็นออบเจ็กต์สำหรับ การเลื่อน Record ไป 1 Record
                 Loop

            Else
                 dgvSelectSize.Rows.Clear()

            End If

       .ActiveConnection = Nothing
       .Close()

       End With

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub btnPrntPrevw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntPrevw.Click


 Dim strDocId As String = cmbOptPrint.Text.ToUpper.Trim
 Dim strQuery As String = ""

  If strDocId <> "" Then

      If dgvSelectSize.Rows.Count <> 0 Then

         PrePrintData(strDocId)
         frmMainPro.lblRptCentral.Text = "A"
         frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim _
                                                                & "' AND eqp_id ='" & strDocId & "'"

         frmRptCentral.ShowDialog()
      Else
         MsgBox("โปรดเลือก SIZE ก่อนพิมพ์รายงาน", MsgBoxStyle.Critical, "ผิดพลาด")
         Exit Sub
      End If

   Else
        MsgBox("โปรดระบุข้อมูลก่อนพิมพ์", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Equipment Empty!!")
        cmbOptPrint.Focus()
   End If
 Me.Close()
End Sub

Private Sub PrePrintData(ByVal strSelectCode As String)

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset
Dim RsdPic As New ADODB.Recordset

Dim strSqlSelc As String
Dim strSqlCmdPic As String
Dim strPicPath As String = "\\10.32.0.15\data1\EquipPicture\"

Dim strLoadFilePic1 As String
Dim strLoadFilePic2 As String
Dim strLoadFilePic3 As String

Dim blnHavePic1 As Boolean
Dim blnHavePic2 As Boolean
Dim blnHavePic3 As Boolean

Dim inImg As Image

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
    End With

     strSqlSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                         & " WHERE eqp_id = '" & strSelectCode.ToString.Trim & "'"


     With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then


                                    '----------------------------------------LoadPicture บรรจุอุปกรณ์------------------------------------------------

                                       strLoadFilePic1 = strPicPath & .Fields("pic_ctain").Value.ToString.Trim
                                       If strLoadFilePic1 <> "" Then

                                               If File.Exists(strLoadFilePic1) Then 'รูปยังมีอยู่ในระบบ
                                                      blnHavePic1 = True
                                               Else
                                                      blnHavePic1 = False
                                                End If

                                       Else
                                            blnHavePic1 = False
                                       End If

                                      '----------------------------------------LoadPicture ภายใน/ภายนอก------------------------------------------------

                                       strLoadFilePic2 = strPicPath & .Fields("pic_io").Value.ToString.Trim
                                       If strLoadFilePic2 <> "" Then

                                               If File.Exists(strLoadFilePic2) Then 'รูปยังมีอยู่ในระบบ
                                                      blnHavePic2 = True
                                               Else
                                                      blnHavePic2 = False
                                                End If

                                       Else
                                            blnHavePic2 = False
                                       End If


                                         '----------------------------------------LoadPicture ชิ้นงาน------------------------------------------------

                                       strLoadFilePic3 = strPicPath & .Fields("pic_part").Value.ToString.Trim
                                       If strLoadFilePic3 <> "" Then

                                               If File.Exists(strLoadFilePic3) Then 'รูปยังมีอยู่ในระบบ
                                                      blnHavePic3 = True
                                               Else
                                                      blnHavePic3 = False
                                                End If

                                       Else
                                            blnHavePic3 = False
                                       End If

                                    '-----------------------------------เพิ่มข้อมูล ไปที่ tmp_eqpmst-------------------------------

                                       strSqlCmdPic = "SELECT * " _
                                                                  & " FROM tmp_eqpmst (NOLOCK)"

                                            RsdPic = New ADODB.Recordset

                                            RsdPic.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                                            RsdPic.LockType = ADODB.LockTypeEnum.adLockOptimistic
                                            RsdPic.Open(strSqlCmdPic, Conn, , , )

                                                    RsdPic.AddNew()
                                                    RsdPic.Fields("user_id").Value = frmMainPro.lblLogin.Text.ToString.Trim
                                                    RsdPic.Fields("prod_sta").Value = .Fields("prod_sta").Value
                                                    RsdPic.Fields("fix_sta").Value = .Fields("fix_sta").Value
                                                    RsdPic.Fields("group").Value = .Fields("group").Value
                                                    RsdPic.Fields("eqp_id").Value = .Fields("eqp_id").Value
                                                    RsdPic.Fields("eqp_name").Value = .Fields("eqp_name").Value
                                                    RsdPic.Fields("pi").Value = .Fields("pi").Value
                                                    RsdPic.Fields("shoe").Value = .Fields("shoe").Value
                                                    RsdPic.Fields("part").Value = .Fields("part").Value
                                                    RsdPic.Fields("eqp_type").Value = .Fields("eqp_type").Value
                                                    RsdPic.Fields("ap_code").Value = .Fields("ap_code").Value
                                                    RsdPic.Fields("ap_code").Value = .Fields("ap_code").Value
                                                    RsdPic.Fields("ap_desc").Value = .Fields("ap_desc").Value
                                                    RsdPic.Fields("doc_ref").Value = .Fields("doc_ref").Value
                                                    RsdPic.Fields("set_qty").Value = .Fields("set_qty").Value
                                                    RsdPic.Fields("pic_ctain").Value = .Fields("pic_ctain").Value
                                                    RsdPic.Fields("pic_io").Value = .Fields("pic_io").Value
                                                    RsdPic.Fields("pic_part").Value = .Fields("pic_part").Value
                                                    RsdPic.Fields("remark").Value = .Fields("remark").Value
                                                    RsdPic.Fields("creat_date").Value = .Fields("creat_date").Value
                                                    RsdPic.Fields("pre_date").Value = .Fields("pre_date").Value
                                                    RsdPic.Fields("pre_by").Value = .Fields("pre_by").Value
                                                    RsdPic.Fields("last_date").Value = .Fields("last_date").Value
                                                    RsdPic.Fields("last_by").Value = .Fields("last_by").Value
                                                    RsdPic.Fields("pi_qty").Value = .Fields("pi_qty").Value
                                                    RsdPic.Fields("eqp_amt").Value = .Fields("eqp_amt").Value
                                                    RsdPic.Fields("exp_id").Value = .Fields("exp_id").Value

                                                    '---------------------------- เพิ่มข้อมูลรูปภาพบรรจุ ----------------------------------

                                                    If blnHavePic1 Then 'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                            Dim RsdSteam1 As New MemoryStream
                                                            Dim bytes1 = File.ReadAllBytes(strLoadFilePic1)

                                                            inImg = Image.FromFile(strLoadFilePic1)
                                                            inImg = ScaleImage(inImg, 227, 340)
                                                            inImg.Save(RsdSteam1, ImageFormat.Bmp)
                                                            bytes1 = RsdSteam1.ToArray
                                                            RsdPic.Fields("bob_ctain").Value = bytes1

                                                            RsdSteam1.Close()
                                                            RsdSteam1 = Nothing

                                                    End If

                                                     '---------------------------- เพิ่มข้อมูลรูปภาพภายนอก/ภายใน ---------------------------

                                                    If blnHavePic2 Then 'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                            Dim RsdSteam2 As New MemoryStream
                                                            Dim bytes2 = File.ReadAllBytes(strLoadFilePic2)

                                                            inImg = Image.FromFile(strLoadFilePic2)
                                                            inImg = ScaleImage(inImg, 227, 340)
                                                            inImg.Save(RsdSteam2, ImageFormat.Bmp)
                                                            bytes2 = RsdSteam2.ToArray
                                                            RsdPic.Fields("bob_io").Value = bytes2

                                                            RsdSteam2.Close()
                                                            RsdSteam2 = Nothing

                                                    End If

                                                    '---------------------------- เพิ่มข้อมูลรูปภาพชิ้นงาน ----------------------------------

                                                    If blnHavePic3 Then 'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                            Dim RsdSteam3 As New MemoryStream
                                                            Dim bytes3 = File.ReadAllBytes(strLoadFilePic3)

                                                            inImg = Image.FromFile(strLoadFilePic3)
                                                            inImg = ScaleImage(inImg, 227, 340)
                                                            inImg.Save(RsdSteam3, ImageFormat.Bmp)
                                                            bytes3 = RsdSteam3.ToArray
                                                            RsdPic.Fields("bob_part").Value = bytes3

                                                            RsdSteam3.Close()
                                                            RsdSteam3 = Nothing

                                                    End If

                                                    RsdPic.Update()


                                        RsdPic.ActiveConnection = Nothing
                                        RsdPic.Close()
                                        RsdPic = Nothing

             End If

            .ActiveConnection = Nothing
            .Close()

        End With

        Rsd = Nothing

Conn.Close()
Conn = Nothing

End Sub

Private Sub dgvSize_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvSize.RowsAdded
   dgvSize.Rows(e.RowIndex).Height = 24
End Sub

Private Sub dgvSelectSize_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvSelectSize.RowsAdded
   dgvSelectSize.Rows(e.RowIndex).Height = 24
End Sub

Private Sub ClearTmpTable(ByVal bytOption As Byte)

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String = ""

    With Conn

         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

            Select Case bytOption

                   Case Is = 0 'ลบข้อมูลหลังจากปิดฟอร์ม

                        strSqlCmd = "Delete FROM tmp_eqptrn_newsize" _
                                           & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"
                        .Execute(strSqlCmd)

                  Case Is = 1

                       strSqlCmd = "Delete FROM pre_tmp_eqptrn_newsize" _
                                         & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"
                      .Execute(strSqlCmd)

                 Case Is = 2

                       strSqlCmd = "Delete FROM tmp_eqptrn_newsize" _
                                         & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"
                      .Execute(strSqlCmd)


           End Select

       End With

   Conn.Close()
   Conn = Nothing

End Sub

End Class