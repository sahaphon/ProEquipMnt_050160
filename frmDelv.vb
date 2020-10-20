Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D

Public Class frmDelv

 Dim intBkPageCount As Integer
 Dim blnHaveFilter As Boolean

 Dim dubNumberStart As Double
 Dim dubNumberEnd As Double

 Dim strSqlFindData As String

 Dim strDocCode As String = "F5"

 Dim da As New System.Data.OleDb.OleDbDataAdapter
 Dim ds As New DataSet
 Dim dsTn As New DataSet

Private Sub frmDelv_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

     InputDeptData()

     If FormCount("frmAeDelv") > 0 Then

         With frmAeDelv

               strSearch = .lblComplete.Text

               If strSearch <> "" Then
                  SearchData(0, strSearch)
               End If

              .Close()

         End With

      End If

    Timer1.Enabled = True       'ให้ Timer1 รีเฟรชหน้าจอ

    Me.Height = Int(lblHeight.Text)
    Me.Width = Int(lblWidth.Text)

    Me.Top = Int(lblTop.Text)
    Me.Left = Int(lblLeft.Text)

End Sub

Private Function FormCount(ByVal frmName As String) As Long
Dim frm As Form

    For Each frm In My.Application.OpenForms

                If frm Is My.Forms.frmAeDelv Then
                        FormCount = FormCount + 1
                End If
    Next

End Function

Public Sub Activating()
   Dim strSearch As String

       strSearch = lblCmd.Text.Trim

       If strSearch <> "" Then
          SearchData(0, strSearch)
       End If

    Timer1.Enabled = False
End Sub

Private Sub frmDelv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Me.WindowState = FormWindowState.Maximized
    StdDateTimeThai()
    tlsBarFmr.Cursor = Cursors.Hand

   dubNumberStart = 1
   dubNumberEnd = 2100

   PreGroupType()

   InputDeptData()
   tabCmd.Focus()

End Sub

Private Sub frmDelv_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

  lblHeight.Text = Me.Height.ToString.Trim
  lblWidth.Text = Me.Width.ToString.Trim

  lblTop.Text = Me.Top.ToString.Trim
  lblLeft.Text = Me.Left.ToString.Trim

End Sub

Private Sub frmDelv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
End Sub

Private Sub PreGroupType()

 Dim strGpTopic(4) As String
 Dim i As Byte

      strGpTopic(0) = "เลขที่เอกสาร"
      strGpTopic(1) = "ผู้โอน"
      strGpTopic(2) = "ผู้รับโอน"
      strGpTopic(3) = "แผนกรับมอบ"
      strGpTopic(4) = "หมายเหตุ"      

      With cmbType

              For i = 0 To 4
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(0)

      End With

      With cmbFilter

              For i = 0 To 4
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(0)

      End With

End Sub


Private Sub InputDeptData()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String = ""
 Dim strDateAdd As String = ""
 Dim strDateEdit As String = ""

 Dim strInDate As String = ""

 Dim intPageCount As Integer
 Dim intPageSize As Integer
 Dim intCounter As Integer

 Dim strSearch As String = txtFilter.Text.ToString.Trim
 Dim strFieldFilter As String = ""

 Dim dteComputer As Date = Now()

 Dim strDateFilter As String = ""
 Dim strYearCnvt As String = ""

 Dim imgStaReq As Image
 Dim staTransfer As String

               With Conn

                        If .State Then .Close()

                            .ConnectionString = strConnAdodb
                            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                            .ConnectionTimeout = 90
                            .Open()

                End With

              If blnHaveFilter Then

                    Select Case cmbFilter.SelectedIndex()

                              Case Is = 0
                                     strFieldFilter = "doc_id like '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 1
                                      strFieldFilter = "send_nm like '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 2
                                      strFieldFilter = "rvc_nm like '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 3
                                      strFieldFilter = "rvc_dep_nm like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 4
                                      strFieldFilter = "remark like '%" & ReplaceQuote(strSearch) & "%'"

                    End Select

                    strSqlCmdSelc = "SELECT * FROM v_delvmst2 (NOLOCK)" _
                                                   & " WHERE " & strFieldFilter _
                                                   & " ORDER BY doc_id DESC"

              Else


                    strSqlCmdSelc = "SELECT * FROM v_delvmst2 (NOLOCK)" _
                                                  & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                  & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                  & " ORDER BY doc_id DESC"

              End If

              intPageSize = 30

              Rsd = New ADODB.Recordset
              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )


                          If .RecordCount <> 0 Then

                                    If intPageSize > .RecordCount Then
                                        intPageSize = .RecordCount
                                    End If

                                    If intPageSize = 0 Then
                                        intPageSize = 30
                                    End If

                                    .PageSize = intPageSize
                                     intPageCount = .PageCount


                                    '--------------------------ถ้ามีการค้นหา----------------------------------------

                                     If strSqlFindData <> "" Then

                                            .MoveFirst()
                                            .Find(strSqlFindData)

                                             If Not .EOF Then
                                                lblPage.Text = Str(.AbsolutePage)
                                             End If

                                            strSqlFindData = ""

                                    End If

                                    '-----------------------------------------------------------------------------

                                    If Int(lblPage.Text.ToString) > intPageCount Then
                                        lblPage.Text = intPageCount.ToString
                                    End If

                                    txtPage.Text = lblPage.Text.ToString
                                    intBkPageCount = .PageCount
                                    lblPageAll.Text = "/ " & .PageCount.ToString
                                    .AbsolutePage = Int(lblPage.Text.ToString)

                                    dgvShoe.Rows.Clear()

                                    intCounter = 0

                                    Do While Not .EOF

                                                  Select Case .Fields("app_sta").Value.ToString.Trim

                                                          Case Is = "0"
                                                                    imgStaReq = My.Resources._16x16ledyellow
                                                                    staTransfer = "กำลังดำเนินการ"
                                                          Case Is = "1"
                                                                    imgStaReq = My.Resources._16x16_ledgreen
                                                                    staTransfer = "โอนแล้ว"
                                                          Case Else
                                                                    imgStaReq = My.Resources._16x16_ledred
                                                                    staTransfer = "รอดำเนินการ"

                                                  End Select


                                                    dgvShoe.Rows.Add( _
                                                                          imgStaReq, staTransfer, _
                                                                          .Fields("doc_id").Value.ToString.Trim, _
                                                                          Mid(.Fields("doc_date").Value.ToString.Trim, 1, 10), _
                                                                          .Fields("send_nm").Value.ToString.Trim, _
                                                                          .Fields("rvc_nm").Value.ToString.Trim, _
                                                                          .Fields("rvc_dep_nm").Value.ToString.Trim, _
                                                                          .Fields("remark").Value.ToString.Trim, _
                                                                           Mid(.Fields("pre_date").Value.ToString.Trim, 1, 10), _
                                                                           .Fields("pre_by").Value.ToString.Trim, _
                                                                           Mid(.Fields("last_date").Value.ToString.Trim, 1, 10), _
                                                                          .Fields("last_by").Value.ToString.Trim _
                                                                          )
                                            intCounter = intCounter + 1

                                              If intCounter = intPageSize Then
                                                    Exit Do
                                              End If

                                         .MoveNext()
                                     Loop

                            Else
                                intBkPageCount = 1
                                txtPage.Text = "1"

                            End If

                   .Close()

              End With

              Rsd = Nothing

              Conn.Close()
              Conn = Nothing


End Sub

Private Sub ViewShoeData()

If dgvShoe.Rows.Count > 0 Then

         ClearTmpTableUser("tmp_delvtrn")
         lblCmd.Text = "2"

         With frmAeDelv
                  .ShowDialog()
                  .Text = "มุมมองข้อมูล"
         End With

         'Me.Hide()
         'frmMainPro.Hide()

End If

End Sub

Private Sub StateLockFind(ByVal Sta As Boolean)
    tabCmd.Enabled = Sta
    dgvShoe.Enabled = Sta
    tlsBarFmr.Enabled = Sta

End Sub

Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click

  lblPage.Text = "1"
  InputDeptData()

End Sub

Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
  lblPage.Text = Str(intBkPageCount)
  InputDeptData()
End Sub

Private Sub btnPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPre.Click

  If Int(lblPage.Text) > 1 Then
        lblPage.Text = Str(Int(lblPage.Text) - 1)
        InputDeptData()
  End If

End Sub

Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
   lblPage.Text = Str(Int(lblPage.Text) + 1)
   InputDeptData()
End Sub

Private Sub txtPage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPage.GotFocus
    txtPage.SelectAll()
End Sub

Private Sub txtPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPage.KeyPress

If e.KeyChar = Chr(13) Then
   dgvShoe.Focus()
End If

End Sub

Private Sub txtPage_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPage.LostFocus
   ChangePage()
End Sub

Private Sub ChangePage()

Dim i, x As Integer
Dim strTmp As String = ""
Dim strMerge As String = ""

Dim intMovePage As Integer

                x = Len(txtPage.Text.ToString.Trim)
                        For i = 1 To x
                                strTmp = Mid(txtPage.Text.ToString.Trim, i, 1)
                                Select Case strTmp
                                          Case Is = ","
                                          Case Is = "+"
                                          Case Is = "-"
                                          Case Is = "_"
                                          Case Is = "."
                                          Case Else
                                                    strMerge = strMerge & Trim(strTmp)
                                End Select
                                strTmp = ""
                        Next i
                Try

                    intMovePage = Int(strMerge)
                    If intMovePage >= Int(lblPage.Text) Then
                            If intMovePage <= intBkPageCount Then

                                    lblPage.Text = intMovePage.ToString.Trim
                                    txtPage.Text = lblPage.Text
                                    InputDeptData()

                            Else

                                    lblPage.Text = intMovePage.ToString.Trim
                                    txtPage.Text = lblPage.Text
                                    InputDeptData()

                            End If
                    Else

                        If intMovePage > 0 Then
                            lblPage.Text = intMovePage.ToString.Trim
                            txtPage.Text = lblPage.Text
                        Else
                            lblPage.Text = "1"
                            txtPage.Text = lblPage.Text
                        End If

                        InputDeptData()

                    End If

                Catch ex As Exception
                    txtPage.Text = lblPage.Text
                End Try

End Sub

Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchTxt As String)

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim intPageCount As Integer
Dim intPageSize As Integer
Dim strSqlCmdSelc As String = ""

Dim strSqlFind As String = ""
Dim i As Integer

Dim strDateFilter As String = ""
Dim strYearCnvt As String = ""


       With Conn
              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

        End With

        Select Case bytColNumber

                         Case Is = 0
                                 strSqlFind = "doc_id "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                         Case Is = 1
                                 strSqlFind = "send_nm "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                         Case Is = 2
                                 strSqlFind = "rvc_nm "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                         Case Is = 3
                                  strSqlFind = "rvc_dep_nm "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                         Case Is = 4
                                  strSqlFind = "remark "
                                  strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"


        End Select


        strSqlCmdSelc = "SELECT * FROM v_delvmst2 (NOLOCK)" _
                                     & " WHERE " & strSqlFind _
                                     & " ORDER BY doc_id DESC"

        intPageSize = 30

        Rsd = New ADODB.Recordset
        With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                          If .RecordCount <> 0 Then

                                        If intPageSize > .RecordCount Then
                                                intPageSize = .RecordCount
                                        End If

                                        If intPageSize = 0 Then
                                                intPageSize = 30
                                        End If

                                        .PageSize = intPageSize
                                        intPageCount = .PageCount

                                        '---------------------------------------ค้นหาข้อมูล------------------------------
                                        .MoveFirst()
                                        .Find(strSqlFind)
                                         lblPage.Text = Str(.AbsolutePage)
                                        '-----------------------------------------------------------------------------

                                      If .Fields("RowNumber").Value >= 2100 Then

                                          dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                                          dubNumberEnd = .Fields("RowNumber").Value + 1000

                                      Else

                                          dubNumberStart = 1
                                          dubNumberEnd = 2100S

                                      End If

                                        strSqlFindData = strSqlFind


                                        InputDeptData()


                                                For i = 0 To dgvShoe.Rows.Count - 1
                                                        If InStr(UCase(dgvShoe.Rows(i).Cells(2).Value), strSearchTxt.Trim.ToUpper) <> 0 Then
                                                                dgvShoe.CurrentCell = dgvShoe.Item(2, i)
                                                                dgvShoe.Focus()
                                                                Exit For
                                                        End If
                                                Next i


                          Else


                               MsgBox("ไม่มีข้อมูล : " & strSearchTxt & " ในระบบ" & vbNewLine _
                                           & "โปรดระบุการค้นหาข้อมูลใหม่!", vbExclamation, "Not Found Data")



                          End If

            .ActiveConnection = Nothing
            .Close()

      End With
      Rsd = Nothing

      Conn.Close()
      Conn = Nothing

StateLockFind(True)
gpbSearch.Visible = False

End Sub

Private Sub dgvShoe_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShoe.CellDoubleClick
 Dim blnReturn As Boolean

    blnReturn = CheckUserEntry(strDocCode, "act_view")
    If blnReturn Then
          ViewShoeData()
    End If

End Sub

Private Sub dgvShoe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvShoe.KeyDown

If e.KeyCode = Keys.Enter Then
    e.Handled = True
End If

End Sub

Private Sub DeleteData()

Dim Conn As New ADODB.Connection

Dim Rsd As New ADODB.Recordset
Dim RsdSb As New ADODB.Recordset

Dim strSqlCmd As String
Dim strSqlSelc As String
Dim strSqlSelcSb As String
Dim strSta As String

Dim btyConsider As Byte
Dim strDocId As String
Dim strDocDate As String

With Conn

              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

End With

With dgvShoe

        If .Rows.Count > 0 Then

                strDocId = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim
                strDocDate = .Rows(.CurrentRow.Index).Cells(3).Value.ToString.Trim

                btyConsider = MsgBox("เลขที่เอกสาร : " & strDocId & vbNewLine _
                                                & "วันที่เอกสาร : " & strDocDate & vbNewLine _
                                                & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then

                        Conn.BeginTrans()

                   '-------------------------------- อัพเดทการส่งมอบที่ตาราง eqptrn,อัพเดทสถานะ ------------------------------------------

                     strSqlSelc = "SELECT [group],eqp_id,size_id,size_desc" _
                                          & " FROM delvtrn (NOLOCK) " _
                                          & " WHERE doc_id ='" & strDocId & "'" _
                                          & " GROUP BY [group],eqp_id,size_id,size_desc"

                    Rsd = New ADODB.Recordset
                    Rsd.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                    Rsd.LockType = ADODB.LockTypeEnum.adLockOptimistic
                    Rsd.Open(strSqlSelc, Conn, , , )

                    If Rsd.RecordCount <> 0 Then

                                Do While Not Rsd.EOF

                                         strSqlCmd = "UPDATE eqptrn SET delvr_sta = '0'" _
                                                               & " WHERE [group] ='" & Rsd.Fields("group").Value.ToString.Trim & "'" _
                                                               & " AND eqp_id ='" & Rsd.Fields("eqp_id").Value.ToString.Trim & "'" _
                                                               & " AND size_id ='" & Rsd.Fields("size_id").Value.ToString.Trim & "'" _
                                                               & " AND size_desc ='" & Rsd.Fields("size_desc").Value.ToString.Trim & "'"

                                         Conn.Execute(strSqlCmd)

                                        Rsd.MoveNext()

                                 Loop

                      End If

                      Rsd.ActiveConnection = Nothing
                      Rsd.Close()
                      Rsd = Nothing

                    '----------------------------อัพเดทสถานะส่งมอบอุปกรณ์ eqpmst ------------------------------

                     strSqlSelc = "SELECT [group],eqp_id" _
                                          & " FROM delvtrn (NOLOCK) " _
                                          & " WHERE doc_id ='" & strDocId & "'" _
                                          & " GROUP BY [group],eqp_id"

                    Rsd = New ADODB.Recordset
                    Rsd.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                    Rsd.LockType = ADODB.LockTypeEnum.adLockOptimistic
                    Rsd.Open(strSqlSelc, Conn, , , )

                    If Rsd.RecordCount <> 0 Then

                                Do While Not Rsd.EOF

                                            strSqlSelcSb = "SELECT [group],eqp_id" _
                                                                       & ",SUM(delvr_pnd) AS delvr1 " _
                                                                       & ",SUM(delvr_snd) AS delvr2 " _
                                                                       & " FROM v_delvr_sta (NOLOCK) " _
                                                                       & " WHERE [group] ='" & Rsd.Fields("group").Value.ToString.Trim & "'" _
                                                                       & " AND eqp_id ='" & Rsd.Fields("eqp_id").Value.ToString.Trim & "'" _
                                                                       & " GROUP BY [group],eqp_id"

                                            RsdSb = New ADODB.Recordset
                                            RsdSb.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                                            RsdSb.LockType = ADODB.LockTypeEnum.adLockOptimistic
                                            RsdSb.Open(strSqlSelcSb, Conn, , , )

                                            If RsdSb.RecordCount <> 0 Then

                                                         If RsdSb.Fields("delvr2").Value = 0 Then
                                                                strSta = "0" 'รอส่งมอบ
                                                         Else
                                                                  If RsdSb.Fields("delvr1").Value > RsdSb.Fields("delvr2").Value Then
                                                                            strSta = "1" 'ส่งมอบไปบางส่วนแล้ว
                                                                  Else
                                                                            strSta = "2" 'ส่งมอบครบ
                                                                  End If

                                                         End If

                                                            strSqlCmd = "UPDATE eqpmst SET prod_sta = '" & strSta & "'" _
                                                                                  & " WHERE [group] ='" & RsdSb.Fields("group").Value.ToString.Trim & "'" _
                                                                                  & " AND eqp_id ='" & RsdSb.Fields("eqp_id").Value.ToString.Trim & "'"

                                                            Conn.Execute(strSqlCmd)

                                             End If

                                             RsdSb.ActiveConnection = Nothing
                                             RsdSb.Close()
                                             RsdSb = Nothing

                                        Rsd.MoveNext()

                                 Loop

                      End If

                      Rsd.ActiveConnection = Nothing
                      Rsd.Close()
                      Rsd = Nothing


                      '------------------------------------ลบตาราง delvmst--------------------------------------------

                        strSqlCmd = "DELETE FROM delvmst" _
                                             & " WHERE doc_id ='" & strDocId & "'"

                        Conn.Execute(strSqlCmd)

                     '------------------------------------ลบตาราง delvtrn--------------------------------------------

                        strSqlCmd = "DELETE FROM delvtrn" _
                                              & " WHERE doc_id ='" & strDocId & "'"

                        Conn.Execute(strSqlCmd)

                        Conn.CommitTrans()

                        .Rows.RemoveAt(.CurrentRow.Index)
                         InputDeptData()

                End If

        End If

      .Focus()

End With

Conn.Close()
Conn = Nothing

End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click

 If blnHaveFilter Then

  blnHaveFilter = False
  InputDeptData()

End If

  StateLockFind(True)
  gpbFilter.Visible = False

End Sub

Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
  FilterData()
End Sub

Sub FilterData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String = ""
Dim strFieldFilter As String = ""

Dim blnHaveData As Boolean

Dim strSearch As String = txtFilter.Text.ToUpper.Trim

Dim strDateFilter As String = ""
Dim strYearCnvt As String = ""

If strSearch <> "" Then


      With Conn

           If .State Then .Close()

              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()

      End With

                    Select Case cmbFilter.SelectedIndex()

                           Case Is = 0
                                     strFieldFilter = "doc_id like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 1
                                      strFieldFilter = "send_nm like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 2
                                      strFieldFilter = "rvc_nm like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 3
                                      strFieldFilter = "rvc_dep_nm like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 4
                                      strFieldFilter = "remark like '%" & ReplaceQuote(strSearch) & "%'"

                    End Select


                    strSqlCmdSelc = "SELECT * FROM v_delvmst (NOLOCK)" _
                                                  & " WHERE " & strFieldFilter _
                                                  & " ORDER BY doc_id DESC"

                    Rsd = New ADODB.Recordset

                    With Rsd

                            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                            .LockType = ADODB.LockTypeEnum.adLockOptimistic
                            .Open(strSqlCmdSelc, Conn, , , )

                            If .RecordCount <> 0 Then
                                blnHaveData = True
                            Else
                                blnHaveData = False
                            End If

                           .ActiveConnection = Nothing
                           .Close()

                   End With

                   Rsd = Nothing


        Conn.Close()
        Conn = Nothing

        If blnHaveData Then

             blnHaveFilter = True
             InputDeptData()

             StateLockFind(True)
             gpbFilter.Visible = False

        Else

            MsgBox("ไม่มีข้อมูลที่ต้องการกรองข้อมูล!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
            txtFilter.Focus()

        End If

Else

    MsgBox("โปรดระบุข้อมูลที่ต้องการกรองก่อน!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
    txtFilter.Focus()

End If

End Sub

Private Sub cmbFilter_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbFilter.KeyPress

If e.KeyChar = Chr(13) Then
    txtFilter.Focus()
End If

End Sub

Private Sub txtFilter_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilter.GotFocus
    txtFilter.SelectAll()
End Sub

Private Sub txtFilter_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFilter.KeyPress

If e.KeyChar = Chr(13) Then
    FilterData()
End If

End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

  StateLockFind(True)
  gpbSearch.Visible = False

End Sub


Sub FindDocID()
Dim strSearch As String = txtSeek.Text.ToUpper.Trim

If strSearch <> "" Then

        Select Case cmbType.SelectedIndex()

                  Case Is = 0 'เลขที่เอกสาร
                          SearchData(0, strSearch)

                  Case Is = 1 'รายละเอียดอุปกรณ์
                          SearchData(2, strSearch)

                  Case Is = 2 'ชิ้นงาน
                          SearchData(3, strSearch)

                  Case Is = 3 'แผนกรับมอบ
                          SearchData(4, strSearch)

                  Case Is = 4 'หมายเหตุ
                          SearchData(5, strSearch)

        End Select

Else

    MsgBox("ไม่มีข้อมูลที่ต้องการค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
    txtSeek.Focus()

End If

End Sub

Private Sub cmbType_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbType.KeyPress

If e.KeyChar = Chr(13) Then
    txtSeek.Focus()
End If

End Sub

Private Sub txtSeek_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSeek.GotFocus
    txtSeek.SelectAll()
End Sub

Private Sub txtSeek_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSeek.KeyPress

If e.KeyChar = Chr(13) Then
    FindDocID()
End If

End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  FindDocID()
End Sub

Private Sub dgvShoe_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvShoe.KeyPress

 Dim blnReturn As Boolean
     If e.KeyChar = Chr(13) Then

        blnReturn = CheckUserEntry(strDocCode, "act_view")
        If blnReturn Then
            ViewShoeData()
        End If

      End If

End Sub

    Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

        Dim blnReturn As Boolean
        Dim strEqpId As String = ""

        With tabCmd

            Select Case .SelectedIndex

                Case Is = 0 'เพิ่มข้อมูล

                    blnReturn = CheckUserEntry(strDocCode, "act_add")
                    If blnReturn Then

                        ClearTmpTableUser("tmp_delvtrn")
                        lblCmd.Text = "0"

                        With frmAeDelv
                            .ShowDialog()
                            .Text = "เพิ่มข้อมูล"

                        End With

                    Else
                        MsnAdmin()
                    End If


                Case Is = 1 'แก้ไขข้อมูล

                    blnReturn = CheckUserEntry(strDocCode, "act_edit")
                    If blnReturn Then

                        If dgvShoe.Rows.Count > 0 Then

                            If checkCompleteApprove() Then         'ตรวจสอบว่าผู้อนุมัติรับของเซ็นแล้วหรือไม่ (กรณีเซ็นแล้วไม่ให้แก้ไข)

                                ClearTmpTableUser("tmp_delvtrn")
                                lblCmd.Text = "1"

                                With frmAeDelv
                                    .ShowDialog()
                                    .Text = "แก้ไขข้อมูล"

                                End With

                            Else
                                MessageBox.Show("เอกสาร Approve แล้ว ไม่สามารถแก้ไขได้",
                                                       "Access denied!..", MessageBoxButtons.OK, MessageBoxIcon.Warning)


                            End If

                        End If

                    Else
                        MsnAdmin()
                    End If

                Case Is = 2 'มุมมอง

                    blnReturn = CheckUserEntry(strDocCode, "act_view")
                    If blnReturn Then
                        ViewShoeData()
                    Else
                        MsnAdmin()
                    End If

                Case Is = 3 'กรองข้อมูล

                    If dgvShoe.Rows.Count > 0 Then

                        With gpbFilter

                            .Top = 230
                            .Left = 210
                            .Width = 348
                            .Height = 125

                            .Visible = True

                            cmbFilter.SelectedItem = cmbFilter.Items(0)
                            txtFilter.Text =
                            dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                            StateLockFind(False)
                            txtFilter.Focus()

                        End With

                    End If

                Case Is = 4 'ค้นหาข้อมูล

                    If dgvShoe.Rows.Count > 0 Then

                        With gpbSearch

                            .Top = 230
                            .Left = 210
                            .Width = 348
                            .Height = 125

                            .BringToFront()
                            .Visible = True

                            cmbType.SelectedItem = cmbType.Items(0)
                            txtSeek.Text =
                            dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                            StateLockFind(False)
                            txtSeek.Focus()

                        End With

                    End If

                Case Is = 6 'พิมพ์ข้อมูล

                    If dgvShoe.Rows.Count > 0 Then

                        With gpbOptPrint

                            .Top = 200
                            .Left = 270
                            .Width = 311
                            .Height = 125
                            .Visible = True

                            FillPoDocData()
                            cmbOptPrint.SelectedItem =
                            dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                            StateLockFind(False)
                            cmbOptPrint.Focus()

                        End With

                    End If

                Case Is = 5 'ฟื้นฟูข้อมูล
                    blnHaveFilter = False
                    InputDeptData()

                Case Is = 7 'ลบข้อมูล

                    blnReturn = CheckUserEntry(strDocCode, "act_delete")
                    If blnReturn Then
                        DeleteData()
                    Else
                        MsnAdmin()
                    End If

                Case Is = 8 'ออก
                    Me.Close()

            End Select

        End With

    End Sub

    Private Function checkCompleteApprove() As Boolean

        Dim Conn As New ADODB.Connection
        Dim Rsd As New ADODB.Recordset

        Dim strSqlCmdSelc As String
        Dim strDocid As String

        strDocid = dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(2).Value.ToString.Trim()

        With Conn

            If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

        End With


        strSqlCmdSelc = "SELECT app_sta FROM delv_approve (NOLOCK)" _
                                      & " WHERE doc_id = '" & strDocid & "'"


        Rsd = New ADODB.Recordset

        With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlCmdSelc, Conn, , , )

             If .RecordCount <> 0 Then

                 If .Fields("app_sta").Value = True Then

                   Return False

                 Else

                    Return True

                  End If

             Else

                Return True

             End If

      .ActiveConnection = Nothing
      .Close()
     End With

  Conn.Close()
  Conn = Nothing

End Function

Private Sub FillPoDocData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String

    With Conn

              If .State Then .Close()
                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

    End With


    strSqlCmdSelc = "SELECT doc_id FROM v_delvmst (NOLOCK)" _
                                      & " ORDER BY doc_id DESC"


     Rsd = New ADODB.Recordset

     With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlCmdSelc, Conn, , , )

             If .RecordCount <> 0 Then

                    cmbOptPrint.Items.Clear()

                    Do While Not .EOF
                          cmbOptPrint.Items.Add(.Fields("doc_id").Value.ToString.Trim)
                         .MoveNext()

                    Loop

             End If

            .ActiveConnection = Nothing
            .Close()

    End With
    Rsd = Nothing


Conn.Close()
Conn = Nothing

End Sub


Private Function SizeImage(ByVal img As Bitmap, ByVal width As Integer, ByVal height As Integer) As Bitmap

        Dim newBit As New Bitmap(width, height) 'new blank bitmap
        Dim g As Graphics = Graphics.FromImage(newBit)
        'change interpolation for reduction quality
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(img, 0, 0, width, height)
        Return newBit

End Function

Private Sub dgvShoe_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvShoe.RowsAdded
 dgvShoe.Rows(e.RowIndex).Height = 27
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

        strSqlSelc = "SELECT doc_id FROM v_delvmst (NOLOCK)" _
                             & " ORDER BY doc_id DESC"

       RsdPnt = New ADODB.Recordset

       With RsdPnt

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlSelc, Conn, , , )

               If .RecordCount <> 0 Then

                        ds.Clear()
                        da.Fill(ds, RsdPnt, "eqpid")
                        cmbOptPrint.DataSource = ds.Tables("eqpid").DefaultView
                        cmbOptPrint.DisplayMember = "doc_id"
                        cmbOptPrint.ValueMember = "doc_id"

                End If

               .ActiveConnection = Nothing
             ' .Close()

     End With

     RsdPnt = Nothing
End Sub

Private Sub btnPrntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntCancel.Click
    StateLockFind(True)
    gpbOptPrint.Visible = False

End Sub

Private Sub btnPrntPrevw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntPrevw.Click
  Dim strDocId As String = cmbOptPrint.Text.ToUpper.Trim

  If strDocId <> "" Then

     frmMainPro.lblRptCentral.Text = "B"
     frmMainPro.lblRptDesc.Text = strDocId

     frmRptDelvApprove.Show()
     'frmRptAccept.Show()  ' แสดง frmRptCentral()

     StateLockFind(True)
     gpbOptPrint.Visible = False
     frmMainPro.Hide()

   Else
     MsgBox("โปรดระบุข้อมูลก่อนพิมพ์", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Equipment Empty!!")
     cmbOptPrint.Focus()

    End If
End Sub

Private Sub PrePrintData(ByVal strSelectCode As String)
Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset
Dim RsdPic As New ADODB.Recordset

Dim strSqlSelc As String
Dim strSqlCmdPic As String

Dim strPicPath As String = "H:\EquipPicture\"

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

     strSqlSelc = "SELECT * " _
                        & " FROM v_delvmst2 (NOLOCK)" _
                        & " WHERE doc_id = '" & strSelectCode.ToString.Trim & "'"

     Rsd = New ADODB.Recordset

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

                                      '----------------------------------------LoadPicture ภายใน/ภายนอก----------------------------------------------

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

                                    '----------------------------------- เพิ่มข้อมูล ไปที่ tmp_delvmst -------------------------------

                                       strSqlCmdPic = "SELECT * " _
                                                                   & " FROM tmp_delvmst (NOLOCK)"

                                       RsdPic = New ADODB.Recordset
                                       RsdPic.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                                       RsdPic.LockType = ADODB.LockTypeEnum.adLockOptimistic
                                       RsdPic.Open(strSqlCmdPic, Conn, , , )

                                                    RsdPic.AddNew()
                                                    RsdPic.Fields("user_id").Value = frmMainPro.lblLogin.Text.ToString.Trim
                                                    RsdPic.Fields("prod_sta").Value = .Fields("prod_sta").Value
                                                    RsdPic.Fields("fix_sta").Value = .Fields("fix_sta").Value
                                                    RsdPic.Fields("group").Value = .Fields("group").Value
                                                    RsdPic.Fields("doc_id").Value = .Fields("doc_id").Value
                                                    RsdPic.Fields("send_nm").Value = .Fields("send_nm").Value
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

                                                    '----------------------------เพิ่มข้อมูลรูปภาพบรรจุ---------------------------------------------------

                                                    If blnHavePic1 Then 'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                             Dim RsdSteam1 As New MemoryStream
                                                             Dim bytes1 = File.ReadAllBytes(strLoadFilePic1)

                                                            inImg = Image.FromFile(strLoadFilePic1)
                                                            inImg = SizeImage(inImg, 230, 200)
                                                            inImg.Save(RsdSteam1, ImageFormat.Bmp)
                                                            bytes1 = RsdSteam1.ToArray
                                                            RsdPic.Fields("bob_ctain").Value = bytes1

                                                            RsdSteam1.Close()
                                                            RsdSteam1 = Nothing

                                                    End If

                                                     '----------------------------เพิ่มข้อมูลรูปภาพภายนอก/ภายใน---------------------------------------------------

                                                    If blnHavePic2 Then 'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                             Dim RsdSteam2 As New MemoryStream
                                                             Dim bytes2 = File.ReadAllBytes(strLoadFilePic2)

                                                            inImg = Image.FromFile(strLoadFilePic2)
                                                            inImg = SizeImage(inImg, 230, 200)
                                                            inImg.Save(RsdSteam2, ImageFormat.Bmp)
                                                            bytes2 = RsdSteam2.ToArray
                                                            RsdPic.Fields("bob_io").Value = bytes2

                                                            RsdSteam2.Close()
                                                            RsdSteam2 = Nothing

                                                    End If

                                                    '----------------------------เพิ่มข้อมูลรูปภาพชิ้นงาน---------------------------------------------------

                                                    If blnHavePic3 Then 'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                             Dim RsdSteam3 As New MemoryStream
                                                             Dim bytes3 = File.ReadAllBytes(strLoadFilePic3)

                                                            inImg = Image.FromFile(strLoadFilePic3)
                                                            inImg = SizeImage(inImg, 230, 200)
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

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  InputDeptData()
End Sub
End Class