Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream
Imports System.Windows.Forms.DataGridView

Public Class frmFixRecv

Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean    'กรณีกรองข้อมูล
Dim IsShowSeek As Boolean

Dim dubNumberStart As Double   'ถูกกำหนด = 1
Dim dubNumberEnd As Double     'ถูกกำหนด = 2100

Dim strSqlFindData As String
Dim strDocCode As String = "F7"

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim dsTn As New DataSet

Private Sub frmFixRecv_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

     If FormCount("frmAeFixRecv") > 0 Then

        With frmAeFixRecv

               strSearch = .lblComplete.Text          'รหัสซ่อม

                If strSearch <> "" Then
                   SearchData(0, strSearch)
                End If

                .Close()

        End With

     End If
     Timer1.Enabled = True          'สั่งรีเฟรชข้อมูลทุก 1 นาที

    Me.Height = Int(lblHeight.Text)
    Me.Width = Int(lblWidth.Text)

    Me.Top = Int(lblTop.Text)
    Me.Left = Int(lblLeft.Text)

End Sub

Private Function FormCount(ByVal frmName As String) As Long
  Dim frm As Form

    For Each frm In My.Application.OpenForms

         If frm Is My.Forms.frmAeFixRecv Then
            FormCount = FormCount + 1
         End If
    Next

End Function

Private Sub frmFixRecv_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

  lblHeight.Text = Me.Height.ToString
  lblWidth.Text = Me.Width.ToString

  lblTop.Text = Me.Top.ToString
  lblLeft.Text = Me.Left.ToString

End Sub

Private Sub frmFixRecv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 Me.Dispose()
End Sub

Private Sub frmFixRecv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

 Me.WindowState = FormWindowState.Maximized     'ขยายขนาดเต็มหน้าจอ
     StdDateTimeThai()                           'เรียก ซับรูทีน StdDateTimeThai
     tlsBarFmr.Cursor = Cursors.Hand             'ให้คอร์เซอร์ตรง Toolstripbar เป็นรูปมือ

     dubNumberStart = 1                          'ให้แถวเเรกใน Recordset = 1
     dubNumberEnd = 2100                         'ให้แถวเเรกใน Recordset = 2100

     PreGroupType()

     InputData()
     tabCmd.Focus()

End Sub

Private Sub PreGroupType()
  Dim strGpTopic(4) As String
  Dim i As Byte

      strGpTopic(0) = "รหัสส่งซ่อม"
      strGpTopic(1) = "รหัสอุปกรณ์"
      strGpTopic(2) = "รายละเอียดอุปกรณ์"
      strGpTopic(3) = "กลุ่มอุปกรณ์"
      strGpTopic(4) = "สถานะส่งซ่อม"

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

Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchtxt As String)

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
                                 strSqlFind = "fix_id "
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 2
                                 strSqlFind = "eqp_id "
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 3
                                 strSqlFind = "eqp_name"
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 4
                                 strSqlFind = "desc_thai "
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 5
                                  strSqlFind = "fix_sta"
                                  strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

           End Select

           strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                        & " WHERE " & strSqlFind _
                                        & " ORDER BY eqp_id"

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

                            '---------------------------------------ค้นหาข้อมูล-------------------------------------------------------------

                            .MoveFirst()
                            .Find(strSqlFind)
                            lblPage.Text = Str(.AbsolutePage)

                            '-------------------------------------------------------------------------------------------------------------

                            If .Fields("RowNumber").Value >= 2100 Then

                               dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                               dubNumberEnd = .Fields("RowNumber").Value + 1000

                            Else

                               dubNumberStart = 1
                               dubNumberEnd = 2100

                            End If

                                strSqlFindData = strSqlFind

                                InputData()

                                       For i = 0 To dgvFix.Rows.Count - 1

                                            If InStr(UCase(dgvFix.Rows(i).Cells(2).Value), strSearchtxt.Trim.ToUpper) <> 0 Then
                                               dgvFix.CurrentCell = dgvFix.Item(2, i)
                                               dgvFix.Focus()
                                               Exit For

                                            End If

                                       Next i

                Else

                     MsgBox("ไม่มีข้อมูล : " & strSearchtxt & " ในระบบ" & vbNewLine _
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

Private Sub StateLockFind(ByVal sta As Boolean)
  tabCmd.Enabled = sta
  dgvFix.Enabled = sta
  tlsBarFmr.Enabled = sta

End Sub

Private Sub InputData()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String = ""
 Dim strDateAdd As String = ""
 Dim strDateEdit As String = ""

 Dim strInDate As String = ""

 Dim intPageCount As Integer          'จำนวนหน้าทั้งหมด
 Dim intPageSize As Integer           'จำนวนรายการใน 1 หน้า
 Dim intCounter As Integer

 Dim strSearch As String = txtFilter.Text.ToString.Trim      'กรองข้อมูล
 Dim strFieldFilter As String = ""

 Dim dteComputer As Date = Now()
 Dim imgStaFix As Image               'รูปสถานะส่งซ่อม

 Dim strDateFilter As String = ""
 Dim strYearCnvt As String = ""

       With Conn

            If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

          If blnHaveFilter Then          'กรณีเลือก กรองข้อมูล

                    Select Case cmbFilter.SelectedIndex()

                               Case Is = 0
                                     strFieldFilter = "fix_id like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 1
                                      strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 2
                                      strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 3
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 4
                                      strFieldFilter = "fix_sta like '%" & ReplaceQuote(strSearch) & "%'"

                    End Select

                    strSqlCmdSelc = "SELECT  * FROM v_fixeqptrn (NOLOCK)" _
                                                   & " WHERE " & strFieldFilter _
                                                   & " AND fix_sta = '2'" _
                                                   & " ORDER BY fix_id"
        Else

              strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                          & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                          & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                          & " AND fix_sta = '2'" _
                                          & " ORDER BY fix_id"

              End If

              intPageSize = 30   'ตัวแปรกำหนดขนาดกระดาษ

              Rsd = New ADODB.Recordset
              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )



                          If .RecordCount <> 0 Then

                                    If intPageSize > .RecordCount Then    'ถ้าจำนวนรายการใน 1 page(30) > จำนวนเรคคอร์ดที่ qurey มา
                                        intPageSize = .RecordCount
                                    End If

                                    If intPageSize = 0 Then
                                        intPageSize = 30
                                    End If

                                    .PageSize = intPageSize        '.PageSize ใช้กำหนดว่าแต่ละหน้าจะให้มีกี่รายการ ในการแสดงผล
                                     intPageCount = .PageCount     '.PageCount นับจำนวนหน้าทั้งหมด ที่ได้จากการกำหนดขนาดของหน้า


                                    '--------------------------กรณีมีการค้นหา-----------------------------------

                                     If strSqlFindData <> "" Then

                                            .MoveFirst()
                                            .Find(strSqlFindData)

                                             If Not .EOF Then
                                                lblPage.Text = Str(.AbsolutePage)    '.AbsolutePage ใช้อ้างอิงไปยังหน้าที่ต้องการ
                                             End If

                                            strSqlFindData = ""

                                     End If


                                    '---------- กำหนดปุ่ม ใน tlsBarFmr ----------------------------------------

                                    If Int(lblPage.Text.ToString) > intPageCount Then
                                        lblPage.Text = intPageCount.ToString
                                    End If

                                    txtPage.Text = lblPage.Text.ToString
                                    intBkPageCount = .PageCount
                                    lblPageAll.Text = "/ " & .PageCount.ToString
                                    .AbsolutePage = Int(lblPage.Text.ToString)

                                    dgvFix.Rows.Clear()            'เคลียร์ Gridview ก่อน Inputdata

                                    intCounter = 0

                                    Do While Not .EOF

                                 '--------------------------------------- สถานะส่งซ่อม --------------------------------------------------

                                            Select Case .Fields("fix_sta").Value.ToString.Trim

                                                   Case Is = "1"     'ส่งซ่อม
                                                        imgStaFix = My.Resources.sign_deny

                                                   Case Is = "2"     'รับคืนส่งซ่อม
                                                        imgStaFix = My.Resources.accept

                                                   Case Else         'ปกติ
                                                        imgStaFix = My.Resources.blank

                                            End Select

                                            dgvFix.Rows.Add( _
                                                                imgStaFix, _
                                                                .Fields("fix_desc").Value.ToString.Trim, _
                                                                .Fields("fix_id").Value.ToString.Trim, _
                                                                .Fields("eqp_id").Value.ToString.Trim, _
                                                                "#" & .Fields("size_id").Value.ToString.Trim, _
                                                                .Fields("issue").Value.ToString.Trim, _
                                                                .Fields("fix_issue").Value.ToString.Trim, _
                                                                .Fields("fix_amount").Value.ToString.Trim, _
                                                                Format(.Fields("fix_price").Value, "##,##0.00"), _
                                                                 Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                                                 Mid(.Fields("due_date").Value.ToString.Trim, 1, 10), _
                                                                 Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("recv_by").Value.ToString.Trim, _
                                                                 Mid(.Fields("pre_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("pre_by").Value.ToString.Trim, _
                                                                 Mid(.Fields("last_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("last_by").Value.ToString.Trim _
                                                             )

                                              intCounter = intCounter + 1

                                              If intCounter = intPageSize Then
                                                    Exit Do
                                              End If

                                         .MoveNext()            'ข้ามไปที่ระเบียนใหม่
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

Private Sub InputGpbRecv()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String = ""
 Dim strDateAdd As String = ""
 Dim strDateEdit As String = ""

 Dim strInDate As String = ""

 Dim dteComputer As Date = Now()
 Dim imgStaFix As Image               'รูปสถานะส่งซ่อม

 Dim strDateFilter As String = ""
 Dim strYearCnvt As String = ""
 Dim ChkboxSta As Boolean

       With Conn

            If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

              strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                          & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                          & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                          & " AND fix_sta = '1'" _
                                          & " ORDER BY fix_id"

              Rsd = New ADODB.Recordset
              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )



                          If .RecordCount <> 0 Then

                                    dgvShow.Rows.Clear()      'เคลียร์ datagrid

                                    Do While Not .EOF

                                    '--------------------------- สถานะส่งซ่อม ------------------------------

                                            Select Case .Fields("fix_sta").Value.ToString.Trim

                                                   Case Is = "1"     'ส่งซ่อม
                                                        imgStaFix = My.Resources.sign_deny
                                                        ChkboxSta = False     '

                                                   Case Is = "2"     'รับคืนส่งซ่อม
                                                        imgStaFix = My.Resources.accept
                                                        ChkboxSta = True     'ให้ chekbox ถูกเลือก

                                                   Case Else         'ปกติ
                                                        imgStaFix = My.Resources.blank

                                            End Select

                                              dgvShow.Rows.Add( _
                                                                ChkboxSta, _
                                                                imgStaFix, _
                                                                .Fields("fix_desc").Value.ToString.Trim, _
                                                                .Fields("fix_id").Value.ToString.Trim, _
                                                                .Fields("eqp_id").Value.ToString.Trim, _
                                                                .Fields("eqp_name").Value.ToString.Trim, _
                                                                "#" & .Fields("size_id").Value.ToString.Trim, _
                                                                .Fields("sup_name").Value.ToString.Trim, _
                                                                 Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("recv_by").Value.ToString.Trim _
                                                              )

                                         .MoveNext()            'ข้ามไปที่ระเบียนใหม่
                                     Loop

                            End If

                   .Close()

              End With

              Rsd = Nothing

  Conn.Close()
  Conn = Nothing
End Sub

Private Function chkfxtranc() As Boolean    'เช็คว่าใน fixeqptrn fix_sta = 1  หรือไม่
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlCmdSelc As String = ""

       With Conn

            If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

              strSqlCmdSelc = "SELECT * FROM fixeqptrn (NOLOCK)" _
                                          & " WHERE fix_sta = '1'"

              Rsd = New ADODB.Recordset

              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                        If .RecordCount <> 0 Then
                           Return True

                        Else
                           Return False

                        End If

             .ActiveConnection = Nothing
             .Close()
             End With

  Conn.Close()
  Conn = Nothing
End Function

Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

 Dim btnReturn As Boolean
 Dim strFixSta As String = ""           'เก็บสถานะส่งซ่อม

     With tabCmd

         Select Case tabCmd.SelectedIndex

                Case Is = 0  'รับเข้าส่งซ่อม

                     If chkfxtranc() Then       'เช็คว่าใน fixeqptrn fix_sta = 1  หรือไม่

                           IsShowSeek = Not IsShowSeek
                           If IsShowSeek Then

                               With gpbfxrecv             'groupbox รับเข้าอุปกรณ์

                                   .Visible = True
                                   .Left = 285
                                   .Top = 200
                                   .Height = 347
                                   .Width = 795

                                   dgvShow.Rows.Clear()
                                   InputGpbRecv()          'นำเข้าข้อมูลส่่งซ่อม
                                   StateLockFindDept(False)

                               End With

                            Else
                                 StateLockFindDept(True)

                            End If

                     End If


                Case Is = 1  'แก้ไขข้อมูล

                    If dgvFix.Rows.Count <> 0 Then

                         btnReturn = CheckUserEntry(strDocCode, "act_edit")
                         If btnReturn Then

                            ClearTmpTableUser("tmp_fixeqptrn")
                            lblCmd.Text = "1"                     'เพื่อกำหนดว่าเป็นการแก้ไข

                            With frmAeFixRecv
                                 .Show()
                                 .Text = "แก้ไขข้อมูล"

                            End With

                            Me.Hide()
                            frmMainPro.Hide()

                         Else
                            MsnAdmin()
                         End If

                    End If


                Case Is = 2    'มุมมอง

                   If dgvFix.Rows.Count <> 0 Then

                                btnReturn = CheckUserEntry(strDocCode, "act_view")
                                If btnReturn Then
                                   ViewShoeData()

                                Else
                                    MsnAdmin()

                                End If

                   End If


                Case Is = 3   'กรองข้อมูล

                     If dgvFix.Rows.Count <> 0 Then

                        With gpbFilter

                        .Top = 230
                        .Left = 210
                        .Width = 348
                        .Height = 125

                        .Visible = True

                        cmbFilter.Text = cmbFilter.Items(0)
                        txtFilter.Text = _
                                   dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(2).Value.ToString.Trim

                        StateLockFind(False)
                        txtFilter.Focus()

                        End With

                     End If


                Case Is = 4   'ค้นหาข้อมูล

                     If dgvFix.Rows.Count <> 0 Then

                        With gpbSearch

                             .Top = 230
                             .Left = 210
                             .Width = 348
                             .Height = 125

                             .Visible = True

                             cmbType.Text = cmbType.Items(0)
                             txtSeek.Text = _
                                     dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(2).Value.ToString.Trim

                             StateLockFind(False)
                             txtSeek.Focus()

                        End With

                     End If


                Case Is = 6            'พิมพ์ข้อมูล

                    If dgvFix.Rows.Count > 0 Then

                        ClearTmpTableUser("tmp_fixeqptrn")

                        frmMainPro.lblRptCentral.Text = "G"     ' ส่งค่าให้ตัวเเปรฟอร์ม MainPro 

                        '------------------------- ส่งค่าให้ตัวแปร lblRptDesc ของฟอร์ม MainPro โดยส่ง Userid กับ Eqpid ----------------------------- 

                        frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim & "'"

                        frmRptCentral.Show()

                        StateLockFind(True)
                        frmMainPro.Hide()

                   Else
                        MsgBox("โปรดระบุข้อมูลก่อนพิมพ์", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Equipment Empty!!")

                   End If

               Case Is = 5           'ฟื้นฟูข้อมูล
                    InputData()

               Case Is = 7           'ลบข้อมูล

                      btnReturn = CheckUserEntry(strDocCode, "act_delete")
                            If btnReturn Then
                                DeleteData()
                            Else
                                MsnAdmin()
                            End If

               Case Is = 8           'ออก
                    Me.Close()

         End Select

    End With
End Sub

Private Sub dgvShow_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShow.CellClick
  On Error Resume Next
     With dgvShow

          If .Rows.Count <> 0 Then

               Select Case .CurrentCell.ColumnIndex

                      Case Is = 0

                          If dgvShow.Rows(e.RowIndex).Cells("recv").Value = False Then          'ตรวจสอบว่า chekboxถูกเลือกอยู่แล้วหรือไม่ (กรณ๊ยังไม่ถูกเลือก)

                                If Convert.ToBoolean(dgvShow.Rows(e.RowIndex).Cells("recv").Value) = True Then      '0 = false 1 = true

                                   dgvShow.Rows(e.RowIndex).Cells("recv").Value = False

                                Else

                                    With frmAeFixRecv
                                         .Show()
                                         lblCmd.Text = "0"    'บ่งบอกว่าเป็นการรับเข้า

                                    End With
                                    Me.Hide()
                                    frmMainPro.Hide()

                                End If

                           StateLockFindDept(True)
                           IsShowSeek = False
                           gpbfxrecv.Visible = False


                           End If

                End Select

          End If

     End With

End Sub

Private Sub UpdateData()                         ' UPDATE สถานะ fix_sta จาก 1='ส่งซ่อม' เป็น 2='รับคืนส่งซ่อม'
 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim strDate As Date = Now()
 Dim strFixID As String
 Dim strEqpID As String
 Dim strSize As String

     strFixID = dgvFix.Rows(dgvFix.CurrentCell.ColumnIndex).Cells(2).Value.ToString
     strEqpID = dgvFix.Rows(dgvFix.CurrentCell.ColumnIndex).Cells(3).Value.ToString
     strSize = dgvFix.Rows(dgvFix.CurrentCell.ColumnIndex).Cells(4).Value.ToString

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

                 strSqlCmd = "UPDATE fixeqpmst SET fix_sta = '2'" _
                                & " WHERE fix_id = '" & strFixID & "'" _
                                & " AND eqp_id = '" & strEqpID & "'"


                 Conn.Execute(strSqlCmd)

                 strSqlCmd = "UPDATE fixeqptrn SET fix_sta = '2'" _
                                & " WHERE fix_id = '" & strFixID & "'" _
                                & " AND eqp_id = '" & strEqpID & "'" _
                                & " AND size_id = '" & strSize & "'"

                 Conn.Execute(strSqlCmd)


 Conn.Close()
 Conn = Nothing

End Sub

Private Sub dgvFix_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvFix.CellMouseDoubleClick
 ViewShoeData()
End Sub

Private Sub dgvFix_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvFix.KeyDown

 If e.KeyCode = Keys.Enter Then
     e.Handled = True
  End If

End Sub

Private Sub dgvFix_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvFix.KeyPress

  Dim blnReturn As Boolean
      If e.KeyChar = Chr(13) Then

         blnReturn = CheckUserEntry(strDocCode, "act_view")
         If blnReturn Then
            ViewShoeData()
         End If

      End If

End Sub

Private Sub ViewShoeData()

  If dgvFix.Rows.Count <> 0 Then

     ClearTmpTableUser("tmp_fixeqptrn")
     lblCmd.Text = "2"

     With frmAeFixRecv
          .Show()
          .Text = "มุมมองข้อมูล"

     End With

     Me.Hide()
     frmMainPro.Hide()

  Else
     MsnAdmin()
  End If

End Sub

Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
  lblPage.Text = "1"
  InputData()
End Sub

Private Sub btnPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPre.Click
  If Int(lblPage.Text) > 1 Then
     lblPage.Text = Str(Int(lblPage.Text) - 1)
     InputData()
  End If
End Sub

Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
 lblPage.Text = Str(Int(lblPage.Text) + 1)
 InputData()
End Sub

Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
 lblPage.Text = Str(intBkPageCount)
 InputData()
End Sub

Private Sub txtPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPage.KeyPress

 If e.KeyChar = Chr(13) Then
   dgvFix.Focus()
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
                                    InputData()

                            Else

                                    lblPage.Text = intMovePage.ToString.Trim
                                    txtPage.Text = lblPage.Text
                                    InputData()

                            End If
                    Else

                        If intMovePage > 0 Then
                            lblPage.Text = intMovePage.ToString.Trim
                            txtPage.Text = lblPage.Text
                        Else
                            lblPage.Text = "1"
                            txtPage.Text = lblPage.Text
                        End If

                        InputData()

                    End If

                Catch ex As Exception
                    txtPage.Text = lblPage.Text
                End Try
End Sub

Private Sub dgvFix_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFix.RowsAdded
  dgvFix.Rows(e.RowIndex).Height = 27
End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  SearchDT()
End Sub

Private Sub SearchDT()                                        'ค้นหาเอกสาร
 Dim strSearch As String = txtSeek.Text.ToUpper.Trim

 If strSearch <> "" Then

           Select Case cmbType.SelectedIndex()

                  Case Is = 0 'รหัสส่งซ่อม
                          SearchData(0, strSearch)             'ส่งตำเงื่อนไข ,Text ให้ ซับรูทีน SearchData

                  Case Is = 1 'รหัสอุปกรณ์
                          SearchData(2, strSearch)

                  Case Is = 2 'รายละเอียดอุปกรณ์
                          SearchData(3, strSearch)

                  Case Is = 3 'กลุ่มอุปกรณ์
                          SearchData(4, strSearch)

                  Case Is = 4 'สถานะส่งซ่อม
                          SearchData(5, strSearch)

          End Select

 Else
       MsgBox("โปรดกรอกข้อมูลเพื่อค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
       txtSeek.Focus()

 End If
End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
 StateLockFind(True)
 gpbSearch.Visible = False
End Sub

Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
  FindDocID()
End Sub

Private Sub FindDocID()     'ค้นหาเอกสาร

 Dim strSearch As String = txtFilter.Text.ToUpper.Trim

     If strSearch <> "" Then

           Select Case cmbFilter.SelectedIndex()

                  Case Is = 0 'รหัสส่งซ่อม
                          SearchData(0, strSearch)     'ส่งตำเงื่อนไข ,Text ให้ ซับรูทีน SearchData

                  Case Is = 1 'รหัสอุปกรณ์
                          SearchData(2, strSearch)

                  Case Is = 2 'รายละเอียดอุปกรณ์
                          SearchData(3, strSearch)

                  Case Is = 3 'กลุ่มอุปกรณ์
                          SearchData(4, strSearch)

                  Case Is = 4 'สถานะส่งซ่อม
                          SearchData(5, strSearch)

          End Select

    Else
         MsgBox("โปรดกรอกข้อมูลเพื่อค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
         txtSeek.Focus()

    End If

End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click
   StateLockFind(True)
   gpbFilter.Visible = False
End Sub

Private Sub DeleteData()
 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim btyConsider As Byte
 Dim strFixID As String
 Dim strEqpID As String
 Dim strSize As String
 Dim strPrice As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

     With dgvFix

          If .Rows.Count <> 0 Then

             strFixID = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim
             strEqpID = .Rows(.CurrentRow.Index).Cells(3).Value.ToString.Trim
             strSize = .Rows(.CurrentRow.Index).Cells(4).Value.ToString.Trim
             strPrice = .Rows(.CurrentRow.Index).Cells(8).Value.ToString.Trim
             strSize = Mid(strSize, 2)          'ตัด # ออก

             btyConsider = MsgBox("รหัสอุปกรณ์: " & strEqpID & vbNewLine _
                                               & "Size : " & strSize & vbNewLine _
                                               & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "ยืนยันการลบข้อมูล")

                If btyConsider = 6 Then

                       If chkFixData(strFixID) Then   'ถ้าข้อมูลใน fixeqptrn มีเพียง 1 เรคคอร์ด

                          Conn.BeginTrans()

                          '---------------------------- ลบตาราง fixeqptrn --------------------------------------------

                          strSqlCmd = "DELETE FROM fixeqptrn" _
                                               & " WHERE fix_id ='" & strFixID & "'" _
                                               & " AND size_id = '" & strSize & "'"

                          Conn.Execute(strSqlCmd)
                          Conn.CommitTrans()

                         .Rows.RemoveAt(.CurrentRow.Index)


                         '------------------------------------ ลบตาราง fixeqpmst ----------------------------------------

                          strSqlCmd = "DELETE FROM fixeqpmst" _
                                                 & " WHERE fix_id ='" & strFixID & "'"

                          Conn.Execute(strSqlCmd)

                          InputData()  'อัพเดทข้อมูลใน datagrid


                       Else

                          Conn.BeginTrans()

                          ChangFixPrice()  'update fix_price ทุกครั้งเมื่อลบ ข้อมูล

                          '---------------------------- ลบตาราง fixeqptrn --------------------------------------------

                          strSqlCmd = "DELETE FROM fixeqptrn" _
                                               & " WHERE fix_id ='" & strFixID & "'" _
                                               & " AND size_id = '" & strSize & "'"

                          Conn.Execute(strSqlCmd)
                          Conn.CommitTrans()

                         .Rows.RemoveAt(.CurrentRow.Index)
                          InputData()  'อัพเดทข้อมูลใน datagrid

                       End If

                End If

          End If
          .Focus()
     End With

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub ChangFixPrice()  'update fix_price ทุกครั้งเมื่อลบ ข้อมูล
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlSelc As String
 Dim strSqlcmd As String
 Dim fxprice As Double
 Dim Sumprice As Double

 Dim strFixID As String
 Dim strEqpID As String
 Dim strSize As String


     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

     With dgvFix
          strFixID = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim
          strEqpID = .Rows(.CurrentRow.Index).Cells(3).Value.ToString.Trim
          strSize = .Rows(.CurrentRow.Index).Cells(4).Value.ToString.Trim
          fxprice = .Rows(.CurrentRow.Index).Cells(8).Value

     End With

            strSqlSelc = " SELECT * FROM fixeqpmst" _
                                 & " WHERE fix_id = '" & strFixID & "'"

             Rsd = New ADODB.Recordset
             With Rsd

                     .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                     .LockType = ADODB.LockTypeEnum.adLockOptimistic
                     .Open(strSqlSelc, Conn, , , )

                     If .RecordCount <> 0 Then

                        Sumprice = .Fields("fix_price").Value
                        Sumprice = Sumprice - fxprice

                        strSqlcmd = "UPDATE fixeqpmst SET fix_price = '" & Sumprice & "'" _
                                                  & "WHERE fix_id = '" & strFixID & "'"

                        Conn.Execute(strSqlcmd)
                     End If

             .ActiveConnection = Nothing
             .Close()
             End With

    Conn.Close()
    Conn = Nothing
End Sub

Private Function chkFixData(ByVal txtFixid As String) As Boolean        'เช็คข้อมูลใน fixeqptrn ว่าเหลือ เรคคอร์ดสุดท้ายหรือไม่
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlSelc As String

       With Conn

           If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

              strSqlSelc = "SELECT * FROM fixeqptrn (NOLOCK)" _
                                        & " WHERE fix_id = '" & txtFixid & "'"


              Rsd = New ADODB.Recordset

              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlSelc, Conn, , , )

                        If .RecordCount = 1 Then          'เหลือ record สุดท้าย ก่อนลบ
                            Return True

                        Else
                            Return False

                        End If

             .ActiveConnection = Nothing
             .Close()
             End With

 Conn.Close()
 Conn = Nothing

End Function

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  InputData()
End Sub

Private Sub lblGpbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblGpbClose.Click
  With gpbfxrecv

       .Visible = False
        IsShowSeek = False

  End With

  StateLockFindDept(True)

End Sub

Private Sub StateLockFindDept(ByVal sta As Boolean)
  tabCmd.Enabled = sta
  dgvFix.Enabled = sta
  tlsBarFmr.Enabled = sta
End Sub

Private Sub txtSeek_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSeek.LostFocus
 txtSeek.Text = txtSeek.Text.ToUpper.Trim
End Sub

End Class