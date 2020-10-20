Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmMoldInj

Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean

Dim dubNumberStart As Double
Dim dubNumberEnd As Double

Dim strSqlFindData As String

Dim strDocCode As String = "F0"

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim dsTn As New DataSet

Private Sub frmMoldInj_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

Dim strSearch As String

    '-------------------------ปิด Form เพิ่มแก้ไขเอกสาร--------------------------
    If FormCount("frmAeMoldInj") > 0 Then

            With frmAeMoldInj

               strSearch = .lblComplete.Text

                    If strSearch <> "" Then
                       SearchData(0, strSearch)
                    End If

              .Close()

        End With

    Timer1.Enabled = True
    End If

    Me.Height = Int(lblHeight.Text)
    Me.Width = Int(lblWidth.Text)

    Me.Top = Int(lblTop.Text)
    Me.Left = Int(lblLeft.Text)

End Sub

Public Sub Activating()

   Dim strSearch As String
       strSearch = lblCmd.Text.Trim

       If strSearch <> "" Then
          SearchData(0, strSearch)
       End If

    Timer1.Enabled = False

End Sub

Private Function FormCount(ByVal frmName As String) As Long
Dim frm As Form

    For Each frm In My.Application.OpenForms

        If frm Is My.Forms.frmAeMoldInj Then
           FormCount = FormCount + 1
        End If

    Next

End Function

Private Sub frmMoldInj_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Me.WindowState = FormWindowState.Maximized
     StdDateTimeThai()
     tlsBarFmr.Cursor = Cursors.Hand

     dubNumberStart = 1
     dubNumberEnd = 2100

     PreGroupType()

     InputDeptData()
     tabCmd.Focus()

End Sub

Private Sub frmMoldInj_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
  lblHeight.Text = Me.Height.ToString.Trim
  lblWidth.Text = Me.Width.ToString.Trim
  lblTop.Text = Me.Top.ToString.Trim
  lblLeft.Text = Me.Left.ToString.Trim
End Sub

Private Sub frmMoldInj_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTableUser("print_view_allmold")
   Me.Dispose()
End Sub

Private Sub PreGroupType()

Dim strGpTopic(6) As String
Dim i As Byte

      strGpTopic(0) = "รหัสอุปกรณ์"
      strGpTopic(1) = "รายละเอียดอุปกณ์"
      strGpTopic(2) = "ชิ้นส่วนที่ผลิต"
      strGpTopic(3) = "กลุ่มอุปกรณ์"
      strGpTopic(4) = "รายละเอียดกลุ่ม"
      strGpTopic(5) = "สถานะส่งมอบ"
      strGpTopic(6) = "สถานะส่งซ่อม"

      With cmbType

              For i = 0 To 6
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(0)

      End With

      With cmbFilter

              For i = 0 To 6
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(0)

      End With

End Sub

Private Sub InputDeptData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String = ""
Dim intPageCount As Integer
Dim intPageSize As Integer
Dim intCounter As Integer

Dim strSearch As String = txtFilter.Text.ToString.Trim
Dim strFieldFilter As String = ""

Dim dteComputer As Date = Now()
Dim imgStaPrd As Image
Dim imgStaFix As Image

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
                                     strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 1
                                      strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 2
                                      strFieldFilter = "part_nw like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 3
                                      strFieldFilter = "desc_eng like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 4
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 5
                                      strFieldFilter = "sta_pd like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 6
                                      strFieldFilter = "sta_fx like '%" & ReplaceQuote(strSearch) & "%'"


                    End Select

                    strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                                   & " WHERE " & strFieldFilter _
                                                   & " AND ([group] ='A'" _
                                                   & " OR [group] ='B' OR [group] ='C' )" _
                                                   & " ORDER BY eqp_id"

              Else

                    strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                                  & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                  & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                  & " AND ( [group] ='A'" _
                                                  & " OR [group] ='B' OR [group] ='C' )" _
                                                  & " ORDER BY eqp_id"

              End If

        intPageSize = 30

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

                .PageSize = intPageSize     '.PageSize ใช้กำหนดว่าแต่ละหน้าจะให้มีกี่รายการ ในการแสดงผล
                intPageCount = .PageCount  '.PageCount นับจำนวนหน้าทั้งหมด ที่ได้จากการกำหนดขนาดของหน้า

                '--------------------------ถ้ามีการค้นหา----------------------------------------

                If strSqlFindData <> "" Then

                    .MoveFirst()      'เป็นออบเจ็กต์สำหรับ การเลื่อน Record ไปยัง Record แรกสุด
                    .Find(strSqlFindData)

                    If Not .EOF Then
                        lblPage.Text = Str(.AbsolutePage) '.AbsolutePage ใช้อ้างอิงไปยังหน้าที่ต้องการ
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

                Do While Not .EOF       '.EOF เป็นออบเจ็กต์ตรวจสอบ Pointer ในตำแหน่งสุดท้าย, .BOF เป็นออบเจ็กต์ตรวจสอบ Pointer ในตำแหน่งเริ๋มต้น

                    '-------------------------------------------สถานะส่งมอบฝ่ายผลิต------------------------------------------

                    Select Case .Fields("prod_sta").Value.ToString.Trim

                        Case Is = "1"
                            imgStaPrd = My.Resources._16x16_ledlightblue

                        Case Is = "2"
                            imgStaPrd = My.Resources._16x16_ledgreen

                        Case Is = "0"
                            imgStaPrd = My.Resources._16x16_ledred

                        Case Else
                            imgStaPrd = My.Resources.blank

                    End Select
                    '-------------------------------------------สถานะส่งซ่อม----------------------------------------------

                    Select Case .Fields("fix_sta").Value.ToString.Trim

                        Case Is = "1"
                            imgStaFix = My.Resources.sign_deny
                        Case Is = "2"
                            imgStaFix = My.Resources.Chk
                        Case Else
                            imgStaFix = My.Resources.blank

                    End Select


                    dgvShoe.Rows.Add(
                                         .Fields("eqp_id").Value.ToString.Trim,
                                         .Fields("exp_id").Value.ToString.Trim,
                                         .Fields("eqp_name").Value.ToString.Trim,
                                         .Fields("part_nw").Value.ToString.Trim,
                                         .Fields("desc_eng").Value.ToString.Trim,
                                         .Fields("desc_thai").Value.ToString.Trim,
                                         imgStaPrd, .Fields("sta_pd").Value.ToString.Trim,
                                         .Fields("eqptype").Value.ToString.Trim,
                                         Mid(.Fields("pre_date").Value.ToString, 1, 10),
                                         .Fields("pre_by").Value.ToString.Trim,
                                         Mid(.Fields("last_date").Value.ToString, 1, 10),
                                        .Fields("last_by").Value.ToString.Trim()
                                    )

                    intCounter = intCounter + 1

                    If intCounter = intPageSize Then
                        Exit Do
                    End If

                    .MoveNext()   'เป็นออบเจ็กต์สำหรับ การเลื่อน Record ไป 1 Record
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

         ClearTmpTableUser("tmp_eqptrn")
         lblCmd.Text = "2"

         With frmAeMoldInj
              .ShowDialog()
              .Text = "มุมมองข้อมูล"
         End With

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

'-------------------------------------------------- ซับรูทีนเปลี่่่ยนเพจ--------------------------------------------
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
                    strSqlFind = "eqp_id "
                    strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

               Case Is = 2
                    strSqlFind = "eqp_name "
                    strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

               Case Is = 3
                    strSqlFind = "part_nw "
                    strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

               Case Is = 4
                    strSqlFind = "desc_eng "
                    strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

               Case Is = 5
                    strSqlFind = "desc_thai "
                    strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

               Case Is = 7
                    strSqlFind = "sta_pd "
                    strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

               Case Is = 9
                    strSqlFind = "sta_fx "
                    strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

        End Select


        strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                               & " WHERE " & strSqlFind _
                               & " AND ([group] ='A'" _
                               & " OR [group] ='B' OR [group] ='C' )" _
                               & " ORDER BY eqp_id"

        intPageSize = 30

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

                 '-------------------------------------------------------------------------------------------------------------------

                 If .Fields("RowNumber").Value >= 2100 Then

                     dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                     dubNumberEnd = .Fields("RowNumber").Value + 1000

                 Else

                      dubNumberStart = 1
                      dubNumberEnd = 2100

                 End If

                       strSqlFindData = strSqlFind
                       InputDeptData()

                              For i = 0 To dgvShoe.Rows.Count - 1
                                       If InStr(UCase(dgvShoe.Rows(i).Cells(bytColNumber).Value), strSearchTxt.Trim.ToUpper) <> 0 Then
                                          dgvShoe.CurrentCell = dgvShoe.Item(bytColNumber, i)
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
Dim strSqlCmd As String

Dim btyConsider As Byte
Dim strDept As String
Dim strDeptName As String

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
    End With

    With dgvShoe

        If .Rows.Count > 0 Then

                strDept = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
                strDeptName = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim

                btyConsider = MsgBox("รหัสอุปกรณ์ : " & strDept & vbNewLine _
                                                & "รายละเอียดอุปกรณ์ : " & strDeptName & vbNewLine _
                                                & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then

                        Conn.BeginTrans()

                                '------------------------------------ลบตาราง eqpmst--------------------------------------------

                                strSqlCmd = "DELETE FROM eqpmst" _
                                                      & " WHERE eqp_id ='" & strDept & "'"

                                Conn.Execute(strSqlCmd)

                                '------------------------------------ลบตาราง eqptrn--------------------------------------------

                                strSqlCmd = "DELETE FROM eqptrn" _
                                                     & " WHERE eqp_id ='" & strDept & "'"

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
                                     strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                             Case Is = 1
                                      strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                             Case Is = 2
                                     strFieldFilter = "part_nw like '%" & ReplaceQuote(strSearch) & "%'"

                             Case Is = 3
                                      strFieldFilter = "desc_eng like '%" & ReplaceQuote(strSearch) & "%'"

                             Case Is = 4
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                             Case Is = 5
                                      strFieldFilter = "sta_pd like '%" & ReplaceQuote(strSearch) & "%'"

                             Case Is = 6
                                      strFieldFilter = "sta_fx like '%" & ReplaceQuote(strSearch) & "%'"

                    End Select


                    strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                                  & " WHERE " & strFieldFilter _
                                                  & " AND ([group] ='A'" _
                                                  & " OR [group] ='B' OR [group] ='C' )" _
                                                  & " ORDER BY eqp_id"


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

        ' แปลงเป็นตัวพิมพ์ใหญ่ทันที
        If Char.IsLower(e.KeyChar) Then
            txtFilter.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If

        If e.KeyChar = Chr(13) And txtFilter.Text.Length > 0 Then
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

               Case Is = 0 'รหัสอุปกรณ์
                          SearchData(0, strSearch)

               Case Is = 1 'รายละเอียดอุปกรณ์
                          SearchData(2, strSearch)

               Case Is = 2 'ชิ้นงาน
                          SearchData(3, strSearch)

               Case Is = 3 'กลุ่มอุปกรณ์
                          SearchData(4, strSearch)

               Case Is = 4 'รายละเอียดกลุ่ม
                          SearchData(5, strSearch)

               Case Is = 5 'สถานะส่งฝ่ายผลิต
                          SearchData(7, strSearch)

               Case Is = 6 'สถานะส่งซ่อม
                          SearchData(9, strSearch)

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

        ' แปลงเป็นตัวพิมพ์ใหญ่ทันที
        If Char.IsLower(e.KeyChar) Then
            txtSeek.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If

        If e.KeyChar = Chr(13) And txtSeek.Text.Length > 0 Then
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

                         blnReturn = CheckUserEntry(strDocCode, "act_add")    'strDocCode = 'F0', act_add คือสิทธิ์การ เพิ่มข้อมูล
                         If blnReturn Then

                            ClearTmpTableUser("tmp_eqptrn")
                            lblCmd.Text = "0"
                            With frmAeMoldInj
                                 .ShowDialog()
                                 .Text = "เพิ่มข้อมูล"
                            End With

                             Else
                                MsnAdmin()
                            End If


                  Case Is = 1 'แก้ไขข้อมูล

                            blnReturn = CheckUserEntry(strDocCode, "act_edit")  'strDocCode = 'F0', act_edit คือสิทธิ์การแก้ไขข้อมูล
                            If blnReturn Then

                                    If dgvShoe.Rows.Count > 0 Then

                                       ClearTmpTableUser("tmp_eqptrn")
                                       lblCmd.Text = "1"
                                       With frmAeMoldInj
                                            .ShowDialog()
                                            .Text = "แก้ไขข้อมูล"
                                       End With

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

                                     .Top = 80
                                     .Left = 120
                                     .Width = 348
                                     .Height = 125
                                     .Visible = True

                                     cmbFilter.SelectedItem = cmbFilter.Items(0)
                                     txtFilter.Text = _
                                     dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim()

                                     StateLockFind(False)
                                     txtFilter.Focus()

                                 End With

                            End If

                  Case Is = 4 'ค้นหาข้อมูล

                            If dgvShoe.Rows.Count > 0 Then

                                 With gpbSearch
                                      .Top = 80
                                      .Left = 120
                                      .Width = 348
                                      .Height = 125

                                      .BringToFront()
                                      .Visible = True

                                      cmbType.SelectedItem = cmbType.Items(0)
                                      txtSeek.Text = _
                                      dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim()

                                      StateLockFind(False)
                                      txtSeek.Focus()
                                 End With

                            End If

                  Case Is = 5 'พิมพ์ข้อมูล

                             If dgvShoe.Rows.Count > 0 Then
                                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                                ClearTmpTableUser("tmp_eqpmst")

                                 With frmPMoldinj
                                      .ShowDialog()
                                 End With
                                 Me.Cursor = System.Windows.Forms.Cursors.Arrow
                            End If

                  Case Is = 6     'พิมพ์ All Molds

                            If dgvShoe.Rows.Count > 0 Then
                               Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                               Me.BackgroundWorker1.RunWorkerAsync()
                                'frmWaiting.ShowDialog()

                               ClearTmpTableUser("tmp_eqpmst")
                               frmPrntProgress.ShowDialog()

                               frmMainPro.lblRptCentral.Text = "I"
                               frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim
                               frmRptCentral.ShowDialog()
                               Me.Cursor = System.Windows.Forms.Cursors.Arrow
                            End If


                Case Is = 7 'ฟื้นฟูข้อมูล
                    blnHaveFilter = False
                    InputDeptData()

                  Case Is = 8 'ลบข้อมูล

                          blnReturn = CheckUserEntry(strDocCode, "act_delete")
                          If blnReturn Then
                             DeleteData()
                          Else
                                MsnAdmin()
                          End If

                  Case Is = 9 'ออก
                            Me.Close()

         End Select

    End With

End Sub

Private Sub dgvShoe_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvShoe.RowsAdded
  dgvShoe.Rows(e.RowIndex).Height = 27
End Sub

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
   InputDeptData()
End Sub

Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
   For n As Integer = 1 To 50
        Threading.Thread.Sleep(10)
   Next
End Sub

Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

  With frmWaiting
        .Countdown = 7
        .ShowDialog(Me)
  End With
  ' Application.Exit()
End Sub

End Class