Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmEqpSheet

Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean    'กรณีกรองข้อมูล

Dim dubNumberStart As Double   'ถูกกำหนด = 1
Dim dubNumberEnd As Double     'ถูกกำหนด = 2100

Dim strSqlFindData As String
Dim strDocCode As String = "F0"

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim dsTn As New DataSet

Private Sub frmEqpSheet_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

Dim strSearch As String

    '-------------------------ปิด Form เพิ่มแก้ไขเอกสาร--------------------------

    If FormCount("frmPastMold") > 0 Then

        With frmPastMold

               strSearch = .lblComplete.Text

                If strSearch <> "" Then
                    SearchData(0, strSearch)
                End If

              .Close()

        End With

     Timer1.Enabled = True       'ให้ Timer1 รีเฟรชหน้าจอ
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

            If frm Is My.Forms.frmPastMold Then
                FormCount = FormCount + 1
            End If
    Next

End Function

Private Sub frmEqpSheet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Me.WindowState = FormWindowState.Maximized     'ขยายขนาดเต็มหน้าจอ
     StdDateTimeThai()                           'เรียก ซับรูทีน StdDateTimeThai
     tlsBarFmr.Cursor = Cursors.Hand             'ให้คอร์เซอร์ตรง Toolstripbar เป็นรูปมือ

     dubNumberStart = 1                          'ให้แถวเเรกใน Recordset = 1
     dubNumberEnd = 2100                         'ให้แถวเเรกใน Recordset = 2100

     PreGroupType()
     InputDeptData()
     tabCmd.Focus()
End Sub

     'เมื่อเรียกฟอร์มอื่นขึ้นมาแทน
Private Sub frmEqpSheet_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
    lblHeight.Text = Me.Height.ToString.Trim
    lblWidth.Text = Me.Width.ToString.Trim

    lblTop.Text = Me.Top.ToString.Trim
    lblLeft.Text = Me.Left.ToString.Trim
End Sub

Private Sub frmEqpSheet_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTableUser("print_view_allmold")
   Me.Dispose()
End Sub

'------------------------------------- โหลดข้อมูลใส่ cbo Groupboxค้นหา,กรองข้อมูล -----------------------------------------
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
		Dim strDateAdd As String = ""
		Dim strDateEdit As String = ""

		Dim strInDate As String = ""

		Dim intPageCount As Integer
		Dim intPageSize As Integer
		Dim intCounter As Integer

		Dim strSearch As String = txtFilter.Text.ToString.Trim
		Dim strFieldFilter As String = ""

		Dim dteComputer As Date = Now()
		Dim imgStaPrd As Image
		Dim imgStaFix As Image

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

                  strSqlCmdSelc = "SELECT  * FROM v_moldinj_hd (NOLOCK)" _
                                                   & " WHERE " & strFieldFilter _
                                                   & " AND [group] = 'D'" _
                                                   & " ORDER BY eqp_id"
        Else


              strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                          & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                          & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                          & " AND [group] = 'D'" _
                                          & " ORDER BY eqp_id"

              End If

              intPageSize = 30   'ตัวแปรกำหนดขนาดกระดาษ

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

                                    '------------------------------------------------------------------------------------

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

                                    '--------------------------  สถานะส่งมอบฝ่ายผลิต   ---------------------------------------

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

                                            '--------------------------------  สถานะส่งซ่อม  ---------------------------------------

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
                                          .Fields("eqp_name").Value.ToString.Trim,
                                          .Fields("part_nw").Value.ToString.Trim,
                                          .Fields("tech_desc").Value.ToString.Trim,
                                          imgStaPrd,
                                          .Fields("sta_pd").Value.ToString.Trim,
                                          Mid(.Fields("pre_date").Value.ToString, 1, 10),
                                          .Fields("pre_by").Value.ToString.Trim,
                                           Mid(.Fields("last_date").Value.ToString, 1, 10),
                                          .Fields("last_by").Value.ToString.Trim
                                    )

                    'CDate(.Fields("pre_date").Value).ToString("dd-MM-yyyy"),
                    'IIf(IsDate(.Fields("last_date").Value), CDate(.Fields("last_date").Value).ToString("dd-MM-yyyy"), ""),

                    intCounter = intCounter + 1

                    If intCounter = intPageSize Then
                        Exit Do
                    End If

                    .MoveNext()    'ข้ามไปที่ระเบียนใหม่
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

       With frmPastMold
            .ShowDialog()
            .Text = "มุมมองข้อมูล"

       End With

       'Me.Hide()
       'frmMainPro.Hide()
    End If

End Sub

Public Sub StateLockFind(ByVal Sta As Boolean)
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

Private Sub SearchData()
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

                        Select Case cmbType.SelectedIndex()

                               Case Is = 0
                                     strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 1
                                      strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 2
                                      strFieldFilter = "part_nw like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 3
                                      strFieldFilter = "desc_eng like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 4
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 5
                                      strFieldFilter = "sta_pd like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 6
                                      strFieldFilter = "sta_fx like '%" & ReplaceQuote(strSearch) & "'"

                      End Select

                      strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                                  & " WHERE " & strFieldFilter _
                                                  & " AND ([group] ='D'" _
                                                  & " ORDER BY eqp_id"




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
Dim strEqpid As String
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

            strEqpid = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
            strDeptName = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim

            btyConsider = MsgBox("รหัสอุปกรณ์ : " & strEqpid & vbNewLine _
                                                & "รายละเอียดอุปกรณ์ : " & strDeptName & vbNewLine _
                                                & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

           If btyConsider = 6 Then

              Conn.BeginTrans()

              '------------------------------------ลบตาราง eqpmst--------------------------------------------

              strSqlCmd = "DELETE FROM eqpmst" _
                                    & " WHERE eqp_id ='" & strEqpid & "'"

              Conn.Execute(strSqlCmd)

              '------------------------------------ลบตาราง eqptrn--------------------------------------------

              strSqlCmd = "DELETE FROM eqptrn" _
                                       & " WHERE eqp_id ='" & strEqpid & "'"

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

Sub FilterData()    'กรองข้อมูล

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
                                                  & " ORDER BY eqp_id"


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

    '--------------------------------------------- ค้นหาเอกสาร Groupbox ค้นหาข้อมูล -----------------------------------------------------
    Private Sub FindDocID()

        Dim strSearch As String = txtSeek.Text.ToUpper.Trim

        If strSearch <> "" Then

            Select Case cmbType.SelectedIndex()

                Case Is = 0 'รหัสอุปกรณ์
                    SearchData(0, strSearch)     'ส่งตำเงื่อนไข ,Text ให้ ซับรูทีน SearchData

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

            MsgBox("โปรดกรอกข้อมูลเพื่อค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
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

                            blnReturn = CheckUserEntry(strDocCode, "act_add")
                            If blnReturn Then

                               ClearTmpTableUser("tmp_eqptrn")
                               lblCmd.Text = "0"
                               With frmPastMold
                                    .ShowDialog()
                                    .Text = "เพิ่มข้อมูล"
                               End With

                            Else
                                MsnAdmin()
                            End If


                    Case Is = 1     'แก้ไขข้อมูล

                            blnReturn = CheckUserEntry(strDocCode, "act_edit")
                            If blnReturn Then

                                    If dgvShoe.Rows.Count > 0 Then

                                       ClearTmpTableUser("tmp_eqptrn")
                                       lblCmd.Text = "1"

                                       With frmPastMold
                                            .ShowDialog()
                                            .Text = "แก้ไขข้อมูล"
                                       End With

                                    End If

                            Else
                                MsnAdmin()
                            End If

                    Case Is = 2    'มุมมอง

                            blnReturn = CheckUserEntry(strDocCode, "act_view")
                            If blnReturn Then
                                 ViewShoeData()
                            Else
                                MsnAdmin()
                            End If

                    Case Is = 3    'กรองข้อมูล

                            If dgvShoe.Rows.Count > 0 Then

                                With gpbFilter

                                            .Top = 230
                                            .Left = 210
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

                                            .Top = 230
                                            .Left = 210
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
                                ClearTmpTableUser("tmp_eqpmst")

                                  With Me.gpbOptPrint

                                       .Top = 300
                                       .Left = 250
                                       .Width = 374
                                       .Height = 125
                                       .Visible = True

                                       Me.InputEqpDataPrint()
                                       Me.cmbOptPrint.Text = Me.dgvShoe.Rows(Me.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim()

                                       Me.StateLockFind(False)
                                       Me.cmbOptPrint.Focus()

                                End With

                            End If

                    Case Is = 6     'พิมพ์ mold ทั้งหมด

                               If dgvShoe.Rows.Count > 0 Then
                                  Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                                  Me.BackgroundWorker1.RunWorkerAsync()
                                  ClearTmpTableUser("tmp_eqpmst")
                                  PrepareData()    'โหลดรายการโมล์ดทั้งหมดเรียง size ใหม่

                                  frmMainPro.lblRptCentral.Text = "J"
                                  frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim
                                  frmRptCentral.ShowDialog()
                                  Me.Cursor = System.Windows.Forms.Cursors.Hand
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

'------------------------------------- ซับรูทีนค้นหาข้อมูล (รับค่าเงื่อนไข,ข้อความที่ต้องการค้น) -------------------------------------------------------

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
                                          & " AND [group]= 'D'" _
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

'----------------------- ฟังก์ชั่นปรับขนาด Size รูปภาพ -------------------------------------------------------------------

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

Public Sub InputEqpDataPrint()

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

        strSqlSelc = "SELECT DISTINCT eqp_id FROM v_mst_trn (NOLOCK)" _
                                           & " WHERE [group] ='D'" _
                                           & " ORDER BY eqp_id"


       RsdPnt = New ADODB.Recordset

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
   Conn.Close()

End Sub

Private Sub btnPrntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntCancel.Click
  StateLockFind(True)
  gpbOptPrint.Visible = False

End Sub

Private Sub btnPrntPrevw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntPrevw.Click
  Dim strDocId As String = cmbOptPrint.Text.ToUpper.Trim

  If strDocId <> "" Then
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        PrePrintData(strDocId)
        frmMainPro.lblRptCentral.Text = "C"          ' ส่งค่าให้ตัวเเปรฟอร์ม MainPro 

        '-------------------------ส่งค่าให้ตัวแปร lblRptDesc ของฟอร์ม MainPro โดยส่ง Userid กับ Eqpid ----------------------------- 

        frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim _
                                                                & "' AND eqp_id ='" & strDocId & "'"

        frmRptCentral.Show()
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
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

     strSqlSelc = "SELECT * FROM v_mst_trn (NOLOCK)" _
                        & " WHERE eqp_id = '" & strSelectCode.ToString.Trim & "'"

     Rsd = New ADODB.Recordset
     With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                  For i As Integer = 1 To .RecordCount

                      '----------------------------------------LoadPicture รูปแผง ------------------------------------------------

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

                     '----------------------------------------LoadPicture รูป เริ่มชิ้นงาน ---------------------------------------------

                     strLoadFilePic2 = strPicPath & .Fields("pic_io").Value.ToString.Trim

                     If strLoadFilePic2 <> "" Then

                        If File.Exists(strLoadFilePic2) Then 'รูปยังมีอยู่ในระบบ
                           blnHavePic2 = True
                        Else
                           blnHavePic2 = False
                        End If

                    Else
                        blnHavePic3 = False
                    End If

                    '----------------------------------------LoadPicture รูปรองเท้าสำเร็จรูป --------------------------------------------

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

                       strSqlCmdPic = "SELECT * " _
                                            & " FROM tmp_mst_trn (NOLOCK)"

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
                                                    RsdPic.Fields("ap_desc").Value = .Fields("ap_desc").Value
                                                    RsdPic.Fields("doc_ref").Value = .Fields("doc_ref").Value
                                                    RsdPic.Fields("set_qty").Value = .Fields("set_qty").Value
                                                    RsdPic.Fields("pic_ctain").Value = .Fields("pic_ctain").Value
                                                    RsdPic.Fields("pic_io").Value = .Fields("pic_io").Value
                                                    RsdPic.Fields("pic_part").Value = .Fields("pic_part").Value
                                                    RsdPic.Fields("remark").Value = .Fields("remark").Value
                                                    RsdPic.Fields("tech_desc").Value = .Fields("tech_desc").Value
                                                    RsdPic.Fields("tech_thk").Value = .Fields("tech_thk").Value
                                                    RsdPic.Fields("tech_trait").Value = .Fields("tech_trait").Value
                                                    RsdPic.Fields("tech_sht").Value = .Fields("tech_sht").Value
                                                    RsdPic.Fields("tech_eva").Value = .Fields("tech_eva").Value
                                                    RsdPic.Fields("tech_warm").Value = .Fields("tech_warm").Value
                                                    RsdPic.Fields("tech_time1").Value = .Fields("tech_time1").Value
                                                    RsdPic.Fields("tech_time2").Value = .Fields("tech_time2").Value
                                                    RsdPic.Fields("creat_date").Value = .Fields("creat_date").Value
                                                    RsdPic.Fields("size_desc").Value = .Fields("size_desc").Value
                                                    RsdPic.Fields("size_id").Value = .Fields("size_id").Value
                                                    RsdPic.Fields("size_qty").Value = .Fields("size_qty").Value
                                                    RsdPic.Fields("dimns").Value = .Fields("dimns").Value
                                                    RsdPic.Fields("men_rmk").Value = .Fields("men_rmk").Value
                                                    RsdPic.Fields("pre_date").Value = .Fields("pre_date").Value
                                                    RsdPic.Fields("pre_by").Value = .Fields("pre_by").Value
                                                    RsdPic.Fields("last_date").Value = .Fields("last_date").Value
                                                    RsdPic.Fields("last_by").Value = .Fields("last_by").Value
                                                    RsdPic.Fields("pi_qty").Value = .Fields("pi_qty").Value
                                                    RsdPic.Fields("eqp_amt").Value = .Fields("eqp_amt").Value

                                                    '----------------------------เพิ่มข้อมูลรูปภาพแผง---------------------------------------------------

                                                    If blnHavePic1 Then              'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                            Dim RsdSteam1 As New MemoryStream
                                                            Dim bytes1 = File.ReadAllBytes(strLoadFilePic1)

                                                            inImg = Image.FromFile(strLoadFilePic1)  'ดึงรูปขึ้นมา
                                                            'inImg = SizeImage(inImg, 230, 200)     'ปรับขนาด size
                                                            inImg = ScaleImage(inImg, 167, 212)
                                                            inImg.Save(RsdSteam1, ImageFormat.Bmp)  'เปลี่ยนนามสกุล .Bmp
                                                            bytes1 = RsdSteam1.ToArray
                                                            RsdPic.Fields("bob_ctain").Value = bytes1   'พาทค่าไปไว้ที่่ฟิวด์ bob_ctain

                                                            RsdSteam1.Close()              'ปิด RecordSet
                                                            RsdSteam1 = Nothing            'เคลียร์ RecordSet

                                                    End If

                                                    '----------------------------เพิ่มข้อมูลรูปภาพเริ่มชิ้นงาน ----------------------------------------------

                                                    If blnHavePic2 Then            'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                            Dim RsdSteam2 As New MemoryStream
                                                            Dim bytes2 = File.ReadAllBytes(strLoadFilePic2)

                                                            inImg = Image.FromFile(strLoadFilePic2)
                                                            'inImg = SizeImage(inImg, 230, 200)
                                                            inImg = ScaleImage(inImg, 167, 212)
                                                            inImg.Save(RsdSteam2, ImageFormat.Bmp)
                                                            bytes2 = RsdSteam2.ToArray
                                                            RsdPic.Fields("bob_io").Value = bytes2

                                                            RsdSteam2.Close()
                                                            RsdSteam2 = Nothing

                                                    End If

                                                    '----------------------------เพิ่มข้อมูลรูปภาพชิ้นงาน-------------------------------------------------

                                                    If blnHavePic3 Then                  'ถ้ามีรูปภาพให้แปลงเป็น Binary แล้วเพิ่มข้อมูลเข้าไป

                                                            Dim RsdSteam3 As New MemoryStream
                                                            Dim bytes3 = File.ReadAllBytes(strLoadFilePic3)

                                                            inImg = Image.FromFile(strLoadFilePic3)
                                                            'inImg = SizeImage(inImg, 230, 200)
                                                            inImg = ScaleImage(inImg, 167, 212)
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
                                  .MoveNext()     'เลื่อนไปที่ Record ถัดไป
                  Next i

                End If

            .ActiveConnection = Nothing
            .Close()

    End With
    Rsd = Nothing

Conn.Close()
Conn = Nothing

End Sub

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
   InputDeptData()   'สั่งฟื้นฟูข้อมูล
End Sub

Private Sub PrepareData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim sql As String
  Dim sqlCmd As String

  Dim strArr() As String
  Dim SearchWithinThis As String
  Dim newSize As String

      Try

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      sql = "SELECT * FROM view_mold " _
                     & " WHERE [group] ='D'" _
                     & " ORDER BY eqp_id"

      With Rsd

           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .CursorType = ADODB.CursorLocationEnum.adUseClient
           .Open(sql, Conn, , , )

           If .RecordCount <> 0 Then

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
                           & "(user_id,eqp_id,eqp_name,desc_thai,size_id,size_desc" _
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

                  .MoveNext()
               Loop

           End If

          .ActiveConnection = Nothing
          .Close()
      End With

   Conn.Close()

      Catch ex As Exception
           MsgBox(ex.Message)
      End Try
End Sub

Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

  With frmWaiting
        .Countdown = 7
        .ShowDialog(Me)
  End With
  ' Application.Exit()
End Sub

End Class