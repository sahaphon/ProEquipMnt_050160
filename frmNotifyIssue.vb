Imports ADODB
Imports System.IO
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmNotifyIssue
 Dim intBkPageCount As Integer
 Dim blnHaveFilter As Boolean    'กรณีกรองข้อมูล

 Dim dubNumberStart As Double   'ถูกกำหนด = 1
 Dim dubNumberEnd As Double     'ถูกกำหนด = 2100

 Dim strSqlFindData As String
 Dim strDocCode As String = "F10"

 Dim da As New System.Data.OleDb.OleDbDataAdapter
 Dim ds As New DataSet
 Dim dsTn As New DataSet

Private Sub frmNotifyIssue_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

    If FormCount("frmAeNotifyIssue") > 0 Then

       With frmAeNotifyIssue

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

 InputDeptData()

End Sub

Private Sub frmNotifyIssue_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Me.WindowState = FormWindowState.Maximized     'ขยายขนาดเต็มหน้าจอ
  StdDateTimeThai()                           'เรียก ซับรูทีน StdDateTimeThai
  tlsBarFmr.Cursor = Cursors.Hand             'ให้คอร์เซอร์ตรง Toolstripbar เป็นรูปมือ

  dubNumberStart = 1                          'ให้แถวเเรกใน Recordset = 1
  dubNumberEnd = 2100                         'ให้แถวเเรกใน Recordset = 2100

  PreGroupType()
  InputDeptData()
  tabCmd.Focus()

End Sub

Private Sub frmNotifyIssue_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
 lblHeight.Text = Me.Height.ToString.Trim
 lblWidth.Text = Me.Width.ToString.Trim

 lblTop.Text = Me.Top.ToString.Trim
 lblLeft.Text = Me.Left.ToString.Trim

End Sub

Private Sub frmNotifyIssue_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 Me.Dispose()

End Sub

Private Function FormCount(ByVal frmName As String) As Long

 Dim frm As Form

     For Each frm In My.Application.OpenForms

         If frm Is My.Forms.frmAeNotifyIssue Then
            FormCount = FormCount + 1
         End If

     Next

End Function

Private Sub PreGroupType()

 Dim strGpTopic(5) As String
 Dim i As Byte

     strGpTopic(0) = "รหัส"
     strGpTopic(1) = "สถานะ"
     strGpTopic(2) = "รายละเอียดอุปกณ์"           'ชื่ออุปกรณ์
     strGpTopic(3) = "แผนกที่แจ้งปัญหา"
     strGpTopic(4) = "รายละเอียดกลุ่ม"            'DescThai
     strGpTopic(5) = "รุ่นอุปกรณ์"                'รุุ่นอุปกรณ์

     With cmbType

         For i = 0 To 5
             .Items.Add(strGpTopic(i))
         Next i

         .SelectedItem = .Items(0)

     End With

         With cmbFilter

              For i = 0 To 5
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(0)

         End With

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
 Dim numb As Integer


        With Conn
              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

        End With

                   Select Case bytColNumber

                          Case Is = 0
                                 strSqlFind = "req_id "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 1
                                 'strSqlFind = "sta_notify "
                                 'strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"
                                      If frmMainPro.lblLogin.Text = "SUTID" Then

                                             If strSearchTxt = "อยู่ระหว่างรับเเจ้ง" Then
                                                strSqlFind = "req_sta = '" & "2" & "'"

                                             ElseIf strSearchTxt = "รอรับเเจ้ง" Then
                                                strSqlFind = "req_sta = '" & "1" & "'"

                                             Else
                                                strSqlFind = "req_sta = '" & "3" & "'"

                                             End If

                                      Else

                                             If strSearchTxt = "อยู่ระหว่างรับเเจ้ง" Then
                                                strSqlFind = "req_sta = '" & "1" Or "2" & "'"

                                             ElseIf strSearchTxt = "รอรับเเจ้ง" Then
                                                strSqlFind = "req_sta = '" & "0" & "'"

                                             Else
                                                strSqlFind = "req_sta = '" & "3" & "'"

                                             End If

                                          'strFieldFilter = "sta_notify like '" & ReplaceQuote(strSearch) & "%'"
                                      End If

                          Case Is = 2
                                 strSqlFind = "eqp_nm "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 3
                                 strSqlFind = "from_dep "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 4
                                 strSqlFind = "desc_thai "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 5
                                  strSqlFind = "series "
                                  strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                    End Select


                       strSqlCmdSelc = "SELECT * FROM v_notifyissue (NOLOCK)" _
                                               & " WHERE " & strSqlFind _
                                               & " ORDER BY req_id"



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
                                        '-----------------------------------------------------------------------------------------------------------

                                        If .Fields("RowNumber").Value >= 2100 Then

                                            dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                                            dubNumberEnd = .Fields("RowNumber").Value + 1000

                                        Else

                                             dubNumberStart = 1
                                             dubNumberEnd = 2100

                                        End If

                                        strSqlFindData = strSqlFind

                                        InputDeptData()

                                               If bytColNumber = 0 Then
                                                  numb = 2

                                                   ElseIf bytColNumber = 1 Then
                                                         numb = 1

                                                   ElseIf bytColNumber = 2 Then
                                                          numb = 4

                                                   ElseIf bytColNumber = 3 Then
                                                          numb = 3

                                               End If

                                                For i = 0 To dgvIssue.Rows.Count - 1
                                                        'เปลี่ยน  2 เป็น bytColNumber

                                                        If InStr(UCase(dgvIssue.Rows(i).Cells(numb).Value.ToString), strSearchTxt.Trim.ToUpper) <> 0 Then
                                                                dgvIssue.CurrentCell = dgvIssue.Item(numb, i)
                                                                dgvIssue.Focus()
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
 Dim imgStaReq As Image

 Dim strDateFilter As String = ""
 Dim strYearCnvt As String = ""
 Dim strSta As String

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
                                      strFieldFilter = "req_id like '%" & ReplaceQuote(strSearch) & "%'"

                                   Case Is = 1

                                      If frmMainPro.lblLogin.Text = "SUTID" Then

                                             If strSearch = "อยู่ระหว่างรับเเจ้ง" Then
                                                strFieldFilter = "req_sta = '" & "2" & "'"

                                             ElseIf strSearch = "รอรับเเจ้ง" Then
                                                strFieldFilter = "req_sta = '" & "1" & "'"

                                             Else
                                                strFieldFilter = "req_sta = '" & "3" & "'"

                                             End If

                                      Else

                                             If strSearch = "อยู่ระหว่างรับเเจ้ง" Then
                                                strFieldFilter = "req_sta = '" & "1" Or "2" & "'"

                                             ElseIf strSearch = "รอรับเเจ้ง" Then
                                                strFieldFilter = "req_sta = '" & "0" & "'"

                                             Else
                                                strFieldFilter = "req_sta = '" & "3" & "'"

                                             End If

                                          'strFieldFilter = "sta_notify like '" & ReplaceQuote(strSearch) & "%'"
                                      End If

                                   Case Is = 2
                                      strFieldFilter = "eqpnm like '%" & ReplaceQuote(strSearch) & "%'"

                                   Case Is = 3
                                      strFieldFilter = "from_dep like '%" & ReplaceQuote(strSearch) & "%'"

                                   Case Is = 4
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                                   Case Is = 5
                                      strFieldFilter = "shoe like '%" & ReplaceQuote(strSearch) & "%'"

                              End Select

                                          If frmMainPro.lblLogin.Text = "SUTID" Then
                                             strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                  & " WHERE " & strFieldFilter _
                                                                  & " ORDER BY req_id"

                                          Else
                                             strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                  & " WHERE " & strFieldFilter _
                                                                  & " ORDER BY req_id"

                                          End If


              Else

                   If frmMainPro.lblLogin.Text = "SUTID" Then

                      strSqlCmdSelc = "SELECT * FROM v_notifyissue (NOLOCK)" _
                                                 & " WHERE RowNumber >= " & dubNumberStart.ToString.Trim _
                                                 & " AND RowNumber <= " & dubNumberEnd.ToString.Trim _
                                                 & " AND person2_sta = 'True'" _
                                                 & " ORDER BY req_id"

                   Else

                      strSqlCmdSelc = "SELECT * FROM v_notifyissue (NOLOCK)" _
                                          & " WHERE RowNumber >= " & dubNumberStart.ToString.Trim _
                                          & " AND RowNumber <= " & dubNumberEnd.ToString.Trim _
                                          & " ORDER BY req_id"

                   End If

                  
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

                                    dgvIssue.Rows.Clear()

                                    intCounter = 0

                                    Do While Not .EOF

                                     '------------------------------------------- สถานะดำเนินการเอกสาร ---------------------------
                                        If frmMainPro.lblLogin.Text = "SUTID" Then

                                           Select Case .Fields("req_sta").Value.ToString.Trim

                                                  Case Is = "0"
                                                                    imgStaReq = My.Resources._16x16_ledred
                                                                    strSta = "รอรับเเจ้ง"

                                                  Case Is = "1"
                                                                    imgStaReq = My.Resources._16x16_ledred
                                                                     strSta = "รอรับเเจ้ง"

                                                  Case Is = "2"
                                                                    imgStaReq = My.Resources._16x16ledyellow
                                                                    strSta = "อยู่ระหว่างรับเเจ้ง"

                                                  Case Else  'กรณีเป็น 3
                                                                    imgStaReq = My.Resources._16x16_ledgreen
                                                                    strSta = "รับเเจ้งแล้ว"

                                           End Select


                                        Else
                                                Select Case .Fields("req_sta").Value.ToString.Trim

                                                          Case Is = "0"
                                                                    imgStaReq = My.Resources._16x16_ledred
                                                                    strSta = "รอรับเเจ้ง"

                                                          Case Is = "1"
                                                                    imgStaReq = My.Resources._16x16ledyellow
                                                                    strSta = "อยู่ระหว่างรับเเจ้ง"

                                                          Case Is = "2"
                                                                    imgStaReq = My.Resources._16x16ledyellow
                                                                    strSta = "อยู่ระหว่างรับเเจ้ง"

                                                          Case Else   'กรณีเป็น 3
                                                                    imgStaReq = My.Resources._16x16_ledgreen
                                                                    strSta = "รับเเจ้งแล้ว"

                                                 End Select

                                        End If

                                            dgvIssue.Rows.Add( _
                                                                  imgStaReq, strSta, _
                                                                  .Fields("req_id").Value.ToString.Trim, _
                                                                  .Fields("dep_notify").Value.ToString.Trim, _
                                                                  .Fields("eqpnm").Value.ToString.Trim, _
                                                                  .Fields("shoe").Value.ToString.Trim & " / " & .Fields("size").Value.ToString.Trim, _
                                                                  .Fields("amount").Value, _
                                                                  .Fields("issue").Value.ToString.Trim & " / " & .Fields("cause").Value.ToString.Trim, _
                                                                  .Fields("fxissue").Value.ToString.Trim, _
                                                                  Mid(.Fields("person1_date").Value.ToString.Trim, 1, 10), _
                                                                  .Fields("person1").Value.ToString.Trim, _
                                                                  Mid(.Fields("person3_date").Value.ToString.Trim, 1, 10), _
                                                                  .Fields("person3").Value.ToString.Trim, _
                                                                  Mid(.Fields("last_date").Value.ToString.Trim, 1, 10), _
                                                                  .Fields("lastby").Value.ToString.Trim, _
                                                                  .Fields("remark").Value.ToString.Trim _
                                                            )
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

Private Sub StateLockFind(ByVal Sta As Boolean)
 tabCmd.Enabled = Sta
 dgvIssue.Enabled = Sta
 tlsBarFmr.Enabled = Sta

End Sub

Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
 lblPage.Text = "1"
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

Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
 lblPage.Text = Str(intBkPageCount)
 InputDeptData()
End Sub

Private Sub txtPage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPage.GotFocus
 txtPage.SelectAll()
End Sub

Private Sub txtPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPage.KeyPress
 If e.KeyChar = Chr(13) Then
     dgvIssue.Focus()
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

Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

 Dim blnReturn As Boolean
 Dim strEqpId As String = ""

     With tabCmd

          Select Case .SelectedIndex

                    Case Is = 0 'เพิ่มข้อมูล

                            blnReturn = CheckUserEntry(strDocCode, "act_add")
                            If blnReturn Then

                               ClearTmpTableUser("tmp_notifyissue")
                               lblCmd.Text = "0"

                               With frmAeNotifyIssue
                                    .Show()
                                    .Text = "เพิ่มข้อมูล"

                               End With

                               Me.Hide()
                               frmMainPro.Hide()

                            Else
                                MsnAdmin()
                            End If

                    Case Is = 1   'แก้ไขข้อมูล

                            blnReturn = CheckUserEntry(strDocCode, "act_edit")
                            If blnReturn Then

                                    If dgvIssue.Rows.Count > 0 Then

                                        If ChkCompleteapprove() Then     'เช็คว่า approve ครบทุกคนแล้วหรือยัง (ถ้าครบทุกคนแล้ว ไม่ให้แก้ไข)

                                           ClearTmpTableUser("tmp_notifyissue")
                                           lblCmd.Text = "1"

                                              With frmAeNotifyIssue
                                                   .Show()
                                                   .Text = "แก้ไขข้อมูล"
                                              End With

                                              Me.Hide()
                                              frmMainPro.Hide()

                                        Else
                                           MessageBox.Show("ไม่สามารถดำเนินการได้  เอกสารเซ็นยืนยันครบถ้วนแล้ว", _
                                                                         "Access denied!.....", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        End If

                                     End If

                            Else
                                MsnAdmin()
                            End If

                    Case Is = 2    'มุมมอง

                            blnReturn = CheckUserEntry(strDocCode, "act_view")
                            If blnReturn Then
                               ViewData()
                            Else
                                MsnAdmin()
                            End If

                    Case Is = 3    'กรองข้อมูล

                            If dgvIssue.Rows.Count > 0 Then

                                With gpbFilter

                                     .Top = 230
                                     .Left = 210
                                     Width = 348
                                     .Height = 125

                                     .Visible = True

                                     cmbFilter.SelectedItem = cmbFilter.Items(0)
                                     txtFilter.Text = _
                                     dgvIssue.Rows(dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                                     StateLockFind(False)
                                     txtFilter.Focus()

                                 End With

                            End If

                    Case Is = 4 'ค้นหาข้อมูล

                            If dgvIssue.Rows.Count > 0 Then

                               With gpbSearch

                                    .Top = 230
                                    .Left = 210
                                    .Width = 348
                                    .Height = 125

                                    .BringToFront()
                                    .Visible = True

                                    cmbType.SelectedItem = cmbType.Items(0)
                                    txtSeek.Text = _
                                             dgvIssue.Rows(dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                                    StateLockFind(False)
                                    txtSeek.Focus()

                               End With

                            End If

                    Case Is = 5 'ฟื้นฟูข้อมูล
                              InputDeptData()

                    Case Is = 6 'พิมพ์ข้อมูล

                             If dgvIssue.Rows.Count > 0 Then
                                ClearTmpTableUser("tmp_notifyissue")

                                With gpbOptPrint

                                     .Top = 230
                                     .Left = 210
                                     .Width = 374
                                     .Height = 125

                                     .Visible = True

                                     InputEqpDataPrint()

                                     cmbOptPrint.Text = dgvIssue.Rows(dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                                     StateLockFind(False)
                                     cmbOptPrint.Focus()

                                End With

                            End If

                    Case Is = 7 'ลบข้อมูล

                          blnReturn = CheckUserEntry(strDocCode, "act_delete")

                            If blnReturn Then

                               If dgvIssue.Rows.Count <> 0 Then
                                  DeleteData()

                               End If

                            Else
                                MsnAdmin()
                            End If

                    Case Is = 8 'ออก
                            Me.Close()

          End Select

  End With
End Sub

Private Function ChkCompleteapprove() As Boolean
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlSelc As String
 Dim strReqid As String
     strReqid = dgvIssue.Rows(dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim()

     With Conn

          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

        strSqlSelc = "SELECT * FROM v_notifyissue (NOLOCK)" _
                                   & " WHERE req_id = '" & strReqid & "' "


       Rsd = New ADODB.Recordset

       With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            If .RecordCount <> 0 Then

                If .Fields("req_sta").Value.ToString.Trim = "2" Then

                   Return False

                Else

                   Return True

                End If


            Else
                Return False

            End If

         .ActiveConnection = Nothing
         .Close()
     End With


 Conn.Close()
 Conn = Nothing

End Function

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

        strSqlSelc = "SELECT req_id FROM v_notifyissue (NOLOCK)" _
                                              & " ORDER BY req_id"


       RsdPnt = New ADODB.Recordset

       With RsdPnt

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            If .RecordCount <> 0 Then

                  ds.Clear()
                  da.Fill(ds, RsdPnt, "reqid")
                  cmbOptPrint.DataSource = ds.Tables("reqid").DefaultView
                  cmbOptPrint.DisplayMember = "req_id"
                  cmbOptPrint.ValueMember = "req_id"

            End If

         .ActiveConnection = Nothing
         ' .Close()
     End With

     RsdPnt = Nothing

End Sub

Private Sub ViewData()

 If dgvIssue.Rows.Count > 0 Then
    'ClearTmpTableUser("tmp_eqptrn")
    lblCmd.Text = "2"

    With frmAeNotifyIssue
         .Show()
         .Text = "มุมมองข้อมูล"

    End With

    Me.Hide()
    frmMainPro.Hide()

 End If

End Sub

Private Sub DeleteData()

 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim btyConsider As Byte
 Dim strReqid As String
 Dim strSeries As String

     With Conn

              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

     End With

     With dgvIssue

        If .Rows.Count > 0 Then

           strReqid = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim
           strSeries = .Rows(.CurrentRow.Index).Cells(4).Value.ToString.Trim

           btyConsider = MsgBox("รหัส : " & strReqid & vbNewLine _
                                                & "รุ่น / SIZE: " & strSeries & vbNewLine _
                                                & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then

                        Conn.BeginTrans()

                                '------------------------------------ลบตาราง eqpmst--------------------------------------------

                                strSqlCmd = "DELETE FROM notifyissue" _
                                                      & " WHERE req_id ='" & strReqid & "'"

                                Conn.Execute(strSqlCmd)


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

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
  StateLockFind(True)
  gpbSearch.Visible = False

End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  FindDocID()

End Sub

Private Sub FindDocID() 'ค้นหาเอกสาร

 Dim strSearch As String = txtSeek.Text.ToUpper.Trim

   If strSearch <> "" Then

        Select Case cmbType.SelectedIndex()

               Case Is = 0 'รหัส
                    SearchData(0, strSearch)     'ส่งตำเงื่อนไข ,Text ให้ ซับรูทีน SearchData

               Case Is = 1 'สถานะ
                    SearchData(1, strSearch)

               Case Is = 2 'รายละเอียดอุปกรณ์
                    SearchData(2, strSearch)

               Case Is = 3 'แผนกที่แจ้งปัญหา
                    SearchData(3, strSearch)

               Case Is = 4 'รายละเอียดกลุ่ม
                    SearchData(4, strSearch)

               Case Is = 5 'รุ่นอุปกรณ์
                    SearchData(5, strSearch)

        End Select

   Else

     MsgBox("โปรดระบุข้อมูลเพื่อค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
     txtSeek.Focus()

End If

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
                                     strFieldFilter = "req_id like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 1
                                      strFieldFilter = "sta_notify like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 2
                                    strFieldFilter = "eqpnm like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 3
                                      strFieldFilter = "dep_notify like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 4
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 5
                                      strFieldFilter = "shoe like '%" & ReplaceQuote(strSearch) & "%'"

                    End Select


                    strSqlCmdSelc = "SELECT * FROM v_notifyissue (NOLOCK)" _
                                                  & " WHERE " & strFieldFilter _
                                                  & " ORDER BY req_id"


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


                   Rsd = Nothing
                   'Rsd.Close()

                   End With

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

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
 InputDeptData()

End Sub

Private Sub btnPrntPrevw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntPrevw.Click

 Dim strDocId As String = cmbOptPrint.Text.ToUpper.Trim

     If strDocId <> "" Then

        frmMainPro.lblRptDesc.Text = strDocId
        frmRptIssueReceive.Show()

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

     With Conn

         If .State Then .Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

     End With

     strSqlSelc = "SELECT * " _
                          & " FROM notifyissue (NOLOCK)" _
                          & " WHERE req_id = '" & strSelectCode.ToString.Trim & "'"

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                 For i As Integer = 1 To .RecordCount

                                       strSqlCmdPic = "SELECT * " _
                                                               & " FROM tmp_notifyissue (NOLOCK)"

                                       RsdPic = New ADODB.Recordset
                                       RsdPic.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                                       RsdPic.LockType = ADODB.LockTypeEnum.adLockOptimistic
                                       RsdPic.Open(strSqlCmdPic, Conn, , , )

                                                     RsdPic.AddNew()
                                                     RsdPic.Fields("user_id").Value = frmMainPro.lblLogin.Text.ToString.Trim
                                                     RsdPic.Fields("req_id").Value = .Fields("req_id").Value
                                                     RsdPic.Fields("req_sta").Value = .Fields("req_sta").Value
                                                     RsdPic.Fields("group").Value = .Fields("group").Value
                                                     RsdPic.Fields("to_dep").Value = .Fields("to_dep").Value
                                                     RsdPic.Fields("from_notify").Value = .Fields("from_notify").Value
                                                     RsdPic.Fields("dep_notify").Value = .Fields("dep_notify").Value
                                                     RsdPic.Fields("order").Value = .Fields("order").Value
                                                     RsdPic.Fields("eqpnm").Value = .Fields("eqpnm").Value
                                                     RsdPic.Fields("shoe").Value = .Fields("shoe").Value
                                                     RsdPic.Fields("size").Value = .Fields("size").Value
                                                     RsdPic.Fields("amount").Value = .Fields("amount").Value
                                                     RsdPic.Fields("issue").Value = .Fields("issue").Value
                                                     RsdPic.Fields("cause").Value = .Fields("cause").Value
                                                     RsdPic.Fields("needdate").Value = .Fields("needdate").Value
                                                     RsdPic.Fields("needtime").Value = .Fields("needtime").Value
                                                     RsdPic.Fields("fxissue").Value = .Fields("fxissue").Value
                                                     RsdPic.Fields("wantdate").Value = .Fields("wantdate").Value
                                                     RsdPic.Fields("wanttime").Value = .Fields("wanttime").Value
                                                     RsdPic.Fields("pic_Issue").Value = .Fields("pic_Issue").Value
                                                     RsdPic.Fields("person1_sta").Value = .Fields("person1_sta").Value
                                                     RsdPic.Fields("person1").Value = .Fields("person1").Value
                                                     RsdPic.Fields("person1_date").Value = .Fields("person1_date").Value
                                                     RsdPic.Fields("person2_sta").Value = .Fields("person2_sta").Value
                                                     RsdPic.Fields("person2").Value = .Fields("person2").Value
                                                     RsdPic.Fields("person2_date").Value = .Fields("person2_date").Value
                                                     RsdPic.Fields("person3_sta").Value = .Fields("person3_sta").Value
                                                     RsdPic.Fields("person3").Value = .Fields("person3").Value
                                                     RsdPic.Fields("person3_date").Value = .Fields("person3_date").Value
                                                     RsdPic.Fields("person4_sta").Value = .Fields("person4_sta").Value
                                                     RsdPic.Fields("person4").Value = .Fields("person4").Value
                                                     RsdPic.Fields("person4_date").Value = .Fields("person4_date").Value
                                                     RsdPic.Fields("recordby").Value = .Fields("recordby").Value
                                                     RsdPic.Fields("record_date").Value = .Fields("record_date").Value
                                                     RsdPic.Fields("lastby").Value = .Fields("lastby").Value
                                                     RsdPic.Fields("last_date").Value = .Fields("last_date").Value
                                                     RsdPic.Fields("remark").Value = .Fields("remark").Value


                                                     Dim RsdSteam As New ADODB.Stream
                                                     Dim strPicSign02 As String
                                                     Dim strPicSign03 As String
                                                     Dim strPicSign04 As String

                                                     RsdSteam.Type = StreamTypeEnum.adTypeBinary
                                                     RsdSteam.Open()


                                                     '--------------------------------------โหลดรูปภาพลายเซ็นต์ ผู้อนุมัติแจ้ง -------------------------------------

                                                     If .Fields("person2").Value.ToString.Trim <> "" Then
                                                         strPicSign02 = CallPathSignPicture(.Fields("person2").Value.ToString.Trim)
                                                     Else
                                                         strPicSign02 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                                     End If

                                                     RsdSteam.LoadFromFile(strPicSign02)
                                                     RsdPic.Fields("sign_approve2").Value = RsdSteam.Read


                                                     '--------------------------------------โหลดรูปภาพลายเซ็นต์ ผู้รับแจ้ง -------------------------------------

                                                     If .Fields("person3").Value.ToString.Trim <> "" Then
                                                         strPicSign03 = CallPathSignPicture(.Fields("person3").Value.ToString.Trim)
                                                     Else
                                                         strPicSign03 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                                     End If

                                                     RsdSteam.LoadFromFile(strPicSign03)
                                                     RsdPic.Fields("sign_approve3").Value = RsdSteam.Read

                                                     RsdPic.Update()


                                                     '--------------------------------------โหลดรูปภาพลายเซ็นต์ ผู้อนุมัติรับแจ้ง -------------------------------------

                                                     If .Fields("person4").Value.ToString.Trim <> "" Then
                                                         strPicSign04 = CallPathSignPicture(.Fields("person4").Value.ToString.Trim)
                                                     Else
                                                         strPicSign04 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                                     End If

                                                     RsdSteam.LoadFromFile(strPicSign04)
                                                     RsdPic.Fields("sign_approve4").Value = RsdSteam.Read

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

'----------------------- ฟังก์ชั่นปรับขนาด Size รูปภาพ -------------------------------------------------------------------
Private Function ResizeImage(ByVal img As Bitmap, ByVal width As Integer, ByVal height As Integer) As Bitmap

 Dim newBit As New Bitmap(width, height) 'new blank bitmap
 Dim g As Graphics = Graphics.FromImage(newBit)
 'change interpolation for reduction quality
 g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
 g.DrawImage(img, 0, 0, width, height)
 Return newBit

End Function

Private Sub btnPrntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntCancel.Click
 StateLockFind(True)
 gpbOptPrint.Visible = False

End Sub

Private Sub dgvIssue_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvIssue.RowsAdded
   dgvIssue.Rows(e.RowIndex).Height = 27
End Sub
End Class