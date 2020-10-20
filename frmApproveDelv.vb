Imports ADODB
Imports System.IO
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmApproveDelv
  Dim intBkPageCount As Integer
  Dim blnHaveFilter As Boolean    'กรณีกรองข้อมูล

  Dim dubNumberStart As Double   'ถูกกำหนด = 1
  Dim dubNumberEnd As Double     'ถูกกำหนด = 2100

  Dim strSqlFindData As String
  Dim strDocCode As String = "F12"

Private Sub frmApproveDelv_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

     InputDeptData()

     If FormCount("frmRptDelvApprove2") > 0 Then

        With frmRptDelvApprove2

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

 'InputDeptData()

End Sub

Private Sub frmApproveDelv_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
 lblHeight.Text = Me.Height.ToString.Trim
 lblWidth.Text = Me.Width.ToString.Trim

 lblTop.Text = Me.Top.ToString.Trim
 lblLeft.Text = Me.Left.ToString.Trim
End Sub

Private Sub frmApproveDelv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 Me.Dispose()
End Sub

Private Sub frmApproveDelv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Me.WindowState = FormWindowState.Maximized     'ขยายขนาดเต็มหน้าจอ
  StdDateTimeThai()                           'เรียก ซับรูทีน StdDateTimeThai
  tlsBarFmr.Cursor = Cursors.Hand             'ให้คอร์เซอร์ตรง Toolstripbar เป็นรูปมือ

  dubNumberStart = 1                          'ให้แถวเเรกใน Recordset = 1
  dubNumberEnd = 2100                         'ให้แถวเเรกใน Recordset = 2100

  PreviewUser(frmMainPro.lblLogin.Text.Trim)
  PreGroupType()
  InputDeptData()
  tabCmd.Focus()

End Sub

Private Function FormCount(ByVal frmName As String) As Long
 Dim frm As Form

     For Each frm In My.Application.OpenForms

         If frm Is My.Forms.frmRptDelvApprove2 Then
            FormCount = FormCount + 1
         End If

     Next

End Function

Private Sub PreGroupType()

Dim strGpTopic(4) As String
Dim i As Byte

      strGpTopic(0) = "เลขที่เอกสาร"
      'strGpTopic(1) = "สถานะ"
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

Sub PreviewUser(ByVal strUsr As String)

   Select Case strUsr

          Case Is = "PRADIST"
                    lblName.Text = "คุณประดิษฐ์ สังข์ทอง"
                    lblDept.Text = "122000 ตัดชิ้นส่วน"

          Case Is = "ITHISAK"
                    lblName.Text = "คุณอิทธิศักดิ์ ปานอุดมกิตติ์"
                    lblDept.Text = "125000 EVA INJECTION"

          Case Is = "TODSAPORN"
                    lblName.Text = "คุณทศพร ภูมิสุนทรธรรม"
                    lblDept.Text = "126000 ฉีด PU"

          Case Is = "SATHID"
                    lblName.Text = "คุณสถิตถ์ แสนรัก"
                    lblDept.Text = "123000 เย็บ"

          Case Is = "TECHIN"
                    lblName.Text = "คุณเตชินม์ บัวหลง"
                    lblDept.Text = "121000 ผลิตโฟม"

          Case Is = "PEERA"
                    lblName.Text = "คุณพีระ แสงอรุณบริสุทธิ์"
                    lblDept.Text = "124000 ฉีด PVC"

          Case Is = "BOONTUM"
                    lblName.Text = "คุณบุญธรรม น่วมสวัสดิ์"
                    lblDept.Text = "เทคนิคอุปกรณ์"

          Case Is = "SUTID"
                    lblName.Text = "คุณสุทิด กรชม"
                    lblDept.Text = "เทคนิคอุปกรณ์"

    End Select

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
                                      & " ORDER BY doc_id"


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

                                        '---------------------------------------ค้นหาข้อมูล-------------------------------
                                        .MoveFirst()
                                        .Find(strSqlFind)
                                         lblPage.Text = Str(.AbsolutePage)
                                        '-----------------------------------------------------------------------------

                                        If .Fields("RowNumber").Value >= 2100 Then

                                            dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                                            dubNumberEnd = .Fields("RowNumber").Value + 1000

                                        Else

                                            dubNumberStart = 1
                                            dubNumberEnd = 2100

                                        End If

                                        strSqlFindData = strSqlFind

                                        InputDeptData()


                                                For i = 0 To dgvTransfer.Rows.Count - 1

                                                        If InStr(UCase(dgvTransfer.Rows(i).Cells(2).Value), strSearchTxt.Trim.ToUpper) <> 0 Then
                                                                dgvTransfer.CurrentCell = dgvTransfer.Item(2, i)
                                                                dgvTransfer.Focus()
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

Private Sub StateLockFind(ByVal Sta As Boolean)
 tabCmd.Enabled = Sta
 dgvTransfer.Enabled = Sta
 tlsBarFmr.Enabled = Sta

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
 Dim strApprovesta As String = ""

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
                                     strFieldFilter = "doc_id like '" & ReplaceQuote(strSearch) & "%'"

                              'Case Is = 1

                              '        If frmMainPro.lblLogin.Text = "SUTID" Then

                              '               If strSearch = "อนุมัติแล้ว" Then
                              '                  strFieldFilter = "req_sta = '" & "2" & "'"

                              '               ElseIf strSearch = "รอเซ็นอนุมัติ" Then
                              '                  strFieldFilter = "req_sta = '" & "1" & "'"

                              '               Else
                              '                  strFieldFilter = "req_sta = '" & "3" & "'"

                              '               End If

                              '        Else

                              '               If strSearch = "อยู่ระหว่างดำเนินการ" Then
                              '                  strFieldFilter = "req_sta = '" & "1" Or "2" & "'"

                              '               ElseIf strSearch = "รอดำเนินการ" Then
                              '                  strFieldFilter = "req_sta = '" & "0" & "'"

                              '               Else
                              '                  strFieldFilter = "req_sta = '" & "3" & "'"

                              '               End If

                              '        End If
                              '            'strFieldFilter = "sta_notify like '" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 1
                                      strFieldFilter = "send_nm like '" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 2
                                      strFieldFilter = "rvc_nm like '" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 3
                                      strFieldFilter = "rvc_dep_nm like '" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 4
                                      strFieldFilter = "remark like '" & ReplaceQuote(strSearch) & "%'"

                    End Select


                              Select Case frmMainPro.lblLogin.Text

                                     Case Is = "PRADIST"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND rvc_dep_nm LIKE '%" & "122000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                                     Case Is = "ITHISAK"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND rvc_dep_nm LIKE '%" & "125000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                                     Case Is = "TODSAPORN"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND rvc_dep_nm LIKE '%" & "126000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                                     Case Is = "SATHID"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND rvc_dep_nm LIKE '%" & "123000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                                     Case Is = "TECHIN"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND rvc_dep_nm LIKE '%" & "121000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"
                                     Case Is = "PEERA"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND rvc_dep_nm LIKE '%" & "124000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                                     Case Is = "SUTID"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND ps01_result = 'True' " _
                                                                 & " ORDER BY doc_id"


                                     Case Is = "BOONTUM"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND ps02_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                              End Select

              Else

                    Select Case frmMainPro.lblLogin.Text

                           Case Is = "PRADIST"

                                       strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND rvc_dep_nm LIKE '%" & "122000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                           Case Is = "ITHISAK"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND rvc_dep_nm LIKE '%" & "125000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                           Case Is = "TODSAPORN"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND rvc_dep_nm LIKE '%" & "126000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                           Case Is = "SATHID"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND rvc_dep_nm LIKE '%" & "123000" & "'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                           Case Is = "TECHIN"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND rvc_dep_nm LIKE '%" & "121000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                           Case Is = "PEERA"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND rvc_dep_nm  LIKE '%" & "124000" & "%'" _
                                                                 & " AND ps03_result = 'True' " _
                                                                 & " ORDER BY doc_id"

                           Case Is = "SUTID"

                                         strSqlCmdSelc = "SELECT  * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND ps01_result = 'True' " _
                                                                 & " ORDER BY doc_id"


                           Case Is = "BOONTUM"

                                         strSqlCmdSelc = "SELECT * FROM v_delvmst2 (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND ps02_result = 'True' " _
                                                                 & " ORDER BY doc_id"


                   End Select


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

                                    dgvTransfer.Rows.Clear()

                                    intCounter = 0
                                    lblDocnull.Visible = False

                                    Do While Not .EOF

                                                '----------------------------------- Preview ข้อมูลตาม user login ---------------------------

                                                If frmMainPro.lblLogin.Text = "BOONTUM" Then

                                                            Select Case .Fields("ps03_result").Value

                                                                   Case Is = False
                                                                       imgStaReq = My.Resources._16x16_ledred
                                                                       strApprovesta = "รอเซ็นอนุมัติ"

                                                                   Case Else
                                                                       imgStaReq = My.Resources._16x16_ledgreen
                                                                        strApprovesta = "เซ็นอนุมัติแล้ว"

                                                             End Select


                                                 ElseIf frmMainPro.lblLogin.Text = "SUTID" Then

                                                           Select Case .Fields("ps02_result").Value

                                                                   Case Is = False
                                                                       imgStaReq = My.Resources._16x16_ledred
                                                                       strApprovesta = "รอเซ็นอนุมัติ"

                                                                   Case Else
                                                                       imgStaReq = My.Resources._16x16_ledgreen
                                                                       strApprovesta = "เซ็นอนุมัติแล้ว"


                                                           End Select

                                                 Else       'กรณีเป็น user อื่น

                                                          Select Case .Fields("ps04_result").Value

                                                                 Case Is = False
                                                                     imgStaReq = My.Resources._16x16_ledred
                                                                     strApprovesta = "รอเซ็นอนุมัติ"

                                                                 Case Else
                                                                      imgStaReq = My.Resources._16x16_ledgreen
                                                                      strApprovesta = "เซ็นอนุมัติแล้ว"


                                                          End Select


                                       End If


                                            dgvTransfer.Rows.Add( _
                                                                          imgStaReq, strApprovesta, _
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

                                         .MoveNext()    'ข้ามไปที่ระเบียนใหม่

                                     Loop

                            Else
                                intBkPageCount = 1
                                txtPage.Text = "1"

                                lblDocnull.Visible = True

                            End If

                   .Close()

              End With

  Rsd = Nothing

 Conn.Close()
 Conn = Nothing

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
     dgvTransfer.Focus()
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

                    Case Is = 0 'อนุมัติ / พิมพ์

                            blnReturn = CheckUserEntry(strDocCode, "act_print")
                            If blnReturn Then

                               If dgvTransfer.Rows.Count <> 0 Then

                                  'ClearTmpTableUser("tmp_notifyissue")
                                  frmMainPro.lblRptDesc.Text = Me.dgvTransfer.Rows(Me.dgvTransfer.CurrentRow.Index).Cells(2).Value.ToString

                                  With frmRptDelvApprove2
                                       .Show()

                                  End With
                                  Me.Hide()
                                  frmMainPro.Hide()

                               End If

                            Else
                                MsnAdmin()
                            End If

                   Case Is = 1    'กรองข้อมูล

                            If dgvTransfer.Rows.Count > 0 Then

                                With gpbFilter

                                     .Top = 230
                                     .Left = 210
                                      Width = 348
                                     .Height = 125

                                     .Visible = True
                                     .BringToFront()

                                     cmbFilter.SelectedItem = cmbFilter.Items(0)
                                     txtFilter.Text = _
                                              dgvTransfer.Rows(dgvTransfer.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                                     StateLockFind(False)
                                     txtFilter.Focus()

                                 End With

                            End If


                    Case Is = 2 'ค้นหาข้อมูล

                            If dgvTransfer.Rows.Count > 0 Then

                               With gpbSearch

                                    .Top = 230
                                    .Left = 210
                                    .Width = 348
                                    .Height = 125

                                    .BringToFront()
                                    .Visible = True

                                    cmbType.SelectedItem = cmbType.Items(0)
                                    txtSeek.Text = _
                                             dgvTransfer.Rows(dgvTransfer.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                                    StateLockFind(False)
                                    txtSeek.Focus()

                               End With

                            End If

                    Case Is = 3 'ฟื้นฟูข้อมูล
                              InputDeptData()

                    Case Is = 4 'ออก
                            Me.Close()

          End Select

  End With
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
                                     strFieldFilter = "doc_id LIKE '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 1
                                      strFieldFilter = "sta_notify like '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 2
                                      strFieldFilter = "send_nm LIKE '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 3
                                      strFieldFilter = "rvc_nm LIKE '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 4
                                      strFieldFilter = "rvc_dep_nm LIKE '%" & ReplaceQuote(strSearch) & "%'"

                              Case Is = 5
                                      strFieldFilter = "remark LIKE '%" & ReplaceQuote(strSearch) & "%'"

                    End Select


                    strSqlCmdSelc = "SELECT * FROM v_delvmst2 (NOLOCK)" _
                                                  & " WHERE " & strFieldFilter _
                                                  & " ORDER BY doc_id"

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
                             gpbFilter.Visible = True

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

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click

 If blnHaveFilter Then
    blnHaveFilter = False
    InputDeptData()

 End If

    StateLockFind(True)
    gpbFilter.Visible = False

End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  FindDocID()

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

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
  StateLockFind(True)
  gpbSearch.Visible = False

End Sub

Private Sub dgvIssue_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvTransfer.RowsAdded
  dgvTransfer.Rows(e.RowIndex).Height = 27
End Sub

Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  InputDeptData()
End Sub
End Class