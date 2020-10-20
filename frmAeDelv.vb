Imports ADODB
Imports System.IO

Public Class frmAeDelv

Dim strDateDefault As String

Dim IsShowPs As Boolean
Dim IsShowEqp As Boolean

Dim staAction As String
Dim staPs As String

Dim strTxt_C6 As String
Dim strTxt_C7 As String
Dim strTxt_C8 As String
Dim strTxt_C9 As String
Dim strTxt_C10 As String

Protected Overrides ReadOnly Property CreateParams() As CreateParams 'ป้องกันการปิดโดยใช้ปุ่ม Close Button
Get
   Dim cp As CreateParams = MyBase.CreateParams
   Const CS_DBLCLKS As Int32 = &H8
   Const CS_NOCLOSE As Int32 = &H200
   cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE

   Return cp
End Get
End Property

Private Sub ClearTmpTable(ByVal bytOption As Byte, ByVal strPsID As String)

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String = ""

            With Conn

                    If .State Then .Close()

                        .ConnectionString = strConnAdodb
                        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                        .ConnectionTimeout = 90
                        .Open()

            End With

            With Conn

                    Select Case bytOption

                           Case Is = 0   'ลบข้อมูลหลังจากปิดฟอร์ม

                               strSqlCmd = "Delete FROM tmp_delvtrn" _
                                                       & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"
                               .Execute(strSqlCmd)

                           Case Is = 1

                               strSqlCmd = "Delete FROM tmp_delvtrn" _
                                                        & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                        & " AND docno ='" & strPsID.ToString.Trim & "'"
                               .Execute(strSqlCmd)

                    End Select

            End With

    Conn.Close()
    Conn = Nothing

End Sub

Private Sub frmAeDelv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTable(0, "")
   frmDelv.lblCmd.Text = "0" 'เคลียร์สถานะ
   Me.Dispose()

End Sub

Private Sub frmAeDelv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Dim dteComputer As Date = Now()
  Dim strCurrentDate As String

  StdDateTimeThai()
  strCurrentDate = dteComputer.Date.ToString("dd/MM/yyyy")  

  PrePsData()
  PreEqpData()

  FindPsData("", 2)
  FindEqpData("", 5)

   Select Case frmDelv.lblCmd.Text.ToString

          Case Is = "0" 'เพิ่มข้อมูล

               With Me
                    .Text = "เพิ่มข้อมูล"
               End With

               staAction = "0"
               With txtBegin
                    .Text = strCurrentDate
                    strDateDefault = strCurrentDate

               End With

          Case Is = "1" 'แก้ไขข้อมูล

               With Me
                    .Text = "เเก้ไขข้อมูล"
               End With

               staAction = "1"
               LockEditData()

          Case Is = "2" 'มุมมองข้อมูล

               With Me
                    .Text = "มุมมองข้อมูล"
               End With

               staAction = "2"
               LockEditData()
               btnSaveData.Enabled = False

    End Select
    txtRemark.Focus()

End Sub

Private Sub LockEditData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset
Dim RsdWc As New ADODB.Recordset

Dim strSqlSelc As String

Dim strCmd As String

Dim blnHaveData As Boolean
Dim dteComputer As Date = Now()

Dim strSqlCmd As String = ""

Dim strCode As String = frmDelv.dgvShoe.Rows(frmDelv.dgvShoe.CurrentRow.Index).Cells(2).Value.ToString.Trim

            With Conn

                 If .State Then .Close()
                    .ConnectionString = strConnAdodb
                    .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                    .ConnectionTimeout = 90
                    .Open()

              End With

              strSqlSelc = "SELECT * " _
                                    & "FROM v_delvmst2 (NOLOCK)" _
                                    & " WHERE doc_id ='" & strCode & "'"

              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlSelc, Conn, , , )

                          If .RecordCount <> 0 Then

                                    txtBegin.Text = Mid(.Fields("doc_date").Value.ToString.Trim, 1, 10)
                                    strDateDefault = Mid(.Fields("doc_date").Value.ToString.Trim, 1, 10)

                                    lblDocID.Text = .Fields("doc_id").Value.ToString.Trim
                                    txtRemark.Text = .Fields("remark").Value.ToString.Trim

                                    lblDocID.Visible = True
                                    lblDocTopic.Visible = True

                                    lblSendId.Text = .Fields("send_id").Value.ToString.Trim
                                    lblSendNm.Text = .Fields("send_nm").Value.ToString.Trim

                                    lblRvcId.Text = .Fields("rvc_id").Value.ToString.Trim
                                    lblRvcNm.Text = .Fields("rvc_nm").Value.ToString.Trim
                                    lblRvcDep.Text = .Fields("rvc_dep_nm").Value.ToString.Trim

                                    strCmd = frmDelv.lblCmd.Text.ToString

                                    Select Case strCmd

                                           Case Is = "1" 'ให้ล็อคตอนแก้ไข
                                           Case Is = "2" 'ให้ล็อคตอนมุมมอง
                                                      btnSaveData.Enabled = False

                                    End Select

                                    '-------------------------- บันทึกข้อมูลในตาราง tmp_delvtrn -----------------------------

                                    strSqlCmd = "INSERT INTO tmp_delvtrn" _
                                                         & " SELECT user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                                         & "  FROM delvtrn" _
                                                         & " WHERE doc_id ='" & strCode & "'"

                                    Conn.Execute(strSqlCmd)

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
     ShowScrapItem()
End If

End Sub

Private Sub PrePsData()

Dim strGpTopic(1) As String
Dim i As Byte

      strGpTopic(0) = "รหัสอุปกรณ์"
      strGpTopic(1) = "ชื่อพนักงาน"      

      With cmbPs

              For i = 0 To 1
                 .Items.Add(strGpTopic(i))
              Next i

        .SelectedItem = .Items(1)

      End With

End Sub

Private Sub PreEqpData()

Dim strGpTopic(4) As String
Dim i As Byte

      strGpTopic(0) = "กลุ่มอุปกรณ์"
      strGpTopic(1) = "รายละเอียดกลุ่ม"
      strGpTopic(2) = "รหัสอุปกรณ์"
      strGpTopic(3) = "ชุดโมล์ด"
      strGpTopic(4) = "รายละเอียดอุปกรณ์"

      With cmbEqp

              For i = 0 To 4
                 .Items.Add(strGpTopic(i))
              Next i

        .SelectedItem = .Items(2)

      End With

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

    Me.Close()

  'On Error Resume Next

  'Dim strCode As String

  '  With frmDelv.dgvShoe

  '       If .Rows.Count > 0 Then
  '          strCode = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim
  '          lblComplete.Text = strCode

  '       End If

  '  End With

  '  Me.Hide()

  '  frmMainPro.Show()
  '  frmDelv.Show()
End Sub

Private Sub txtBegin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBegin.GotFocus

    With mskBegin
            txtBegin.SendToBack()
            .BringToFront()
            .Focus()

    End With

End Sub

Private Sub mskBegin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskBegin.GotFocus

Dim i, x As Integer

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskBegin

        If txtBegin.Text <> "__/__/____" Then

                        x = Len(txtBegin.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtBegin.Text.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = "_"
                                          Case Else

                                                    If InStr("0123456789/", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select

                         Next i

                        Select Case strMerge.ToString.Length
                                  Case Is = 10
                                            .SelectionStart = 0
                                  Case Is = 7
                                            '.SelectionStart = 1
                                  Case Is = 6
                                            '.SelectionStart = 2
                                  Case Is = 5
                                            '.SelectionStart = 3
                                  Case Is = 4
                                            '.SelectionStart = 4
                                  Case Is = 3
                                            '.SelectionStart = 5
                        End Select

                        .SelectedText = strMerge

                End If

        .SelectAll()

End With

End Sub

Private Sub mskBegin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskBegin.KeyDown

Dim intChkPoint As Integer

    With mskBegin

            Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                              
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง
                              txtRemark.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select

    End With

End Sub

Private Sub mskBegin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskBegin.KeyPress

Select Case e.KeyChar
          Case Is = Chr(13)
                  txtRemark.Focus()
          Case Is = Chr(46) 'เครื่องหมายจุลภาค(.)
                  'mskBegin.SelectionStart = 6
End Select

End Sub

Private Sub mskBegin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskBegin.LostFocus

Dim i, x As Integer
Dim z As Date

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskBegin

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = ","
                                          Case Is = "+"
                                          Case Is = "_"
                                          Case Else

                                                    If InStr("0123456789/", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskBegin.Text = ""
                    strMerge = "#" & strMerge & "#"
                    z = CDate(strMerge)
                    txtBegin.Text = z.ToString("dd/MM/yyyy")


               Catch ex As Exception
                    txtBegin.Text = strDateDefault
                    mskBegin.Text = ""

               End Try

    mskBegin.SendToBack()
    txtBegin.BringToFront()

End With

End Sub

Private Sub txtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.GotFocus
  txtRemark.SelectAll()
End Sub

Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
    Dim intChkPoint As Integer

    With txtRemark

            Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น 
                              txtBegin.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง                              
                              txtBegin.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select

    End With

End Sub

Private Sub txtRemark_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRemark.KeyPress
If e.KeyChar = Chr(13) Then
      txtBegin.Focus()
End If

End Sub

Private Sub txtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.LostFocus
    With txtRemark
            .Text = .Text.ToString.Trim.ToUpper
    End With

End Sub

Private Sub btnShwPs1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShwPs1.Click
   ShowPsData1()
End Sub

Private Sub ShowPsData1() 'ผู้โอน

   IsShowPs = Not IsShowPs
   If IsShowPs Then

        With gpbPs

                .Visible = True
                .Left = 285
                .Top = 64
                .Height = 384
                .Width = 734
                .Text = "ค้นหาผู้โอน"

                 dgvPs.BackgroundColor = Color.Chartreuse

         End With

        StateLockFindDept(False)
        dgvPs.Focus()
        staPs = "1"

   Else

       StateLockFindDept(True)
       dgvPs.Visible = False
       staPs = "0"

   End If

End Sub

Private Sub ShowPsData2() 'ผู้รับอุปกรณ์

   IsShowPs = Not IsShowPs
   If IsShowPs Then

        With gpbPs

                .Visible = True
                .Left = 285
                .Top = 64
                .Height = 384
                .Width = 734
                .Text = "ค้นหาผู้รับอุปกรณ์"

                 dgvPs.BackgroundColor = Color.DarkOliveGreen    'เปลี่ยนสี Background 

         End With

        StateLockFindDept(False)
        dgvPs.Focus()
        staPs = "2"

   Else

       StateLockFindDept(True)
       dgvPs.Visible = False
       staPs = "0"

   End If

End Sub

Private Sub StateLockFindDept(ByVal Sta As Boolean)

        gpbHead.Enabled = Sta
        gpbItem.Enabled = Sta

        btnAdd.Enabled = Sta
        btnDel.Enabled = Sta
        btnSaveData.Enabled = Sta

        Select Case staAction

                  Case Is = "1" 'แก้ไขข้อมูล                        
                  Case Is = "2" 'มุมมองข้อมูล
                          btnSaveData.Enabled = False

        End Select

End Sub

Private Sub btnPsExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPsExit.Click

  StateLockFindDept(True)
  gpbPs.Visible = False
  IsShowPs = False
  staPs = "0"

End Sub

Private Sub FindPsData(ByVal strWording As String, ByVal bytOption As Byte)

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String

        With Conn

                If .State Then .Close()

                   .ConnectionString = strConnDbHr2
                   .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                   .ConnectionTimeout = 90
                   .Open()

         End With

        Select Case bytOption

                  Case Is = 0  'ค้นหาโดยระบุรหัส

                          strSqlCmdSelc = "SELECT pfs_id" _
                                                         & ",'คุณ ' +  RTRIM(name) + '   ' + RTRIM(last_name) AS nw_name " _
                                                         & ",RTRIM(division) + ' : ' + RTRIM([DESC]) AS nw_dept" _
                                                         & " FROM v_depmain_job2 (NOLOCK)" _
                                                         & " WHERE pfs_id LIKE '%" & ReplaceQuote(strWording) & "%'" _
                                                         & " AND retire =0" _
                                                         & " ORDER BY pfs_id"

                  Case Is = 1  'ค้นหาโดยรายละเอียด

                          strSqlCmdSelc = "SELECT pfs_id" _
                                                        & ",'คุณ ' +  RTRIM(name) + '   ' + RTRIM(last_name) AS nw_name " _
                                                        & ",RTRIM(division) + ' : ' + RTRIM([DESC]) AS nw_dept" _
                                                        & " FROM v_depmain_job2 (NOLOCK)" _
                                                        & " WHERE 'คุณ ' +  RTRIM(name) + '   ' + RTRIM(last_name) LIKE '%" & ReplaceQuote(strWording) & "%'" _
                                                        & " AND retire =0" _
                                                        & " ORDER BY pfs_id"

                  Case Else    'ค่าตั้งต้น

                          strSqlCmdSelc = "SELECT TOP 200 pfs_id" _
                                                        & ",'คุณ ' +  RTRIM(name) + '   ' + RTRIM(last_name) AS nw_name " _
                                                        & ",RTRIM(division) + ' : ' + RTRIM([DESC]) AS nw_dept" _
                                                        & " FROM v_depmain_job2 (NOLOCK)" _
                                                        & " WHERE retire =0" _
                                                        & " ORDER BY pfs_id"

         End Select

                With Rsd

                     .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                     .LockType = ADODB.LockTypeEnum.adLockOptimistic
                     .Open(strSqlCmdSelc, Conn, , , )

                     If .RecordCount <> 0 Then

                         dgvPs.Rows.Clear()
                         Do While Not .EOF

                                dgvPs.Rows.Add( _
                                                 .Fields("pfs_id").Value.ToString.Trim, _
                                                 .Fields("nw_name").Value.ToString.Trim, _
                                                 "เลือก", _
                                                 .Fields("nw_dept").Value.ToString.Trim _
                                               )

                                .MoveNext()

                                  Loop


                         Else

                              MsgBox("ไม่พบข้อมูล :" & strWording & " ในระบบ" & vbNewLine _
                                          & "โปรดระบุการค้นหาข้อมูลใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Not Found Data!!")
                              txtPsSeek.Focus()

                         End If


                        .ActiveConnection = Nothing
                        .Close()

                End With

    Rsd = Nothing

    Conn.Close()
    Conn = Nothing

    ShowEmployeePicture()

End Sub

Private Sub ShowEmployeePicture()
Dim strDocID As String

    With dgvPs

             If .Rows.Count > 0 Then
                strDocID = .Rows(.CurrentRow.Index).Cells(0).Value.ToString

                lblPsID.Text = strDocID
                lblPsName.Text = .Rows(.CurrentRow.Index).Cells(1).Value.ToString

                LoadEmployeePicture(strDocID)

             Else

                 lblPsID.Text = ""
                 lblPsName.Text = ""
                 picPs.Image = My.Resources.blank
             End If

     End With

End Sub

Private Sub LoadEmployeePicture(ByVal strPsID As String)

Dim strLoadFilePicture As String
Dim strPathPicture As String = "\\10.32.0.15\data2\Employee\"


      strLoadFilePicture = strPathPicture & strPsID & ".jpg"

      If File.Exists(strLoadFilePicture) Then

            '---------------------------Load รูปภาพมาแล้วสามาแก้ไขชื่อไฟล์ต้นฉบับได้---------------------
            Dim img As Image
                  img = Image.FromFile(strLoadFilePicture)
            Dim s As String = ImageToBase64(img, System.Drawing.Imaging.ImageFormat.Jpeg)
                  img.Dispose()
                  picPs.Image = Base64ToImage(s)

      Else

             picPs.Image = My.Resources.blank

      End If

End Sub

Private Sub btnPsSeek_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPsSeek.Click
   Dim strSearch As String

    strSearch = txtPsSeek.Text.Trim.ToUpper
    If Len(strSearch) <> 0 Then

        FindPsData(strSearch, cmbPs.SelectedIndex)
        btnPsSeek.Focus()

    End If

End Sub

Private Sub dgvPs_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPs.CellClick

With dgvPs

        If .Rows.Count > 0 Then

                Select Case .CurrentCell.ColumnIndex

                          Case Is = 2 'คลิกเลือกข้อมูล

                                     Select Case staPs

                                            Case Is = "1"
                                                  lblSendId.Text = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
                                                  lblSendNm.Text = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim
                                                  btnShwPs1.Focus()

                                            Case Is = "2"
                                                  lblRvcId.Text = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
                                                  lblRvcNm.Text = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim
                                                  lblRvcDep.Text = .Rows(.CurrentRow.Index).Cells(3).Value.ToString.Trim
                                                  btnShwPs2.Focus()

                                     End Select

                                     StateLockFindDept(True)
                                     IsShowPs = False
                                     gpbPs.Visible = False


                End Select

        End If

End With

ShowEmployeePicture()

End Sub

Private Sub dgvPs_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvPs.KeyDown
    If e.KeyCode = Keys.Enter Then
        e.Handled = True
    End If

End Sub

Private Sub cmbPs_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbPs.KeyPress
    If e.KeyChar = Chr(13) Then
            txtPsSeek.Focus()
    End If

End Sub

Private Sub txtPsSeek_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPsSeek.GotFocus
  txtPsSeek.SelectAll()
End Sub

Private Sub txtPsSeek_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPsSeek.KeyPress
Dim strSearch As String

If e.KeyChar = Chr(13) Then

    strSearch = txtPsSeek.Text.Trim.ToUpper
    If Len(strSearch) <> 0 Then
        FindPsData(strSearch, cmbPs.SelectedIndex)
        btnPsSeek.Focus()
    End If

End If

End Sub

Private Sub txtPsSeek_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPsSeek.LostFocus
With txtPsSeek
     .Text = .Text.ToString.Trim.ToUpper
End With

End Sub

Private Sub dgvPs_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvPs.KeyPress

If e.KeyChar = Chr(13) Then

    With dgvPs

            If .Rows.Count > 0 Then

                     Select Case staPs

                                  Case Is = "1"
                                              lblSendId.Text = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
                                              lblSendNm.Text = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim
                                              btnShwPs1.Focus()

                                   Case Is = "2"
                                             lblRvcId.Text = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
                                             lblRvcNm.Text = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim
                                             lblRvcDep.Text = .Rows(.CurrentRow.Index).Cells(3).Value.ToString.Trim
                                             btnShwPs2.Focus()

                End Select

                StateLockFindDept(True)
                IsShowPs = False
                gpbPs.Visible = False

          End If

        End With

End If

End Sub

Private Sub dgvPs_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvPs.KeyUp

    Select Case e.KeyCode

          Case 33, 34, 35, 36, 38, 40
                 ShowEmployeePicture()

     End Select

End Sub

Private Sub btnShwPs2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShwPs2.Click
  ShowPsData2()
End Sub

Private Sub FindEqpData(ByVal strWording As String, ByVal bytOption As Byte)

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

        Select Case bytOption

                  Case Is = 0 'ค้นหากลุ่มอุปกรณ์

                         strSqlCmdSelc = "SELECT [group],desc_eng,desc_thai " _
                                                         & ",eqp_id,eqp_name,size_act,size_desc,size_id " _
                                                         & " FROM v_eqp_delivr (NOLOCK)" _
                                                         & " WHERE desc_eng LIKE '%" & ReplaceQuote(strWording) & "%'" _
                                                         & " AND delvr_sta = 0" _
                                                         & " AND sent_sta = 0" _
                                                         & " ORDER BY [group],eqp_id,size_id"

                  Case Is = 1 'ค้นหารายละเอียดกลุ่ม

                           strSqlCmdSelc = "SELECT [group],desc_eng,desc_thai " _
                                                         & ",eqp_id,eqp_name,size_act,size_desc,size_id " _
                                                         & " FROM v_eqp_delivr (NOLOCK)" _
                                                         & " WHERE desc_thai LIKE '%" & ReplaceQuote(strWording) & "%'" _
                                                         & " AND delvr_sta = 0" _
                                                         & " AND sent_sta = 0" _
                                                         & " ORDER BY [group],eqp_id,size_id"

                   Case Is = 2 'ค้นหารหัสอุปกรณ์

                           strSqlCmdSelc = "SELECT [group],desc_eng,desc_thai " _
                                                         & ",eqp_id,eqp_name,size_act,size_desc,size_id " _
                                                         & " FROM v_eqp_delivr (NOLOCK)" _
                                                         & " WHERE size_act LIKE '%" & ReplaceQuote(strWording) & "%'" _
                                                         & " AND delvr_sta = 0" _
                                                         & " AND sent_sta = 0" _
                                                         & " ORDER BY [group],eqp_id,size_id"

                  Case Is = 3 'ค้นหาชุดโมลด์

                           strSqlCmdSelc = "SELECT [group],desc_eng,desc_thai " _
                                                         & ",eqp_id,eqp_name,size_act,size_desc,size_id " _
                                                         & " FROM v_eqp_delivr (NOLOCK)" _
                                                         & " WHERE size_desc LIKE '%" & ReplaceQuote(strWording) & "%'" _
                                                         & " AND delvr_sta = 0" _
                                                         & " AND sent_sta = 0" _
                                                         & " ORDER BY [group],eqp_id,size_id"


                    Case Is = 4 'รายละเอียดอุปกรณ์

                           strSqlCmdSelc = "SELECT [group],desc_eng,desc_thai " _
                                                         & ",eqp_id,eqp_name,size_act,size_desc,size_id " _
                                                         & " FROM v_eqp_delivr (NOLOCK)" _
                                                         & " WHERE eqp_name LIKE '%" & ReplaceQuote(strWording) & "%'" _
                                                         & " AND delvr_sta = 0" _
                                                         & " AND sent_sta = 0" _
                                                         & " ORDER BY [group],eqp_id,size_id"

                  Case Else 'ค่าตั้งต้น

                          strSqlCmdSelc = "SELECT [group],desc_eng,desc_thai " _
                                                         & ",eqp_id,eqp_name,size_act,size_desc,size_id " _
                                                         & " FROM v_eqp_delivr (NOLOCK)" _
                                                         & " WHERE delvr_sta = 0" _
                                                         & " AND sent_sta = 0" _
                                                         & " ORDER BY [group],eqp_id,size_id"

         End Select

         With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                          If .RecordCount <> 0 Then

                                  dgvEqpList.Rows.Clear()

                                  Do While Not .EOF

                                            dgvEqpList.Rows.Add( _
                                                                                         .Fields("desc_eng").Value.ToString.Trim, _
                                                                                         .Fields("desc_thai").Value.ToString.Trim, _
                                                                                         .Fields("size_act").Value.ToString.Trim, _
                                                                                         .Fields("size_desc").Value.ToString.Trim, _
                                                                                         .Fields("eqp_name").Value.ToString.Trim, _
                                                                                          "เลือก", _
                                                                                         .Fields("group").Value.ToString.Trim, _
                                                                                         .Fields("eqp_id").Value.ToString.Trim, _
                                                                                         .Fields("size_id").Value.ToString.Trim _
                                                                                         )
                                            .MoveNext()

                                  Loop


                         Else

                               If bytOption <> 5 Then

                                        MsgBox("ไม่พบข้อมูล :" & strWording & " ในระบบ" & vbNewLine _
                                                  & "โปรดระบุการค้นหาข้อมูลใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Not Found Data!!")
                                        txtEqp.Focus()

                                End If


                         End If


                        .ActiveConnection = Nothing
                        .Close()

                End With

    Rsd = Nothing

    Conn.Close()
    Conn = Nothing

End Sub

Private Sub ShowEquipList()

   IsShowEqp = Not IsShowEqp
   If IsShowEqp Then

        With gpbEqpList

                .Visible = True
                .Left = 115
                .Top = 243
                .Height = 460
                .Width = 910

         End With

        StateLockFindDept(False)
        dgvEqpList.Focus()

   Else
       StateLockFindDept(True)
       dgvEqpList.Visible = False

   End If

End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
    ShowEquipList()
End Sub

Private Sub dgvEqpList_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvEqpList.CellClick

Dim strGrp As String
Dim strEqp As String
Dim strSizeDesc As String
Dim strSizeId As String

Dim i As Integer

With dgvEqpList

        If .Rows.Count > 0 Then

                Select Case .CurrentCell.ColumnIndex

                          Case Is = 5 'เลือกข้อมูล

                                 strGrp = .Rows(.CurrentRow.Index).Cells(6).Value 'group
                                 strEqp = .Rows(.CurrentRow.Index).Cells(7).Value
                                 strSizeDesc = .Rows(.CurrentRow.Index).Cells(3).Value 'ชุดโมลด์
                                 strSizeId = .Rows(.CurrentRow.Index).Cells(8).Value

                                 SeekCodeData(strGrp, strEqp, strSizeDesc, strSizeId)
                                 ArrangeNumber()
                                 ShowScrapItem()

                                 '------------------------------ค้นหารหัสที่เพิ่มเข้าไปใหม่------------------------------------------

                                 For i = 0 To dgvItem.Rows.Count - 1

                                    If dgvItem.Rows(i).Cells(4).Value.ToString = strSizeDesc Then
                                       dgvItem.CurrentCell = dgvItem.Item(6, i)
                                       dgvItem.Focus()
                                       Exit For

                                    End If

                                 Next i

                                 StateLockFindDept(True)
                                 IsShowEqp = False
                                 gpbEqpList.Visible = False
                                 txtRemark.Focus()

                End Select

        End If

    End With

End Sub

Private Sub dgvEqpList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvEqpList.KeyDown
    If e.KeyCode = Keys.Enter Then
        e.Handled = True
    End If

End Sub

Private Sub dgvEqpList_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvEqpList.KeyPress

Dim strGrp As String
Dim strEqp As String
Dim strSizeDesc As String
Dim strSizeId As String

Dim i As Integer

If e.KeyChar = Chr(13) Then

 With dgvEqpList

        If .Rows.Count > 0 Then

           strGrp = .Rows(.CurrentRow.Index).Cells(6).Value 'group
           strEqp = .Rows(.CurrentRow.Index).Cells(7).Value
           strSizeDesc = .Rows(.CurrentRow.Index).Cells(3).Value 'ชุดโมลด์
           strSizeId = .Rows(.CurrentRow.Index).Cells(8).Value

           SeekCodeData(strGrp, strEqp, strSizeDesc, strSizeId)
           ArrangeNumber()
           ShowScrapItem()

           '------------------------------ค้นหารหัสที่เพิ่มเข้าไปใหม่------------------------------------------

           For i = 0 To dgvItem.Rows.Count - 1

               If dgvItem.Rows(i).Cells(4).Value.ToString = strSizeDesc Then
                  dgvItem.CurrentCell = dgvItem.Item(6, i)
                  dgvItem.Focus()
                  Exit For

               End If

           Next i

         StateLockFindDept(True)
         IsShowEqp = False
         gpbEqpList.Visible = False
         txtRemark.Focus()

        End If

  End With

End If

End Sub

Private Sub dgvEqpList_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvEqpList.RowsAdded
    dgvEqpList.Rows(e.RowIndex).Height = 30
End Sub

Private Sub cmbEqp_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbEqp.KeyPress
    If e.KeyChar = Chr(13) Then
       txtEqp.Focus()
    End If
End Sub

Private Sub btnEqpExt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqpExt.Click
  StateLockFindDept(True)
  gpbEqpList.Visible = False
  IsShowEqp = False
End Sub

Private Sub btnEqpAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqpAdd.Click

   Dim strSearch As String

   strSearch = txtEqp.Text.Trim.ToUpper
    If Len(strSearch) <> 0 Then

       FindEqpData(strSearch, cmbEqp.SelectedIndex)
       btnEqpAdd.Focus()
    End If

End Sub

Private Sub txtEqp_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqp.GotFocus
   txtEqp.SelectAll()
End Sub

Private Sub txtEqp_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqp.KeyPress

Dim strSearch As String

    If e.KeyChar = Chr(13) Then

       strSearch = txtEqp.Text.Trim.ToUpper
       If Len(strSearch) <> 0 Then
          FindEqpData(strSearch, cmbEqp.SelectedIndex)
          btnEqpAdd.Focus()
       End If

   End If

End Sub

Private Sub txtEqp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqp.LostFocus

    With txtEqp
         .Text = .Text.ToString.Trim.ToUpper
    End With

End Sub

Private Function SeekCodeData(ByVal strGrp As String, ByVal strEqp As String, ByVal strSizeDesc As String, ByVal strSizeId As String) As Byte

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlSelc As String
Dim strSqlCmd As String

Dim bytReturn As Byte

Dim datSave As Date = Now()
Dim strDate As String = ""

        With Conn

             If .State Then .Close()

                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .Open()

         End With

        strSqlSelc = "SELECT * FROM v_eqp_delivr " _
                             & " WHERE [group] = '" & strGrp & "'" _
                             & " AND eqp_id = '" & strEqp & "'" _
                             & " AND size_desc = '" & strSizeDesc & "'" _
                             & " AND size_id = '" & strSizeId & "'"

         With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                          '-------------------------------------------- วันที่เอกสารตัดจ่าย -----------------------------------------

                          strDate = datSave.Date.ToString("yyyy-MM-dd")
                          strDate = SaveChangeEngYear(strDate)

                          '-------------------------------------------- ลบข้อมูลเก่าออก  ------------------------------------------

                          strSqlCmd = "Delete FROM tmp_delvtrn" _
                                              & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim & "'" _
                                              & " AND [group] = '" & strGrp & "'" _
                                              & " AND eqp_id = '" & strEqp & "'" _
                                              & " AND size_desc = '" & strSizeDesc & "'" _
                                              & " AND size_id = '" & strSizeId & "'"

                          Conn.Execute(strSqlCmd)

                         '--------------- ตัดจำนวนเต็ม แล้วสร้างฟิวด์ใหม่ เพื่อเรียงลำดับ size ใน report size(เช่น 13-14 ให้เป็น 13) ---------------

                          Dim SearchWithinThis As String = .Fields("size_id").Value.ToString.Trim
                          Dim SearchForThis As String = "-"
                          Dim strSizeInt As String = ""
                          Dim FirstCharacter As Integer = SearchWithinThis.IndexOf(SearchForThis)          'หา Index ของ -

                              If FirstCharacter <= 0 Then                                     ' ถ้าเป็น sizeเดี่ยวๆ เช่น 7 ไม่ใช้ 7-9
                                 strSizeInt = SearchWithinThis

                              Else
                                   strSizeInt = Mid(SearchWithinThis, 1, FirstCharacter)      'ตัด - ออก

                              End If

                          '--------------------------  บันทึกข้อมูลในตาราง tmp_delvtrn-(เพิ่ม bob_acpx)  ----------------------------

                          strSqlCmd = "INSERT INTO tmp_delvtrn " _
                                                & "(user_id,doc_id,[no],[group]" _
                                                & ",eqp_id,size_id,size_desc,set_qty,pi,shoe" _
                                                & ",notice,maintn,rmk,int_size" _
                                                & ")" _
                                                & " VALUES (" _
                                                & "'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                                & ",'" & lblDocID.Text.ToString.Trim & "'" _
                                                & "," & "1" _
                                                & ",'" & .Fields("group").Value.ToString.Trim & "'" _
                                                & ",'" & .Fields("eqp_id").Value.ToString.Trim & "'" _
                                                & ",'" & .Fields("size_id").Value.ToString.Trim & "'" _
                                                & ",'" & .Fields("size_desc").Value.ToString.Trim & "'" _
                                                & "," & .Fields("set_qty").Value.ToString.Trim _
                                                & ",'" & "" & "'" _
                                                & ",'" & "" & "'" _
                                                & ",'" & "" & "'" _
                                                & ",'" & "" & "'" _
                                                & ",'" & "" & "'" _
                                                & "," & CInt(strSizeInt) _
                                                & ")"

                         Conn.Execute(strSqlCmd)

                        bytReturn = 1

                  Else
                        bytReturn = 0

                  End If

                 .ActiveConnection = Nothing
                 .Close()

          End With

          Rsd = Nothing

    Conn.Close()
    Conn = Nothing
           SeekCodeData = bytReturn

End Function

Private Sub ShowScrapItem()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String

Dim dubQty As Double
Dim strQty As String

        With Conn

                If .State Then .Close()

                   .ConnectionString = strConnAdodb
                   .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                   .ConnectionTimeout = 90
                   .Open()

         End With

         strSqlCmdSelc = "SELECT * " _
                                        & " FROM v_tmp_delvtrn (NOLOCK)" _
                                        & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                        & " ORDER BY [no]"

         With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                         dgvItem.Rows.Clear()
                         dgvItem.ScrollBars = ScrollBars.None 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

                          If .RecordCount <> 0 Then

                                Do While Not .EOF()

                                        dubQty = .Fields("set_qty").Value - Int(.Fields("set_qty").Value)
                                        If dubQty > 0 Then
                                           strQty = Format(.Fields("set_qty").Value, "#,##0.00")
                                        Else
                                            strQty = Format(.Fields("set_qty").Value, "#,##0")
                                        End If

                                        dgvItem.Rows.Add( _
                                                                     .Fields("eqp_id").Value.ToString.Trim, _
                                                                     .Fields("size_id").Value.ToString.Trim, _
                                                                     .Fields("no").Value, _
                                                                     .Fields("eqp_name").Value.ToString.Trim, _
                                                                     .Fields("size_act").Value.ToString.Trim, _
                                                                     .Fields("size_desc").Value.ToString.Trim, _
                                                                      strQty, _
                                                                     .Fields("pi").Value.ToString.Trim, _
                                                                     .Fields("notice").Value.ToString.Trim, _
                                                                     .Fields("maintn").Value.ToString.Trim, _
                                                                     .Fields("rmk").Value.ToString.Trim _
                                                                     )

                                        .MoveNext()

                                Loop

                          End If

                        .ActiveConnection = Nothing
                        .Close()
                         Rsd = Nothing

                         dgvItem.ScrollBars = ScrollBars.Both 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

        End With

    Conn.Close()
    Conn = Nothing

End Sub

Private Sub dgvItem_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvItem.CellBeginEdit

     With dgvItem

         Select Case e.ColumnIndex

                Case Is = 6 'ราคาทุน
                             strTxt_C6 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.Trim
                            .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "เริ่มแก้ไข"

                Case Is = 7 'จำนวน
                          strTxt_C7 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.Trim

                Case Is = 8 'หมายเหตุ
                          strTxt_C8 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.Trim

                Case Is = 9 'หมายเหตุ
                          strTxt_C9 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.Trim

                Case Is = 10 'หมายเหตุ
                          strTxt_C10 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.Trim

        End Select

   End With

End Sub

Private Sub dgvItem_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItem.CellEndEdit
Dim Conn As New ADODB.Connection

Dim strTmpTxt_C6 As String
Dim strTmpTxt_C7 As String
Dim strTmpTxt_C8 As String
Dim strTmpTxt_C9 As String
Dim strTmpTxt_C10 As String

Dim strMerge As String = ""
Dim strTmp As String

Dim strValue As String
Dim strVal As String

Dim strTableName As String

Dim strEqp As String
Dim strSizeId As String
Dim strSizeDesc As String

Dim dubQty As Double
Dim i, x As Integer

With dgvItem


        If .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim <> "" Then 'ต้องมีรหัสอุปกรณ์

           strEqp = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim 'รหัส
           strSizeId = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim 'Size อุปกรณ์
           strSizeDesc = .Rows(.CurrentRow.Index).Cells(5).Value.ToString.Trim 'Size อุปกรณ์

                       Select Case e.ColumnIndex

                              Case Is = 6 'จำนวน

                                              strTmpTxt_C6 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                              If strTmpTxt_C6 = "" Then
                                                     strVal = "0"
                                              Else
                                                     strVal = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString.Trim
                                              End If

                                              If Not (Microsoft.VisualBasic.Information.IsNumeric(strVal) OrElse strVal.Contains(".")) Then
                                                         .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Val(ChangFormat(strTxt_C6))
                                              Else
                                                          x = Len(strVal)
                                                          For i = 1 To x
                                                                 strTmp = Mid(strVal, i, 1)
                                                                 Select Case strTmp
                                                                               Case Is = ","
                                                                               Case Is = "+"
                                                                               Case Is = "-"
                                                                               Case Is = "_"
                                                                               Case Else
                                                                                           strMerge = strMerge & Trim(strTmp)
                                                                   End Select
                                                                   strTmp = ""

                                                             Next i

                                                             If Val(strMerge) <= 99 Then
                                                                     dubQty = Val(strMerge)
                                                             Else
                                                                     dubQty = 1
                                                              End If

                                                            .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = dubQty

                                                            strTableName = "tmp_delvtrn"
                                                            strValue = "set_qty =" & dubQty.ToString.Trim
                                                            strValue = strValue & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"
                                                            strValue = strValue & " AND eqp_id ='" & strEqp & "'"
                                                            strValue = strValue & " AND size_id ='" & strSizeId & "'"
                                                            strValue = strValue & " AND size_desc ='" & strSizeDesc & "'"
                                                            EditValue(strTableName, strValue)


                                             End If

                              Case Is = 7 'ORDER

                                            strTmpTxt_C7 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                           .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = UCase(Trim(strTmpTxt_C7))

                                            strTableName = "tmp_delvtrn"
                                            strValue = "pi ='" & ReplaceQuote(UCase(Trim(strTmpTxt_C7))) & "'"
                                            strValue = strValue & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"
                                            strValue = strValue & " AND eqp_id ='" & strEqp & "'"
                                            strValue = strValue & " AND size_id ='" & strSizeId & "'"
                                            strValue = strValue & " AND size_desc ='" & strSizeDesc & "'"
                                            EditValue(strTableName, strValue)

                              Case Is = 8 'ข้อควรระวัง

                                            strTmpTxt_C8 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                           .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = UCase(Trim(strTmpTxt_C8))

                                            strTableName = "tmp_delvtrn"
                                            strValue = "notice  ='" & ReplaceQuote(UCase(Trim(strTmpTxt_C8))) & "'"
                                            strValue = strValue & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"
                                            strValue = strValue & " AND eqp_id ='" & strEqp & "'"
                                            strValue = strValue & " AND size_id ='" & strSizeId & "'"
                                            strValue = strValue & " AND size_desc ='" & strSizeDesc & "'"
                                            EditValue(strTableName, strValue)


                              Case Is = 9 'การบำรุงรักษา

                                            strTmpTxt_C9 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                           .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = UCase(Trim(strTmpTxt_C9))

                                            strTableName = "tmp_delvtrn"
                                            strValue = "maintn  ='" & ReplaceQuote(UCase(Trim(strTmpTxt_C9))) & "'"
                                            strValue = strValue & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"
                                            strValue = strValue & " AND eqp_id ='" & strEqp & "'"
                                            strValue = strValue & " AND size_id ='" & strSizeId & "'"
                                            strValue = strValue & " AND size_desc ='" & strSizeDesc & "'"
                                            EditValue(strTableName, strValue)


                              Case Is = 10 'หมายเหตุรายการ

                                            strTmpTxt_C10 = .Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                                           .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = UCase(Trim(strTmpTxt_C10))

                                            strTableName = "tmp_delvtrn"
                                            strValue = "rmk  ='" & ReplaceQuote(UCase(Trim(strTmpTxt_C10))) & "'"
                                            strValue = strValue & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"
                                            strValue = strValue & " AND eqp_id ='" & strEqp & "'"
                                            strValue = strValue & " AND size_id ='" & strSizeId & "'"
                                            strValue = strValue & " AND size_desc ='" & strSizeDesc & "'"
                                            EditValue(strTableName, strValue)

                  End Select

        Else

            Select Case e.ColumnIndex

                    Case Is = 6
                               .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Val(ChangFormat(strTxt_C6))
                    Case Is = 7
                               .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = strTxt_C7
                    Case Is = 8
                               .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = strTxt_C8
                     Case Is = 9
                               .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = strTxt_C9
                     Case Is = 10
                               .Rows(e.RowIndex).Cells(e.ColumnIndex).Value = strTxt_C10

            End Select

        End If

End With

End Sub

Private Sub EditValue(ByVal TableName As String, _
                                        ByVal strValue As String _
                                        )

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String = ""

            With Conn

                    If .State Then .Close()
                        .ConnectionString = strConnAdodb
                        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                        .ConnectionTimeout = 90
                        .Open()

            End With

            Conn.BeginTrans()

            strSqlCmd = "UPDATE " & TableName & " SET " & strValue
            Conn.Execute(strSqlCmd)

            Conn.CommitTrans()

            Conn.Close()
            Conn = Nothing

End Sub

Private Sub dgvItem_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvItem.RowsAdded
    dgvItem.Rows(e.RowIndex).Height = 28
End Sub

Private Sub ArrangeNumber()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim intNo As Integer

Dim strSqlSelc As String

        With Conn

                If .State Then .Close()

                   .ConnectionString = strConnAdodb
                   .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                   .ConnectionTimeout = 90
                   .Open()

         End With

         strSqlSelc = "SELECT * " _
                               & " FROM  tmp_delvtrn (NOLOCK)" _
                               & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                               & " ORDER BY [group],eqp_id,size_id"

         With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlSelc, Conn, , , )

                          If .RecordCount <> 0 Then

                                 intNo = 1

                                Do While Not .EOF()

                                        .Fields("no").Value = intNo.ToString.Trim
                                        .Update()

                                         intNo = intNo + 1

                                        .MoveNext()

                                Loop

                          End If

                        .ActiveConnection = Nothing
                        .Close()
                         Rsd = Nothing

        End With

    Conn.Close()
    Conn = Nothing

End Sub

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
    DeleteSubData()
End Sub

Private Sub DeleteSubData()

Dim btyConsider As Byte

Dim strEqp As String
Dim strSizeId As String
Dim strSizeDesc As String

With dgvItem

        If .Rows.Count > 0 Then

                strEqp = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim 'รหัส
                strSizeId = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim 'Size อุปกรณ์
                strSizeDesc = .Rows(.CurrentRow.Index).Cells(5).Value.ToString.Trim 'กรุ๊ป size

                If strSizeId <> "" Then

                                              btyConsider = MsgBox("รหัสอุปกรณ์ : " & strEqp.ToString.Trim & vbNewLine _
                                                                        & "SIZE : " & strSizeId.ToString.Trim & vbNewLine _
                                                                       & "รายละเอียด : " & strSizeDesc.ToString.Trim & vbNewLine _
                                                                       & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                                        + MsgBoxStyle.Exclamation, "Confirm Delete Data")

                                              If btyConsider = 6 Then

                                                            Dim Conn As New ADODB.Connection
                                                            Dim strSqlCmd As String

                                                            If Conn.State Then Conn.Close()

                                                               Conn.ConnectionString = strConnAdodb
                                                               Conn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                                               Conn.ConnectionTimeout = 90
                                                               Conn.Open()

                                                                strSqlCmd = "DELETE FROM tmp_delvtrn" _
                                                                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                                                     & " AND eqp_id ='" & strEqp & "'" _
                                                                                     & " AND size_id ='" & strSizeId & "'" _
                                                                                     & " AND size_desc ='" & strSizeDesc & "'"

                                                                Conn.Execute(strSqlCmd)

                                                                Conn.Close()

                                                                Conn = Nothing

                                                               .Rows.RemoveAt(.CurrentRow.Index)
                                                                ArrangeNumber()
                                                                ShowScrapItem()

                                                Else
                                                   .Focus()

                                                End If


                End If

        Else
                MsgBox("ไม่มีรายการ SIZE ที่ต้องการลบข้อมูล!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
                dgvItem.Focus()

        End If


End With

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBeforeSave()
End Sub

Private Sub CheckDataBeforeSave()

Dim intListWc As Integer = dgvItem.Rows.Count
Dim strProd As String = ""
Dim strProdNm As String = ""

Dim bytConSave As Byte

   If intListWc > 0 Then 'จำนวนรายการ

      If lblSendId.Text.ToString.Trim <> "" Then 'กำหนดผู้ส่ง

         If lblRvcId.Text.ToString.Trim <> "" Then 'กำหนดผู้รับ

            bytConSave = MsgBox("คุณต้องการบันทึกข้อมูลใช่หรือไม่!" _
                                        , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

            If bytConSave = 6 Then

               Select Case Me.Text

                      Case Is = "เพิ่มข้อมูล"
                           SaveNewRecord()

                      Case Else
                          SaveEditRecord()

               End Select

             Else
                  dgvItem.Focus()

             End If


         Else
               MsgBox("โปรดระบุผู้รับโอน " & vbNewLine _
                                 & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
               ShowPsData2()
         End If

       Else
             MsgBox("โปรดระบุผู้โอน  " & vbNewLine _
                                & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
             ShowPsData1()
       End If

  Else

      MsgBox("โปรดระบุข้อมูลรายการโอนอุปกรณ์ " _
                        & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
      ShowEquipList()

  End If

End Sub

Private Sub SaveNewRecord()

Dim Conn As New ADODB.Connection

Dim strSqlCmd As String

Dim datSave As Date = Now()
Dim strDate As String = ""

Dim strReturnID As String
Dim strDateDoc As String

Dim Rsd As New ADODB.Recordset
Dim RsdSb As New ADODB.Recordset

Dim strSqlSelc As String
Dim strSqlSelcSb As String

Dim strSta As String = "0"

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()
    End With


                   Conn.BeginTrans()

                    '------------------------------------------------บันทึกข้อมูลในตาราง rawtrnh-------------------------------------------------------

                    strReturnID = AutoGenerateID()
                    strDate = datSave.Date.ToString("yyyy-MM-dd")
                    strDate = SaveChangeEngYear(strDate)

                    strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                          & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                          & Mid(txtBegin.Text.ToString, 1, 2)
                    strDateDoc = SaveChangeEngYear(strDateDoc)

                    strSqlCmd = "INSERT INTO delvmst " _
                                          & "(doc_id,doc_date,send_id,send_nm,rvc_id,rvc_nm" _
                                          & ",rvc_dep_nm,pre_by,pre_date,remark" _
                                          & ")" _
                                          & " VALUES (" _
                                          & "'" & strReturnID & "'" _
                                          & ",'" & strDateDoc & "'" _
                                          & ",'" & lblSendId.Text.ToString.Trim & "'" _
                                          & ",'" & lblSendNm.Text.ToString.Trim & "'" _
                                          & ",'" & lblRvcId.Text.ToString.Trim & "'" _
                                          & ",'" & lblRvcNm.Text.ToString.Trim & "'" _
                                          & ",'" & lblRvcDep.Text.ToString.Trim & "'" _
                                          & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                          & ",'" & strDate & "'" _
                                          & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                          & ")"

                      Conn.Execute(strSqlCmd)

                      '---------------------------------------- บันทึกข้อมูลในตาราง delvtrn  ----------------------------------------------------

                       strSqlCmd = "INSERT INTO delvtrn " _
                                           & " SELECT doc_id ='" & strReturnID & "'" _
                                           & ",doc_date ='" & strDateDoc & "'" _
                                           & ",[no],[group],eqp_id,size_id,size_desc,set_qty" _
                                           & ",pi,shoe,notice,maintn,rmk,int_size" _
                                           & "  FROM tmp_delvtrn" _
                                           & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                      Conn.Execute(strSqlCmd)

                     '-------------------------------------- อัพเดทการส่งมอบที่ตาราง eqptrn,อัพเดทสถานะ -----------------------------------------------

                     strSqlSelc = "SELECT [group],eqp_id,size_id,size_desc" _
                                          & " FROM tmp_delvtrn (NOLOCK) " _
                                          & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                          & " GROUP BY [group],eqp_id,size_id,size_desc"

                    Rsd = New ADODB.Recordset
                    Rsd.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                    Rsd.LockType = ADODB.LockTypeEnum.adLockOptimistic
                    Rsd.Open(strSqlSelc, Conn, , , )

                    If Rsd.RecordCount <> 0 Then

                                Do While Not Rsd.EOF

                                         strSqlCmd = "UPDATE eqptrn SET delvr_sta = '1'" _
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

                     '---------------------------------อัพเดทสถานะส่งมอบอุปกรณ์ eqpmst -------------------------------------------------------------------

                     strSqlSelc = "SELECT [group],eqp_id" _
                                          & " FROM tmp_delvtrn (NOLOCK) " _
                                          & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
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

                      Conn.CommitTrans()

                      frmDelv.lblCmd.Text = strReturnID  'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
                      frmDelv.Activating()
                      Me.Close()

    Conn.Close()
    Conn = Nothing

End Sub

Private Function AutoGenerateID() As String

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim datSave As Date = Now()

Dim strSqlCmdSelc As String = ""
Dim strFirstID As String = ""
Dim strLastID As String = ""
Dim strNewID As String = ""

Dim strDate As String = ""
Dim strMonth As String = ""
Dim strYear As String = ""

Dim strNwDocID As String = ""

            With Conn

                    If .State Then .Close()

                        .ConnectionString = strConnAdodb
                        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                        .ConnectionTimeout = 90
                        .Open()

            End With

            strDate = datSave.Date.ToString("yyyy-MM-dd")

            strYear = Mid(strDate, 3, 2)
            strMonth = Mid(strDate, 6, 2)
            strNwDocID = strYear & strMonth

            strSqlCmdSelc = "SELECT TOP 1 doc_id " _
                                            & " FROM delvmst (NOLOCK)" _
                                            & " WHERE SUBSTRING(doc_id,1,6) = 'DV" & strNwDocID.ToString.Trim & "'" _
                                            & " ORDER BY doc_id DESC"

              With Rsd

                   .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                   .LockType = ADODB.LockTypeEnum.adLockOptimistic
                   .Open(strSqlCmdSelc, Conn, , , )

                          If .RecordCount <> 0 Then
                             strFirstID = .Fields("doc_id").Value.ToString.Substring(0, 2)
                             strLastID = .Fields("doc_id").Value.ToString.Substring(6, 3)
                             strNewID = strFirstID & Format(strLastID + 1, strNwDocID & "000")

                          Else
                               strNewID = "DV" & strNwDocID & "001"

                          End If

                        .ActiveConnection = Nothing
                        .Close()

             End With

             Rsd = Nothing
             Conn.Close()
             Conn = Nothing

AutoGenerateID = strNewID

End Function

Private Sub SaveEditRecord()

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String

Dim datSave As Date = Now()
Dim strDate As String = ""

Dim strDateDoc As String

Dim Rsd As New ADODB.Recordset
Dim RsdSb As New ADODB.Recordset

Dim strSqlSelc As String
Dim strSqlSelcSb As String
Dim strSta As String

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()
    End With


                    Conn.BeginTrans()

                    strDate = datSave.Date.ToString("yyyy-MM-dd")
                    strDate = SaveChangeEngYear(strDate)

                    strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                            & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                            & Mid(txtBegin.Text.ToString, 1, 2)
                    strDateDoc = SaveChangeEngYear(strDateDoc)


                    '---------------------------------- อัพเดทการส่งมอบที่ตาราง eqptrn,อัพเดทสถานะ -------------------------------------------------------------------

                     strSqlSelc = "SELECT [group],eqp_id,size_id,size_desc" _
                                          & " FROM delvtrn (NOLOCK) " _
                                          & " WHERE doc_id ='" & lblDocID.Text.ToString.Trim & "'" _
                                          & " GROUP BY [group],eqp_id,size_id,size_desc"

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

                     '-----------------------------------------  อัพเดทสถานะส่งมอบอุปกรณ์ eqpmst ----------------------------------------

                     strSqlSelc = "SELECT [group],eqp_id" _
                                          & " FROM delvtrn (NOLOCK) " _
                                          & " WHERE doc_id ='" & lblDocID.Text.ToString.Trim & "'" _
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

'------------------------------------------------------- จบอัพเดทข้อมูลเก่า ให้ลบทิ้ง-----------------------------------------------------------------------------

                     strSqlCmd = "Delete FROM delvtrn" _
                                          & " WHERE doc_id ='" & lblDocID.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlCmd)

 '------------------------------------------------บันทึกข้อมูลในตาราง delvtrn--------------------------------------------------------------------

                       strSqlCmd = "INSERT INTO delvtrn " _
                                           & " SELECT doc_id ='" & lblDocID.Text.ToString.Trim & "'" _
                                           & ",doc_date ='" & strDateDoc & "'" _
                                           & ",[no],[group],eqp_id,size_id,size_desc,set_qty" _
                                           & ",pi,shoe,notice,maintn,rmk,int_size" _
                                           & " FROM tmp_delvtrn" _
                                           & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                      Conn.Execute(strSqlCmd)

'-------------------------------------------------------------------------------อัพเดทการส่งมอบที่ตาราง eqptrn,อัพเดทสถานะ-----------------------------------------------------------------------------------------------------

                     strSqlSelc = "SELECT [group],eqp_id,size_id,size_desc" _
                                          & " FROM tmp_delvtrn (NOLOCK) " _
                                          & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                          & " GROUP BY [group],eqp_id,size_id,size_desc"

                    Rsd = New ADODB.Recordset
                    Rsd.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                    Rsd.LockType = ADODB.LockTypeEnum.adLockOptimistic
                    Rsd.Open(strSqlSelc, Conn, , , )

                    If Rsd.RecordCount <> 0 Then

                                Do While Not Rsd.EOF

                                         strSqlCmd = "UPDATE eqptrn SET delvr_sta = '1'" _
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

                     '--------------------------------------------------อัพเดทสถานะส่งมอบอุปกรณ์ eqpmst ------------------------------------------------------

                     strSqlSelc = "SELECT [group],eqp_id" _
                                          & " FROM tmp_delvtrn (NOLOCK) " _
                                          & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
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

                     ' --------------------------------อัพเดทหัวเอกสารส่งมอบ----------------------------------------------------------------------------------

                      strSqlCmd = "UPDATE delvmst SET doc_date = '" & strDateDoc & "'" _
                                           & ",send_id = '" & lblSendId.Text.ToString.Trim & "'" _
                                           & ",send_nm = '" & lblSendNm.Text.ToString.Trim & "'" _
                                           & ",rvc_id = '" & lblRvcId.Text.ToString.Trim & "'" _
                                           & ",rvc_nm = '" & lblRvcNm.Text.ToString.Trim & "'" _
                                           & ",rvc_dep_nm = '" & lblRvcDep.Text.ToString.Trim & "'" _
                                           & ",last_by = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                           & ",last_date = '" & strDate & "'" _
                                           & ",remark = '" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                           & " WHERE doc_id ='" & lblDocID.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)
                      Conn.CommitTrans()

                      frmDelv.lblCmd.Text = lblDocID.Text  'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
                      frmDelv.Activating()
                      Me.Close()

    Conn.Close()
    Conn = Nothing

End Sub

End Class
