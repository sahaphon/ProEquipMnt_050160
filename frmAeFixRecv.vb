Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeFixRecv
 Dim IsShowSeek As Boolean        'ตัวเเปรแสดงสถานะ gpbSeek
 Dim strDateDefault As String     'ตัวแปรสำหรับวันที่ทั่วไป

Protected Overrides ReadOnly Property CreateParams() As CreateParams       'ป้องกันการปิดโดยใช้ปุ่ม Close Button(ปุ่มกากบาท)
    Get
         Dim cp As CreateParams = MyBase.CreateParams
         Const CS_DBLCLKS As Int32 = &H8
         Const CS_NOCLOSE As Int32 = &H200
         cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
         Return cp
    End Get

End Property

Private Sub frmAeFixRecv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearTmpTable(0, "")        'ลบข้อมูล Table tmp_fixeqptrn
  frmFixRecv.lblCmd.Text = "0"  'เคลียร์สถานะ
  Me.Dispose()                'ทำลายฟอร์ม คืนหน่วยความจำ
End Sub

Private Sub frmAeFixRecv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Dim dtComputer As Date = Now       'ตัวแปรเก็บค่าวันที่ปัจจุบัน
  Dim strCurrentDate As String       'เก็บค่าสตริงวันที่ปัจจุบัน

      Me.WindowState = FormWindowState.Maximized  'ให้ฟอร์มขยายเต็มหน้าจอ
      StdDateTimeThai()        'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module
      strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

      PrePartSeek()
      ClearDataGpbHead()

      Select Case frmFixRecv.lblCmd.Text.ToString

             Case Is = "0"   'รับเข้าส่งซ่อม
                  LockEditData()
                  CallData()            ' เรียกข้อมูลรายละเอียดส่งซ่อม

             Case Is = "1"   'แก้ไข
                  LockEditData1()
                  CallData1()          ' เรียกข้อมูลรายละเอียดส่งซ่อม

             Case Is = "2"   'มุมมอง
                  LockEditData1()
                  CallData1()          ' เรียกข้อมูลรายละเอียดส่งซ่อม
                  btnSaveData.Enabled = False

       End Select

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  On Error Resume Next    'หากเกิด error โปรแกรมยังจะทำงานต่อไปโดยไม่สนใจ error ที่เกิดขึ้น
  Dim strCode As String

        If MessageBox.Show("ต้องการออกจากฟอร์ม หรือไม่", "กรุณายืนยันออกจากฟอร์ม", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            With frmFixRecv.dgvFix
                If .Rows.Count > 0 Then   'ถ้ามีข้อมูลใน Grid
                    strCode = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim          'ให้strCode = ข้อมูลในแถวปัจจุบัน Cell แรก
                    lblComplete.Text = strCode  'ให้ label แสดงข้อมูลใน Cell ปัจจุบัน   
                End If
            End With
            Me.Close()

            frmMainPro.Show()
            frmFixRecv.Show()
        Else

        End If
End Sub

Private Sub PrePartSeek()
 Dim strEqptype(6) As String
 Dim i As Integer

     strEqptype(0) = "โมล์ดฉีด EVA INJECTION"
     strEqptype(1) = "โมล์ดฉีด PVC INJECTION"
     strEqptype(2) = "โมล์ดหยอด PU"
     strEqptype(3) = "โมล์ดแผงอัดลายหนังหน้า,พื้น"
     strEqptype(4) = "มีดตัด"
     strEqptype(5) = "บล็อกสกรีน"
     strEqptype(6) = "บล็อกอาร์ค"

  With cmbType

       For i = 0 To 6
           .Items.Add(strEqptype(i))
       Next i

 End With
End Sub

Private Sub ClearDataGpbHead()
  txtEqp_id.Text = ""
  txtEqpnm.Text = ""
  txtRemark.Text = ""

End Sub

Private Sub LockEditData()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String = ""                          ' เก็บสตริง sql select

 Dim strSqlCmd As String = ""                           ' เก็บสตริง Command
 Dim blnHavedata As Boolean                             ' เก็บค่าตัวเเปร สำหรับเช็คว่ามีข้อมูลหรือไม่

 Dim strPart As String = ""
 Dim strFixid As String
 Dim strSize As String
 Dim strGpType As String = ""

     strFixid = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(3).Value.ToString
     strSize = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(6).Value.ToString
     strSize = Mid(strSize, 2, 5)        'ต้ด # ออก.
    
     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

     strSqlSelc = "SELECT * " _
                           & "FROM v_fixeqptrn (NOLOCK)" _
                           & " WHERE fix_id = '" & strFixid & "'"

     Rsd = New ADODB.Recordset
     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then

             cmbType.Text = .Fields("desc_thai").Value.ToString.Trim
             txtEqp_id.Text = .Fields("eqp_id").Value.ToString.Trim
             lblFix_id.Text = .Fields("fix_id").Value.ToString.Trim
             txtEqpnm.Text = .Fields("eqp_name").Value.ToString.Trim
             lblAmount.Text = "0"
             lblAmt.Text = "0.00"
             txtRemark.Text = .Fields("remark").Value.ToString.Trim

            '------------------------------- เพิ่มข้อมูลงในตาราง tmp_fixeqptrn ----------------------------

             strSqlSelc = "INSERT INTO tmp_fixeqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE fix_id = '" & strFixid & "'"

             Conn.Execute(strSqlSelc)

             blnHavedata = True                'บ่งบอกว่ามีข้อมูล
             lockGpbHead(False)                'ล็อค Groupbox Head

         Else

             blnHavedata = False

         End If
         .ActiveConnection = Nothing           'คือ ReccordSet
         .Close()

     End With
     Rsd = Nothing
  Conn.Close()
  Conn = Nothing

End Sub

Private Sub LockEditData1()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String = ""                          ' เก็บสตริง sql select

 Dim strSqlCmd As String = ""                           ' เก็บสตริง Command
 Dim blnHavedata As Boolean                             ' เก็บค่าตัวเเปร สำหรับเช็คว่ามีข้อมูลหรือไม่

 Dim strPart As String = ""
 Dim strFixid As String
 Dim strSize As String
 Dim strGpType As String = ""

     strFixid = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(2).Value.ToString
     strSize = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(4).Value.ToString
     strSize = Mid(strSize, 2)        'ต้ด # ออก.
     
     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With


     strSqlSelc = "SELECT * " _
                           & "FROM fixeqpmst (NOLOCK)" _
                           & " WHERE fix_id = '" & strFixid & "'"

     Rsd = New ADODB.Recordset
     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then
             cmbType.Text = checkGroup(.Fields("group").Value.ToString.Trim)
             txtEqp_id.Text = .Fields("eqp_id").Value.ToString.Trim
             lblFix_id.Text = .Fields("fix_id").Value.ToString.Trim
             txtEqpnm.Text = .Fields("eqp_name").Value.ToString.Trim
             txtRemark.Text = .Fields("remark").Value.ToString.Trim

            '------------------------------- เพิ่มข้อมูลงในตาราง tmp_fixeqptrn ----------------------------

             strSqlSelc = "INSERT INTO tmp_fixeqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE fix_id = '" & strFixid & "'" _
                                  & " AND size_id = '" & strSize & "'"

             Conn.Execute(strSqlSelc)

             blnHavedata = True                'บ่งบอกว่ามีข้อมูล
             lockGpbHead(False)                'ล็อค Groupbox Head

         Else
               MsgBox("ไม่มีข้อมู,")
               blnHavedata = False

         End If
         .ActiveConnection = Nothing           'คือ ReccordSet
         .Close()

     End With
     Rsd = Nothing
     Conn.Close()
     Conn = Nothing

End Sub

Private Function checkGroup(ByVal txtGroup As String) As String    'ฟังก์ชั่นแปลง  GroupSize
Dim GroupDesc As String = ""

    Select Case txtGroup

           Case Is = "A"
                 GroupDesc = "โมล์ดฉีด EVA INJECTION"

           Case Is = "B"
                 GroupDesc = "โมล์ดฉีด PVC INJECTION"

           Case Is = "C"
                 GroupDesc = "โมล์ดหยอด PU"

           Case Is = "D"
                 GroupDesc = "โมล์ดแผงอัดลายหนังหน้า,พื้น"

           Case Is = "E"
                 GroupDesc = "มีดตัด"

           Case Is = "F"
                 GroupDesc = "บล็อกสกรีน"

           Case Is = "G"
                 GroupDesc = "บล็อกอาร์ค"

    End Select

    Return GroupDesc

End Function

Private Sub CallData()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim strEqpid As String
 Dim strSize As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

             strEqpid = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(4).Value.ToString.Trim     'รหัสอุปกรณ์
             strSize = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(6).Value.ToString.Trim     'เก็บ Size
             strSize = Mid(strSize, 2)       'ตัด # หน้า size ออก

             strSqlSelc = " SELECT * FROM v_tmp_fixeqptrn (NOLOCK) " _
                                       & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                                       & " AND size_id = '" & strSize & "'"

             Rsd = New ADODB.Recordset

             With Rsd

                  .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                  .LockType = ADODB.LockTypeEnum.adLockOptimistic
                  .Open(strSqlSelc, Conn, , , )

                 If .RecordCount <> 0 Then

                     txtSize.Text = .Fields("size_id").Value.ToString.Trim
                     txtSetQty.Text = .Fields("fix_amount").Value.ToString.Trim
                     txtPr.Text = .Fields("pr_doc").Value.ToString.Trim
                     txtSupp.Text = .Fields("sup_name").Value.ToString.Trim
                     txtFixnm.Text = .Fields("fix_by").Value.ToString.Trim
                     txtIssue.Text = .Fields("issue").Value.ToString.Trim
                     txtRmk.Text = .Fields("fix_rmk").Value.ToString.Trim

                     If .Fields("fix_date").Value.ToString <> "" Then
                        txtFixdate.Text = Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10)
                     Else
                        txtFixdate.Text = "__/__/____"
                     End If


                     If .Fields("due_date").Value.ToString <> "" Then
                        txtDueDate.Text = Mid(.Fields("due_date").Value.ToString.Trim, 1, 10)
                     Else
                        txtDueDate.Text = "__/__/____"
                     End If


                     If .Fields("recv_date").Value.ToString <> "" Then
                        txtRecvDate.Text = Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10)
                     Else
                        txtRecvDate.Text = "__/__/____"
                     End If


                     txtRecvBy.Text = .Fields("recv_by").Value.ToString.Trim
                     txtRecvAmt.Text = .Fields("fix_amount").Value.ToString.Trim
                     txtFxPrice.Text = .Fields("fix_price").Value.ToString.Trim
                     txtFixDetail.Text = .Fields("fix_issue").Value.ToString.Trim

                     lblAmount.Text = txtSetQty.Text
                     lblAmt.Text = txtFxPrice.Text

                     txtRecvDate.Focus()
                     lockFixdetail(False)          'ล็อครายละเอียดส่งซ่อม

                  End If

              .ActiveConnection = Nothing
              .Close()
              End With
              Rsd = Nothing

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub CallData1()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim strEqpid As String
 Dim strSize As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With


             strEqpid = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(3).Value.ToString.Trim     'รหัสอุปกรณ์
             strSize = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(4).Value.ToString.Trim     'เก็บ Size
             strSize = Mid(strSize, 2)       'ตัด # หน้า size ออก


             strSqlSelc = " SELECT * FROM v_tmp_fixeqptrn (NOLOCK) " _
                                       & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                                       & " AND size_id = '" & strSize & "'"

             Rsd = New ADODB.Recordset

             With Rsd

                  .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                  .LockType = ADODB.LockTypeEnum.adLockOptimistic
                  .Open(strSqlSelc, Conn, , , )

                 If .RecordCount <> 0 Then

                     txtSize.Text = .Fields("size_id").Value.ToString.Trim
                     txtSetQty.Text = .Fields("fix_amount").Value.ToString.Trim
                     txtPr.Text = .Fields("pr_doc").Value.ToString.Trim
                     txtSupp.Text = .Fields("sup_name").Value.ToString.Trim
                     txtFixnm.Text = .Fields("fix_by").Value.ToString.Trim
                     txtIssue.Text = .Fields("issue").Value.ToString.Trim
                     txtRmk.Text = .Fields("fix_rmk").Value.ToString.Trim

                     If .Fields("fix_date").Value.ToString <> "" Then
                        txtFixdate.Text = Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10)
                     Else
                        txtFixdate.Text = "__/__/____"
                     End If


                     If .Fields("due_date").Value.ToString <> "" Then
                        txtDueDate.Text = Mid(.Fields("due_date").Value.ToString.Trim, 1, 10)
                     Else
                        txtDueDate.Text = "__/__/____"
                     End If


                     If .Fields("recv_date").Value.ToString <> "" Then
                        txtRecvDate.Text = Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10)
                     Else
                        txtRecvDate.Text = "__/__/____"
                     End If


                     txtRecvBy.Text = .Fields("recv_by").Value.ToString.Trim
                     txtRecvAmt.Text = .Fields("fix_amount").Value.ToString.Trim
                     txtFxPrice.Text = .Fields("fix_price").Value.ToString.Trim
                     txtFixDetail.Text = .Fields("fix_issue").Value.ToString.Trim

                     lblAmount.Text = txtSetQty.Text
                     lblAmt.Text = txtFxPrice.Text

                     txtRecvDate.Focus()
                     lockFixdetail(False)          'ล็อครายละเอียดส่งซ่อม

                  End If

              .ActiveConnection = Nothing
              .Close()
              End With
              Rsd = Nothing

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub lockGpbHead(ByVal sta As Boolean)
  txtEqp_id.Enabled = sta
  txtEqpnm.Enabled = sta
  cmbType.Enabled = sta
  txtRemark.Enabled = sta

End Sub

Private Sub lockFixdetail(ByVal sta As Boolean)
  txtSize.Enabled = sta
  txtSetQty.Enabled = sta
  txtFixdate.Enabled = sta
  txtDueDate.Enabled = sta
  txtPr.Enabled = sta
  txtSupp.Enabled = sta
  txtFixnm.Enabled = sta
  txtIssue.Enabled = sta
  txtRmk.Enabled = sta

End Sub

Private Sub ClearTmpTable(ByVal byOption As Byte, ByVal strPsID As String)

 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

     With Conn

         If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

             Select Case byOption

                    Case Is = 0

                     strSqlCmd = "DELETE tmp_fixeqptrn " _
                                    & "WHERE user_id = '" & frmMainPro.lblLogin.Text & "'"
                     .Execute(strSqlCmd)

                    Case Is = 1

                     strSqlCmd = "DELETE tmp_fixeqptrn " _
                               & "WHERE user_id = '" & frmMainPro.lblLogin.Text & "'" _
                               & "AND docno ='" & strPsID.ToString.Trim & "'"
                    .Execute(strSqlCmd)

              End Select

     End With
     Conn.Close()
     Conn = Nothing

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  ChekDataBfSave()
End Sub

Private Sub ChekDataBfSave()

  If txtRecvDate.Text <> "__/__/____" Then

        If txtRecvBy.Text <> "" Then

                If txtFixDetail.Text <> "" Then

                   'SPrice()          'ฟังก์ช่ันหาผลรวมค่าซ่อมอุปกรณ์
                   SaveEditdata()    'อัพเดทข้อมูล

                Else
                   MsgBox("โปรดระบุรายละเอียดการซ่อม  " _
                                 & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                   txtFixDetail.Focus()

                End If


        Else
            MsgBox("โปรดระบุผู้รับเข้า  " _
                             & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
            txtRecvBy.Focus()

        End If

  Else
      MsgBox("โปรดระบุวันที่รับเข้า  " _
                            & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
      txtRecvDate.Focus()

  End If
End Sub

Private Sub txtRecvDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRecvDate.GotFocus
 With mskRecvDate
      .BringToFront()
      txtRecvDate.SendToBack()
      .Focus()
 End With
End Sub

Private Sub mskRecvDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvDate.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerg As String = ""

        With mskRecvDate
            If txtRecvDate.Text.Trim <> "__/__/____" Then
                x = Len(txtRecvDate.Text)

                For i = 1 To x

                    strTmp = Mid(txtRecvDate.Text.Trim, i, 1)
                    Select Case strTmp
                        Case Is = "_"
                        Case Else
                            If InStr("0123456789/", strTmp) > 0 Then
                                strMerg = strMerg & strTmp
                            End If
                    End Select
                Next i

                Select Case strMerg.ToString.Length    ' Check ความยาวสตริง
                    Case Is = 10
                        .SelectionStart = 0
                    Case Is = 9
                        ' .SelectionStart = 1
                    Case Is = 8
                        ' .SelectionStart = 2
                    Case Is = 7
                        ' .SelectionStart = 3
                    Case Is = 6
                        '.SelectionStart = 4
                    Case Is = 5
                        '.SelectionStart = 5
                    Case Is = 4
                        '.SelectionStart = 6
                End Select
                .SelectedText = strMerg
            End If
            .SelectAll()
        End With
End Sub

Private Sub mskRecvDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskRecvDate.KeyDown
 Dim intChkPoint As Integer
        With mskRecvDate
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                Case Is = 38 'ลูกศรขึ้น
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                        txtRecvBy.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtRecvBy.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtFixDetail.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub mskRecvDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskRecvDate.KeyPress
   If e.KeyChar = Chr(13) Then
     txtRecvBy.Focus()
  End If
End Sub

Private Sub mskRecvDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvDate.LostFocus
 Dim i, x As Integer
 Dim z As Date

 Dim strTmp As String = ""
 Dim strMerge As String = ""

        With mskRecvDate
            x = .Text.Length

            For i = 1 To x
                strTmp = Mid(.Text.ToString.Trim, i, 1)
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
                mskRecvDate.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)

                If Year(z) < 2500 Then
                    txtRecvDate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))

                Else
                    txtRecvDate.Text = z.ToString("dd/MM/yyyy")
                End If
            Catch ex As Exception
                mskRecvDate.Text = "__/__/____"
                txtRecvDate.Text = "__/__/____"

            End Try
            mskRecvDate.SendToBack()
            txtRecvDate.BringToFront()

        End With
End Sub

Private Sub txtRecvBy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRecvBy.KeyDown

 Dim intChkPoint As Integer

     With txtRecvBy

     Select Case e.KeyCode

            Case Is = 35 'ปุ่ม End 

            Case Is = 36 'ปุ่ม Home

            Case Is = 37 'ลูกศรซ้าย

                 If .SelectionStart = 0 Then
                    txtRecvDate.Focus()
                 End If

            Case Is = 38 'ปุ่มลูกศรขึ้น

            Case Is = 39   'ปุ่มลูกศรขวา
                 If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                    txtRecvAmt.Focus()

                 Else

                    intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                           txtRecvAmt.Focus()

                        End If
                 End If

            Case Is = 40 'ปุ่มลง
                      txtFixDetail.Focus()
            Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

     End Select
  End With
End Sub

Private Sub txtRecvAmt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRecvAmt.GotFocus
 With mskRecvAmt
      .BringToFront()
      txtRecvAmt.SendToBack()
      .Focus()
 End With
End Sub

Private Sub mskRecvAmt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvAmt.GotFocus
 Dim i, x As Integer
 Dim strTmp As String = ""
 Dim strMerg As String = ""

        With mskRecvAmt

            If txtRecvAmt.Text.ToString.Trim <> "" Then
                x = Len(txtRecvAmt.Text.ToString)

                For i = 1 To x
                    strTmp = Mid(txtRecvAmt.Text.ToString, i, 1)

                    Select Case strTmp
                           Case Is = "_"
                           Case Else
                            If InStr("0123456789.", strTmp) > 0 Then    'ค้นหาสตริง
                                strMerg = strMerg & strTmp
                            End If
                    End Select

                Next i

                Select Case strMerg.IndexOf(".")

                    Case Is = -1
                        .SelectionStart = 0
                    Case Is = 1
                        .SelectionStart = 1
                    Case Is = 2
                        .SelectionStart = 0
                    Case Is = 3
                        .SelectionStart = 0
                    Case Else
                        .SelectionStart = 0
                End Select
                .SelectedText = strMerg
            End If
            .SelectAll()

        End With
End Sub

Private Sub mskRecvAmt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskRecvAmt.KeyDown
 Dim intChkpoint As Integer

     With mskRecvAmt

         Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                    If .SelectionStart = 0 Then
                        txtRecvBy.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น  

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtFxPrice.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkpoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtFxPrice.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtFixDetail.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
         End Select
        End With
End Sub

Private Sub mskRecvAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskRecvAmt.KeyPress
 If e.KeyChar = Chr(13) Then
    txtFxPrice.Focus()
 End If
End Sub

Private Sub mskRecvAmt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvAmt.LostFocus
 Dim i, x, intFull As Integer
 Dim z As Double

 Dim strTmp As String = ""
 Dim strMerg As String = ""

      With mskRecvAmt
            x = Len(.Text.Length)

            For i = 1 To x
                strTmp = Mid(.Text.ToString, i, 1)
                Select Case strTmp
                    Case Is = "_"
                    Case Else
                        If InStr("0123456789.", strTmp) > 0 Then
                            strMerg = strMerg & strTmp
                        End If
                End Select
                strTmp = ""
            Next i

            Try
                mskRecvAmt.Text = ""     'เคลียร์ mskSizeQty
                z = CDbl(strMerg)        'แปลง Type dbl
                intFull = CInt(z)

                If (z - intFull) > 0 Then
                    txtRecvAmt.Text = z.ToString("#,##0.0")
                Else
                    txtRecvAmt.Text = z.ToString("0")
                End If
            Catch ex As Exception
                txtRecvAmt.Text = "0"
                mskRecvAmt.Text = ""
            End Try

            mskRecvAmt.SendToBack()
            txtRecvAmt.BringToFront()
        End With
End Sub

Private Sub txtFxPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFxPrice.GotFocus
  With mskFxPrice
       .BringToFront()
       txtFxPrice.SendToBack()
       .Focus()
  End With
End Sub

Private Sub txtFxPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFxPrice.KeyPress
  If e.KeyChar = Chr(13) Then
     txtFixDetail.Focus()
  End If
End Sub

Private Sub mskFxPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFxPrice.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerge As String = ""

     With mskFxPrice

           If txtFxPrice.Text <> "0.00" Then

                x = Len(txtFxPrice.Text.ToString)

                For i = 1 To x

                    strTmp = Mid(txtFxPrice.Text.ToString, i, 1)

                    Select Case strTmp
                        Case Is = "_"
                        Case Else

                            If InStr(",.0123456789", strTmp) > 0 Then
                                strMerge = strMerge & strTmp
                            End If

                    End Select
                Next i

                Select Case strMerge.IndexOf(".") 'หาตำแหน่งที่พบเป็นครั้งแรก

                                  Case Is = 7
                                            .SelectionStart = 0
                                  Case Is = 6
                                            .SelectionStart = 1
                                  Case Is = 5
                                            .SelectionStart = 2
                                  Case Is = 3
                                            .SelectionStart = 3
                                  Case Is = 2
                                            .SelectionStart = 5
                                  Case Is = 1
                                            .SelectionStart = 7
                                 Case Else
                                            .SelectionStart = 7
                End Select
                .SelectedText = strMerge

            End If
            .SelectAll()

        End With
End Sub

Private Sub mskFxPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskFxPrice.KeyDown
 Dim intChkPoint As Integer
     With mskFxPrice

       Select Case e.KeyCode
              Case Is = 35 'ปุ่ม End 
              Case Is = 36 'ปุ่ม Home
              Case Is = 37 'ลูกศรซ้าย
                        If .SelectionStart = 0 Then
                            txtRecvAmt.Focus()
                        End If
              Case Is = 38   'ปุ่มลูกศรขึ้น

              Case Is = 39   'ปุ่มลูกศรขวา
                        If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                            txtFixDetail.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                            If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                               txtFixDetail.Focus()
                            End If
                        End If
             Case Is = 40 'ปุ่มลง
                      txtFixDetail.Focus()
             Case Is = 113 'ปุ่ม F2
                      .SelectionStart = .Text.Trim.Length
       End Select

  End With
End Sub

Private Sub mskFxPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskFxPrice.KeyPress
  If e.KeyChar = Chr(13) Then
     txtFixDetail.Focus()
  End If
End Sub

Private Sub mskFxPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFxPrice.LostFocus
  Dim i, x As Integer
  Dim z As Double

  Dim strTmp As String = ""
  Dim strMerge As String = ""

        With mskFxPrice

            x = .Text.Length

            For i = 1 To x

                strTmp = Mid(.Text.ToString, i, 1)
                Select Case strTmp
                    Case Is = ","
                    Case Is = "+"
                    Case Is = "_"
                    Case Else

                        If InStr(".0123456789", strTmp) > 0 Then
                            strMerge = strMerge & strTmp
                        End If

                End Select
                strTmp = ""
            Next i
            Try

                mskFxPrice.Text = ""
                z = CDbl(strMerge)
                txtFxPrice.Text = z.ToString("#,##0.00")


            Catch ex As Exception
                txtFxPrice.Text = "0.00"
                mskFxPrice.Text = ""
            End Try

            mskFxPrice.SendToBack()
            txtFxPrice.BringToFront()

        End With
End Sub

Private Sub txtFixDetail_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFixDetail.KeyDown
 Dim intChkPoint As Integer
 With txtFixDetail

     Select Case e.KeyCode

            Case Is = 35 'ปุ่ม End 

            Case Is = 36 'ปุ่ม Home

            Case Is = 37 'ลูกศรซ้าย
                 If .SelectionStart = 0 Then
                    txtFxPrice.Focus()
                 End If

            Case Is = 38 'ปุ่มลูกศรขึ้น
                 txtRecvBy.Focus()

            Case Is = 39   'ปุ่มลูกศรขวา
                 If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                    btnSaveData.Focus()
                 Else

                    intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                           btnSaveData.Focus()

                        End If
                 End If

            Case Is = 40 'ปุ่มลง

            Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

     End Select
  End With
End Sub

Private Sub txtRecvBy_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRecvBy.KeyPress
   If e.KeyChar = Chr(13) Then
     txtRecvAmt.Focus()
   End If
End Sub

Private Sub SaveEditdata()

 Dim Conn As New ADODB.Connection
 Dim strSqlcmd As String

 Dim DateSave As Date = Now()
 Dim strDate As String = ""
 Dim strDocDate As String
 Dim strRecvDate As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 150
             .Open()

     End With

               'Conn.BeginTrans()      'จุดเริ่มต้น Transection

               strDate = DateSave.Date.ToString("yyyy-MM-dd")
               strDate = SaveChangeEngYear(strDate)

              '------------------------- วันที่เอกสาร ----------------------------------------------------

               strDocDate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocDate = SaveChangeEngYear(strDocDate)

               '---------------------------------------- วดป.รับคืนส่งซ่อม -----------------------------------

               strRecvDate = Mid(txtRecvDate.Text.ToString.Trim, 7, 4) & "-" _
                            & Mid(txtRecvDate.Text.ToString.Trim, 4, 2) & "-" _
                            & Mid(txtRecvDate.Text.ToString.Trim, 1, 2)
               strRecvDate = SaveChangeEngYear(strRecvDate)


                     strSqlcmd = "UPDATE tmp_fixeqptrn SET recv_date = '" & strRecvDate & "'" _
                                             & "," & "fix_sta = '2' " _
                                             & "," & "recv_by = '" & ReplaceQuote(txtRecvBy.Text.ToString.Trim) & "'" _
                                             & "," & "fix_price = " & ChangFormat(txtFxPrice.Text.ToString.Trim) _
                                             & "," & "fix_issue = '" & ReplaceQuote(txtFixDetail.Text.ToString.Trim) & "'" _
                                             & " WHERE  user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                             & " AND fix_id = '" & lblFix_id.Text.ToString.Trim & "'" _
                                             & " AND size_id = '" & txtSize.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)


                     '--------------------------------------- ลบข้อมูลในตาราง eqptrn -----------------------------------

                     strSqlcmd = "DELETE FROM fixeqptrn" _
                                           & " WHERE fix_id ='" & lblFix_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)

                     '--------------------------------------- เพิ่มข้อมูลในตาราง fixEqptrn -------------------------------

                     strSqlcmd = "INSERT INTO fixeqptrn " _
                                     & " SELECT fix_sta " _
                                     & ",fix_id = '" & lblFix_id.Text.ToUpper.Trim & "'" _
                                     & ",[group],eqp_id,size_id" _
                                     & ",fix_amount,fix_price,fix_date,fix_by,pr_doc" _
                                     & ",issue,fix_issue,sup_name,due_date,recv_date" _
                                     & ",recv_by,fix_rmk" _
                                     & " FROM tmp_fixeqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"

                     Conn.Execute(strSqlcmd)
                     'Conn.CommitTrans()                   'สั่ง Commit transection

                     lblComplete.Text = txtSize.Text.ToString.Trim  'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
                     Me.Hide()

                           If ChkFixStatus() Then         'ตรวจสอบว่ารับเข้าส่งซ่อมครบหรือไม่
                              SaveEditFixmst()            'อัพเดทข้อมูล table fixeqpmst

                           End If

                     frmMainPro.Show()
                     frmFixRecv.Show()


  Conn.Close()
  Conn = Nothing

End Sub

Private Function ChkFixStatus() As Boolean     'ตรวจสอบสถานะการส่งซ่อม
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 150
             .Open()

     End With

      '------------------- ถ้ารับเข้าส่งซ่อม ครบทุก size -----------------------

      strSqlSelc = "SELECT * " _
                        & " FROM tmp_fixeqptrn (NOLOCK)" _
                        & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "' " _
                        & " AND fix_sta = '1'"

      Rsd = New ADODB.Recordset

      With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(strSqlSelc, Conn, , , )

           If .RecordCount <> 0 Then         'ถ้ายังรับเข้าไม่ครบ (fix_sta = 1)
              Return False

           Else                              'กรณีรับเข้า ครบถ้วน
              Return True

           End If

      End With
      Rsd = Nothing
      Rsd.Close()


  Conn = Nothing
  Conn.Close()

End Function

Private Sub SaveEditFixmst()    'อัพเดทสถานะรับเข้าส่งซ่อม tb_fixeqpmst

 Dim Conn As New ADODB.Connection
 Dim strSqlcmd As String

 Dim DateSave As Date = Now()
 Dim strDate As String = ""
 Dim strDocDate As String
 Dim strRecvDate As String
 Dim strGpType As String = ""

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 150
             .Open()

     End With

            Conn.BeginTrans()      'จุดเริ่มต้น Transection

            strDate = DateSave.Date.ToString("yyyy-MM-dd")
            strDate = SaveChangeEngYear(strDate)

            '------------------------- วันที่เอกสาร ----------------------------------------------------

            strDocDate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
            strDocDate = SaveChangeEngYear(strDocDate)

            '---------------------------------------- วดป.รับคืนส่งซ่อม -----------------------------------

            strRecvDate = Mid(txtRecvDate.Text.ToString.Trim, 7, 4) & "-" _
                            & Mid(txtRecvDate.Text.ToString.Trim, 4, 2) & "-" _
                            & Mid(txtRecvDate.Text.ToString.Trim, 1, 2)
            strRecvDate = SaveChangeEngYear(strRecvDate)


            Select Case cmbType.Text

                   Case Is = "โมล์ดฉีด EVA INJECTION"
                           strGpType = "A"

                   Case Is = "โมล์ดฉีด PVC INJECTION"
                           strGpType = "B"

                   Case Is = "โมล์ดหยอด PU"
                           strGpType = "C"

                   Case Is = "โมล์ดแผงอัดลายหนังหน้า,พื้น"
                           strGpType = "D"

                   Case Is = "มีดตัด"
                           strGpType = "E"

                   Case Is = "บล็อกสกรีน"
                           strGpType = "F"

                   Case Is = "บล็อกอาร์ค"
                           strGpType = "G"

             End Select

                     strSqlcmd = " UPDATE fixeqpmst SET fix_sta= '2'" _
                                            & "," & "[group] ='" & strGpType & "'" _
                                            & "," & "eqp_id ='" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                            & "," & "eqp_name ='" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                            & "," & "fix_amount ='" & ReplaceQuote(lblAmount.Text.ToString.Trim) & "'" _
                                            & "," & "fix_price = " & ChangFormat(SPrice) _
                                            & "," & "remark ='" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                            & "," & "last_date = '" & strDate & "'" _
                                            & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & "," & "pic_bf = '" & "" & "'" _
                                            & "," & "pic_af = '" & "" & "'" _
                                            & "," & "pro_sta ='" & "0" & "'" _
                                            & " WHERE fix_id ='" & lblFix_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)

                     '--------------- อัพเดทสถานะ eqpmst เป็น  2(รับกลับส่งซ่อม)-------------------------------------------------

                     strSqlcmd = " UPDATE eqpmst SET fix_sta = '2'" _
                                            & "," & "last_date = '" & strDate & "'" _
                                            & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)

                     '---------------- อัพเดทสถานะส่่งซ่อม eqptrn fix_sta = 2 (รับกลับส่งซ่อม) ------------------------------------------

                     strSqlcmd = " UPDATE eqptrn SET fix_sta = '2'" _
                                            & "," & "last_date = '" & strDate & "'" _
                                            & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'" _
                                            & " AND size_id = '" & txtSize.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)

                     Conn.CommitTrans()               'สั่ง Commit transection


   Conn.Close()
   Conn = Nothing
End Sub

Private Function SPrice() As Double        'ฟังก์ช่ันหาผลรวมค่าซ่อมอุปกรณ์
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim price As String

     With Conn
          .ConnectionString = strConnAdodb
          .CursorLocation = ADODB.CursorLocationEnum.adUseClient
          .ConnectionTimeout = 150
          .Open()
     End With

       strSqlSelc = " SELECT SUM(fix_price)as SumPrice FROM v_tmp_fixeqptrn (NOLOCK) " _
                              & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

       With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            price = .Fields("SumPrice").Value.ToString

            Return price

       .ActiveConnection = Nothing
       .Close()
       End With

 Conn.Close()
 Conn = Nothing
End Function

Private Sub txtFixnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFixnm.KeyPress
  If e.KeyChar = Chr(13) Then
     txtIssue.Focus()
  End If
End Sub

Private Sub txtIssue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIssue.KeyPress
  If e.KeyChar = Chr(13) Then
     txtDueDate.Focus()
  End If
End Sub
End Class