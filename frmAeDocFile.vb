Imports ADODB

Public Class frmAeDocFile

Protected Overrides ReadOnly Property CreateParams() As CreateParams 'ป้องกันการปิดโดยใช้ปุ่ม Close Button
Get
Dim cp As CreateParams = MyBase.CreateParams
Const CS_DBLCLKS As Int32 = &H8
Const CS_NOCLOSE As Int32 = &H200
cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
Return cp
End Get
End Property

Private Sub frmAeDocFile_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    Me.Dispose()
End Sub


Private Sub frmAeDocFile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

Me.WindowState = FormWindowState.Maximized
StdDateTimeThai()
PreGroupType()

Select Case frmDocFile.lblCmd.Text.ToString
          Case Is = "0" 'เพิ่มข้อมูล
                  ShowUserDefault()
          Case Is = "1" 'แก้ไขข้อมูล
                  LockEditData()
          Case Is = "2" 'มุมมองข้อมูล
                  LockEditData()
                  btnSaveData.Enabled = False
 End Select

End Sub

Private Sub ShowUserDefault()

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

         strSqlCmdSelc = "SELECT * FROM usermst (NOLOCK)" _
                              & " ORDER BY user_id"

         Rsd = New ADODB.Recordset
         With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                         dgvPs.Rows.Clear()

                          If .RecordCount <> 0 Then


                                Do While Not .EOF()

                                        dgvPs.Rows.Add( _
                                                                  IIf(.Fields("act_usr").Value.ToString.Trim = "A", My.Resources.admin, My.Resources.users), _
                                                                  .Fields("act_usr").Value.ToString, _
                                                                  .Fields("user_id").Value.ToString.Trim, _
                                                                  .Fields("sname").Value.ToString.Trim, _
                                                                   "", 0, "", "", _
                                                                    0, 0, 0, 0, 0, 0, 0, 0, _
                                                                    My.Resources.lock_red, My.Resources.lock_red, _
                                                                    My.Resources.lock_red, My.Resources.lock_red, _
                                                                    My.Resources.lock_red, My.Resources.lock_red, _
                                                                    My.Resources.lock_red, My.Resources.lock_red _
                                                                    , 0, 0, 0, 0, 0, 0, _
                                                                    My.Resources.lock_red, My.Resources.lock_red, _
                                                                    My.Resources.lock_red, My.Resources.lock_red, _
                                                                    My.Resources.lock_red, My.Resources.lock_red _
                                                                 )
                                        .MoveNext()
                                Loop

                                lblPsQty.Text = .RecordCount.ToString

                          Else

                                lblPsQty.Text = "0"

                          End If

                        .ActiveConnection = Nothing
                        .Close()
                         Rsd = Nothing

        End With

    Conn.Close()
    Conn = Nothing

End Sub


Private Sub LockEditData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String
Dim strDateEdit As String
Dim strDocFile As String
Dim strDocFileName As String

        strDocFile = frmDocFile.dgvShoe.Rows(frmDocFile.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim
        strDocFileName = frmDocFile.dgvShoe.Rows(frmDocFile.dgvShoe.CurrentRow.Index).Cells(1).Value.ToString.Trim

        txtDept.Text = strDocFile
        txtDeptName.Text = strDocFileName

        txtDept.ReadOnly = True

        With Conn

                If .State Then .Close()
                   .ConnectionString = strConnAdodb
                   .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                   .ConnectionTimeout = 90
                   .Open()

         End With

         strSqlCmdSelc = "SELECT * FROM v_usr_permit (NOLOCK)" _
                              & " WHERE file_icon ='" & strDocFile & "'" _
                              & " ORDER BY user_id"

         Rsd = New ADODB.Recordset
         With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                         dgvPs.Rows.Clear()

                         If .RecordCount <> 0 Then


                                Do While Not .EOF()

                                        strDateEdit = Mid(.Fields("last_date").Value.ToString, 1, 10)

                                        dgvPs.Rows.Add( _
                                                                  IIf(.Fields("act_usr").Value.ToString.Trim = "A", My.Resources.admin, My.Resources.users), _
                                                                  .Fields("act_usr").Value.ToString, _
                                                                  .Fields("user_id").Value.ToString.Trim, _
                                                                  .Fields("sname").Value.ToString.Trim, _
                                                                  .Fields("file_icon").Value.ToString.Trim, _
                                                                  .Fields("open_cnt").Value, strDateEdit, _
                                                                  .Fields("last_time").Value.ToString.Trim, _
                                                                   IIf(.Fields("act_open").Value, 1, 0), _
                                                                   IIf(.Fields("act_view").Value, 1, 0), _
                                                                   IIf(.Fields("act_add").Value, 1, 0), _
                                                                   IIf(.Fields("act_edit").Value, 1, 0), _
                                                                   IIf(.Fields("act_delete").Value, 1, 0), _
                                                                   IIf(.Fields("act_copy").Value, 1, 0), _
                                                                   IIf(.Fields("act_print").Value, 1, 0), _
                                                                   IIf(.Fields("act_other").Value, 1, 0), _
                                                                   IIf(.Fields("act_open").Value, My.Resources.unlock_green, My.Resources.lock_red), _
                                                                   IIf(.Fields("act_view").Value, My.Resources.unlock_green, My.Resources.lock_red), _
                                                                   IIf(.Fields("act_add").Value, My.Resources.unlock_green, My.Resources.lock_red), _
                                                                   IIf(.Fields("act_edit").Value, My.Resources.unlock_green, My.Resources.lock_red), _
                                                                   IIf(.Fields("act_delete").Value, My.Resources.unlock_green, My.Resources.lock_red), _
                                                                   IIf(.Fields("act_copy").Value, My.Resources.unlock_green, My.Resources.lock_red), _
                                                                   IIf(.Fields("act_print").Value, My.Resources.unlock_green, My.Resources.lock_red), _
                                                                   IIf(.Fields("act_other").Value, My.Resources.unlock_green, My.Resources.lock_red) _
                                                                 )
                                        .MoveNext()
                                Loop

                                lblPsQty.Text = .RecordCount.ToString

                          Else

                                lblPsQty.Text = "0"

                          End If

                        .ActiveConnection = Nothing
                        .Close()
                         Rsd = Nothing

        End With

    Conn.Close()
    Conn = Nothing

    CheckStatusCheckBox()

End Sub


Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

  Me.Hide()
  frmMainpro.Show()
  frmDocFile.Show()

End Sub

Private Sub dgvPs_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPs.CellDoubleClick

With dgvPs

        If .Rows.Count > 0 Then

                Select Case .CurrentCell.ColumnIndex

                          Case Is = 16 'สถานะ Open

                                    If .Rows(.CurrentRow.Index).Cells(8).Value = 0 Then
                                         .Item(8, .CurrentRow.Index).Value = 1
                                         .Item(16, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(8, .CurrentRow.Index).Value = 0
                                         .Item(16, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                          Case Is = 17 'สถานะ View

                                    If .Rows(.CurrentRow.Index).Cells(9).Value = 0 Then
                                         .Item(9, .CurrentRow.Index).Value = 1
                                         .Item(17, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(9, .CurrentRow.Index).Value = 0
                                         .Item(17, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                         Case Is = 18 'สถานะ Add

                                    If .Rows(.CurrentRow.Index).Cells(10).Value = 0 Then
                                         .Item(10, .CurrentRow.Index).Value = 1
                                         .Item(18, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(10, .CurrentRow.Index).Value = 0
                                         .Item(18, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 19 'สถานะ Edit

                                    If .Rows(.CurrentRow.Index).Cells(11).Value = 0 Then
                                         .Item(11, .CurrentRow.Index).Value = 1
                                         .Item(19, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(11, .CurrentRow.Index).Value = 0
                                         .Item(19, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 20 'สถานะ Del

                                    If .Rows(.CurrentRow.Index).Cells(12).Value = 0 Then
                                         .Item(12, .CurrentRow.Index).Value = 1
                                         .Item(20, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(12, .CurrentRow.Index).Value = 0
                                         .Item(20, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 21 'สถานะ Copy

                                    If .Rows(.CurrentRow.Index).Cells(13).Value = 0 Then
                                         .Item(13, .CurrentRow.Index).Value = 1
                                         .Item(21, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(13, .CurrentRow.Index).Value = 0
                                         .Item(21, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 22 'สถานะ Print

                                    If .Rows(.CurrentRow.Index).Cells(14).Value = 0 Then
                                         .Item(14, .CurrentRow.Index).Value = 1
                                         .Item(22, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(14, .CurrentRow.Index).Value = 0
                                         .Item(22, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 23 'สถานะ Other

                                    If .Rows(.CurrentRow.Index).Cells(15).Value = 0 Then
                                         .Item(15, .CurrentRow.Index).Value = 1
                                         .Item(23, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(15, .CurrentRow.Index).Value = 0
                                         .Item(23, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If


                End Select

                CheckStatusCheckBox()

        End If

End With

End Sub

Private Sub dgvPs_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvPs.MouseMove

Dim objMousePosition As Point = dgvPs.PointToClient(Control.MousePosition)
Dim objHitTestInfo As DataGridView.HitTestInfo
      objHitTestInfo = dgvPs.HitTest(objMousePosition.X, objMousePosition.Y)

With dgvPs

        Select Case objHitTestInfo.ColumnIndex
                  Case 16, 17, 18, 19, 20, 21, 22, 23
                         .Cursor = Cursors.Hand
                  Case Else
                        .Cursor = Cursors.Default
        End Select

End With
End Sub


Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

  If dgvPs.Rows.Count > 0 Then

      With gpbSearch

                 .Top = 230
                 .Left = 300
                 .Width = 311
                 .Height = 125

                 .Visible = True

                 cmbType.SelectedItem = cmbType.Items(0)
                 txtSeek.Text = _
                 dgvPs.Rows(dgvPs.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                 StateLockFind(False)
                 txtSeek.Focus()

       End With

    End If

End Sub

Private Sub PreGroupType()

Dim strGpTopic(1) As String
Dim i As Byte

      strGpTopic(0) = "User LogIn"
      strGpTopic(1) = "ชื่อ-นามสกุล"


      With cmbType

              For i = 0 To 1
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(0)

      End With

End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

  StateLockFind(True)
  gpbSearch.Visible = False

End Sub

Private Sub StateLockFind(ByVal Sta As Boolean)

    gpbWc.Enabled = Sta
    gpbSub.Enabled = Sta
    btnSaveData.Enabled = Sta
    btnSearch.Enabled = Sta

End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
    FindData()
End Sub

Private Sub FindData()

Dim i, x As Integer
Dim z As Boolean

Dim strSearchTxt As String


    Select Case cmbType.SelectedIndex()
              Case Is = 0 'รหัสระบบงาน                        
                       x = 2
              Case Is = 1 'ชื่อระบบงาน
                       x = 3
    End Select

    strSearchTxt = txtSeek.Text.ToString.Trim.ToUpper

    With dgvPs


            For i = 0 To .Rows.Count - 1

                    If InStr(UCase(.Rows(i).Cells(x).Value), strSearchTxt) <> 0 Then
                        .CurrentCell = .Item(x, i)
                        .Focus()
                         z = True
                        Exit For
                    End If

            Next i

    End With

    If Not z Then
        MsgBox("ข้อมูล : " & strSearchTxt & " ไม่มีอยู่ในระบบ!!" & vbNewLine _
                     & "โปรดระบุข้อมูลใหม่", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Not Found Data!!")
    End If

StateLockFind(True)
gpbSearch.Visible = False
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
    FindData()
End If

End Sub

Private Sub dgvPs_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvPs.KeyDown

If e.KeyCode = Keys.Enter Then
    e.Handled = True
End If

End Sub

Private Sub txtDept_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDept.GotFocus
    txtDept.SelectAll()
End Sub

Private Sub txtDept_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDept.KeyDown
Dim intChkPoint As Integer

    With txtDept

            Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย
                               If .SelectionStart = 0 Then
                                    txtDeptName.Focus()
                               End If
                      Case Is = 38 'ปุ่มลูกศรขึ้น                                
                      Case Is = 39 'ปุ่มลูกศรขวา
                                If .SelectionLength = .Text.Trim.Length Then
                                    txtDeptName.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtDeptName.Focus()
                                    End If
                                End If
                      Case Is = 40
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtDept_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDept.KeyPress

    If e.KeyChar = Chr(13) Then
        txtDeptName.Focus()
    End If

End Sub

Private Sub txtDept_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDept.LostFocus

With txtDept
     .Text = .Text.ToString.Trim.ToUpper
End With

End Sub

Private Sub txtDeptName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDeptName.GotFocus
    txtDeptName.SelectAll()
End Sub

Private Sub txtDeptName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDeptName.KeyDown
Dim intChkPoint As Integer

    With txtDeptName

            Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย
                               If .SelectionStart = 0 Then
                                    txtDept.Focus()
                               End If
                      Case Is = 38 'ปุ่มลูกศรขึ้น                                
                      Case Is = 39 'ปุ่มลูกศรขวา
                                If .SelectionLength = .Text.Trim.Length Then
                                    txtDept.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtDept.Focus()
                                    End If
                                End If
                      Case Is = 40
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub txtDeptName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDeptName.KeyPress

    If e.KeyChar = Chr(13) Then
        btnSaveData.Focus()
    End If

End Sub

Private Sub txtDeptName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDeptName.LostFocus

    With txtDeptName
          .Text = .Text.ToString.Trim.ToUpper
    End With

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBeforeSave()
End Sub

Private Sub CheckDataBeforeSave()

Dim bytConSave As Byte

    If txtDept.Text.Trim.ToString <> "" Then 'รหัสแฟ้ม

            If txtDeptName.Text.Trim.ToString <> "" Then 'ชื่อแฟ้มระบบ


                                        bytConSave = MsgBox("คุณต้องการบันทึกข้อมูลใช่หรือไม่!" _
                                                                        , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                        If bytConSave = 6 Then
                                                    Select Case Me.Text
                                                              Case Is = "เพิ่มข้อมูลแฟ้มระบบ"
                                                                      SaveNewRecord()
                                                                Case Is = "แก้ไขข้อมูลแฟ้มระบบ"
                                                                      SaveEditRecord()
                                                                Case Is = "คัดลอกข้อมูล"
                                                                      'SaveNewRecord()
                                                                Case Else
                                                     End Select
                                        Else
                                            txtDept.Focus()
                                        End If


            Else


                MsgBox("โปรดระบุชื่อแฟ้มข้อมูล ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                txtDeptName.Focus()

            End If

    Else

        MsgBox("โปรดระบุรหัสแฟ้มข้อมูล ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
        txtDept.Focus()

    End If

End Sub

Private Sub SaveNewRecord()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String
Dim strSqlCmd As String

Dim strDept As String = txtDept.Text.ToString.Trim

Dim datSave As Date = Now()
Dim strDate As String = ""

Dim bytHaveData As Byte

Dim i, z As Integer

        With Conn

                If .State Then .Close()
                        .ConnectionString = strConnAdodb
                        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                        .ConnectionTimeout = 90
                        .Open()

        End With

        strSqlCmdSelc = "SELECT file_icon FROM filemst (NOLOCK)" _
                              & " WHERE file_icon ='" & strDept & "'"

         Rsd = New ADODB.Recordset
         With Rsd

                 .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                 .LockType = ADODB.LockTypeEnum.adLockOptimistic
                 .Open(strSqlCmdSelc, Conn, , , )

                  If .RecordCount <> 0 Then

                          MsgBox("ข้อมูลรหัสแฟ้มระบบ : " & strDept & " มีอยู่แล้ว!" & vbNewLine _
                                      & "โปรดระบุข้อมูลใหม่", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Duplication Data!!")
                          txtDept.Focus()
                          bytHaveData = 0

                  Else
                          bytHaveData = 1
                  End If

                 .ActiveConnection = Nothing
                 .Close()

          End With
          Rsd = Nothing

          If bytHaveData = 1 Then


                    Conn.BeginTrans()

                    strDate = datSave.Date.ToString("yyyy-MM-dd")
                    strDate = SaveChangeEngYear(strDate)

                    '------------------------------------------------เพิ่มรหัสแฟ้มระบบ -----------------------------------------------------------------                                
                    strSqlCmd = "INSERT INTO filemst " _
                                   & "(file_icon,file_name,pre_by,pre_date)" _
                                   & " VALUES (" _
                                   & "'" & ReplaceQuote(strDept) & "'" _
                                   & ",'" & ReplaceQuote(txtDeptName.Text.ToString.Trim) & "'" _
                                   & ",'" & frmMainpro.lblLogin.Text.Trim.ToString & "'" _
                                   & ",'" & strDate & "'" _
                                   & ")"
                    Conn.Execute(strSqlCmd)

                    '------------------------------------------------เพิ่มรหัสแฟ้มระบบ ในแต่ละUser-----------------------------------------------------------------                                
                    With dgvPs
                            z = .Rows.Count - 1

                            If z > 0 Then

                                For i = 0 To z

                                    strSqlCmd = "INSERT INTO usertrn " _
                                                    & "(user_id,file_icon," _
                                                    & "act_open,act_view,act_add,act_edit,act_delete,act_copy,act_print,act_other" _
                                                    & ")" _
                                                    & " VALUES (" _
                                                    & "'" & .Rows(i).Cells(2).Value & "'" _
                                                    & ",'" & ReplaceQuote(strDept) & "'" _
                                                    & "," & .Rows(i).Cells(8).Value _
                                                    & "," & .Rows(i).Cells(9).Value _
                                                    & "," & .Rows(i).Cells(10).Value _
                                                    & "," & .Rows(i).Cells(11).Value _
                                                    & "," & .Rows(i).Cells(12).Value _
                                                    & "," & .Rows(i).Cells(13).Value _
                                                    & "," & .Rows(i).Cells(14).Value _
                                                    & "," & .Rows(i).Cells(15).Value _
                                                    & ")"
                                    Conn.Execute(strSqlCmd)

                                Next i

                            End If

                    End With


                    Conn.CommitTrans()

                    lblComplete.Text = strDept  'บ่งบอกว่าบันทึกข้อมูลสำเร็จ

                    Me.Hide()
                    frmMainpro.Show()
                    frmDocFile.Show()


          End If

          Conn.Close()
          Conn = Nothing

End Sub


Private Sub SaveEditRecord()

Dim Conn As New ADODB.Connection

Dim strSqlCmd As String
Dim strDept As String = txtDept.Text.ToString.Trim

Dim datSave As Date = Now()
Dim strDate As String = ""
Dim strDateEntry As String = ""

Dim i, z As Integer


        With Conn

                If .State Then .Close()
                        .ConnectionString = strConnAdodb
                        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                        .ConnectionTimeout = 90
                        .Open()

        End With


                    Conn.BeginTrans()

                    '------------------------------------------------บันทึกข้อมูลในตาราง filemst-------------------------------------------------------                    

                      strDate = datSave.Date.ToString("yyyy-MM-dd")
                      strDate = SaveChangeEngYear(strDate)

                      strSqlCmd = "UPDATE filemst SET file_name ='" & ReplaceQuote(txtDeptName.Text.ToString.Trim) & "'" _
                                   & "," & "last_date ='" & strDate & "'" _
                                   & "," & "last_by ='" & frmMainpro.lblLogin.Text.Trim.ToString & "'" _
                                   & " WHERE file_icon ='" & strDept & "'"
                     Conn.Execute(strSqlCmd)


                    '------------------------------------------------ลบข้อมูลในตาราง usertrn----------------------------------------------------
                     strSqlCmd = "Delete FROM usertrn" _
                                    & " WHERE file_icon ='" & strDept & "'"
                     Conn.Execute(strSqlCmd)


                    '------------------------------------------------เพิ่มรหัสแฟ้มระบบ ในแต่ละUser-----------------------------------------------------------------                                
                    With dgvPs
                            z = .Rows.Count - 1

                            If z > 0 Then

                                For i = 0 To z

                                    If .Rows(i).Cells(6).Value.ToString.Trim <> "" Then
                                            strDateEntry = Mid(.Rows(i).Cells(6).Value.ToString, 7, 4) & "-" _
                                                               & Mid(.Rows(i).Cells(6).Value.ToString, 4, 2) & "-" _
                                                               & Mid(.Rows(i).Cells(6).Value.ToString, 1, 2)
                                    Else
                                            strDateEntry = ""
                                    End If

                                    strSqlCmd = "INSERT INTO usertrn " _
                                                    & "(user_id,file_icon," _
                                                    & "act_open,act_view,act_add,act_edit,act_delete,act_copy,act_print,act_other," _
                                                    & "open_cnt,last_date,last_time" _
                                                    & ")" _
                                                    & " VALUES (" _
                                                    & "'" & .Rows(i).Cells(2).Value & "'" _
                                                    & ",'" & ReplaceQuote(strDept) & "'" _
                                                    & "," & .Rows(i).Cells(8).Value _
                                                    & "," & .Rows(i).Cells(9).Value _
                                                    & "," & .Rows(i).Cells(10).Value _
                                                    & "," & .Rows(i).Cells(11).Value _
                                                    & "," & .Rows(i).Cells(12).Value _
                                                    & "," & .Rows(i).Cells(13).Value _
                                                    & "," & .Rows(i).Cells(14).Value _
                                                    & "," & .Rows(i).Cells(15).Value _
                                                    & "," & IIf(.Rows(i).Cells(5).Value.ToString.Trim <> "", .Rows(i).Cells(5).Value, "0") _
                                                    & ",'" & strDateEntry & "'" _
                                                    & ",'" & .Rows(i).Cells(7).Value.ToString & "'" _
                                                    & ")"
                                    Conn.Execute(strSqlCmd)

                                Next i

                            End If

                    End With

                    Conn.CommitTrans()

                    lblComplete.Text = strDept  'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
                    Me.Hide()
                    frmMainpro.Show()
                    frmDocFile.Show()


          Conn.Close()
          Conn = Nothing

End Sub


Private Sub CheckStatusCheckBox()
Dim i, x, z As Integer

With dgvPs

            If .Rows.Count > 0 Then

                    For i = 8 To 15 'column


                        '-------------------------------------------------ลูป Perimission---------------------------------------------
                        For z = 0 To .Rows.Count - 1 'row
                             If .Item(i, z).Value = 1 Then
                                    x = x + 1
                             End If
                        Next z

                        Select Case i
                                  Case Is = 8 'CheckBox Open

                                            If z = x Then
                                               chkOpen.Checked = True
                                            Else
                                               chkOpen.Checked = False
                                            End If

                                  Case Is = 9 'CheckBox View

                                            If z = x Then
                                               chkView.Checked = True
                                            Else
                                               chkView.Checked = False
                                            End If

                                  Case Is = 10 'CheckBox Add

                                            If z = x Then
                                               chkAdd.Checked = True
                                            Else
                                               chkAdd.Checked = False
                                            End If

                                  Case Is = 11 'CheckBox Edit

                                            If z = x Then
                                               chkEdit.Checked = True
                                            Else
                                               chkEdit.Checked = False
                                            End If

                                  Case Is = 12 'CheckBox Del

                                            If z = x Then
                                               chkDel.Checked = True
                                            Else
                                               chkDel.Checked = False
                                            End If

                                  Case Is = 13 'CheckBox Copy

                                            If z = x Then
                                               chkCopy.Checked = True
                                            Else
                                               chkCopy.Checked = False
                                            End If

                                  Case Is = 14 'CheckBox Print

                                            If z = x Then
                                               chkPrint.Checked = True
                                            Else
                                               chkPrint.Checked = False
                                            End If

                                  Case Is = 15 'CheckBox Other

                                            If z = x Then
                                               chkOther.Checked = True
                                            Else
                                               chkOther.Checked = False
                                            End If

                        End Select

                        x = 0

                    Next i


            End If

End With

End Sub

Private Sub chkOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOpen.Click
Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkOpen.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(8, i).Value = 1
                                .Item(16, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(8, i).Value = 0
                                .Item(16, i).Value = My.Resources.lock_red
                             Next i


                    End If

            End If
End With

End Sub

Private Sub chkView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkView.Click
Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkView.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(9, i).Value = 1
                                .Item(17, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(9, i).Value = 0
                                .Item(17, i).Value = My.Resources.lock_red
                             Next i

                    End If

            End If
End With

End Sub

Private Sub chkAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAdd.Click
  Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkAdd.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(10, i).Value = 1
                                .Item(18, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(10, i).Value = 0
                                .Item(18, i).Value = My.Resources.lock_red
                             Next i

                    End If

            End If
End With

End Sub

Private Sub chkEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkEdit.Click
Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkEdit.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(11, i).Value = 1
                                .Item(19, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(11, i).Value = 0
                                .Item(19, i).Value = My.Resources.lock_red
                             Next i

                    End If

            End If
End With
End Sub

Private Sub chkDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDel.Click
Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkDel.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(12, i).Value = 1
                                .Item(20, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(12, i).Value = 0
                                .Item(20, i).Value = My.Resources.lock_red
                             Next i

                    End If

            End If
End With

End Sub

Private Sub chkCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCopy.Click
Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkCopy.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(13, i).Value = 1
                                .Item(21, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(13, i).Value = 0
                                .Item(21, i).Value = My.Resources.lock_red
                             Next i

                    End If

            End If
End With

End Sub

Private Sub chkPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPrint.Click
Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkPrint.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(14, i).Value = 1
                                .Item(22, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(14, i).Value = 0
                                .Item(22, i).Value = My.Resources.lock_red
                             Next i

                    End If

            End If
End With

End Sub

Private Sub chkOther_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOther.Click
Dim i As Integer

With dgvPs
            If .Rows.Count > 0 Then

                    If chkOther.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(15, i).Value = 1
                                .Item(23, i).Value = My.Resources.unlock_green
                             Next i

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(15, i).Value = 0
                                .Item(23, i).Value = My.Resources.lock_red
                             Next i

                    End If

            End If
End With

End Sub

End Class