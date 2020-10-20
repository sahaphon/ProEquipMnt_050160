Imports ADODB

Public Class frmUserPermit

Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean

    Private Function FormCount(ByVal frmName As String) As Long

        Dim frm As Form

        For Each frm In My.Application.OpenForms
            If frm Is My.Forms.frmAeUser Then
                FormCount = FormCount + 1
            End If
        Next

    End Function

    Private Sub frmUserPermit_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Dim strSearch As String

        If FormCount("frmAeUser") > 0 Then

            With frmAeUser

                strSearch = .lblComplete.Text

                If strSearch <> "" Then
                    SearchData(4, strSearch)
                End If

                .Close()

                Me.Height = Int(lblHeight.Text)
                Me.Width = Int(lblWidth.Text)

                Me.Top = Int(lblTop.Text)
                Me.Left = Int(lblLeft.Text)

            End With

        End If

    End Sub

    Private Sub frmUserPermit_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

        lblHeight.Text = Me.Height.ToString.Trim
        lblWidth.Text = Me.Width.ToString.Trim

        lblTop.Text = Me.Top.ToString.Trim
        lblLeft.Text = Me.Left.ToString.Trim

    End Sub

    Private Sub frmUserPermit_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub

    Private Sub frmUserPermit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        StdDateTimeThai()
        tlsBarFmr.Cursor = Cursors.Hand

        PreGroupType()
        InputDeptData()
        tabCmd.Focus()

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

        With cmbFilter

            For i = 0 To 1
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
        Dim strDateLog As String = ""

        Dim intPageCount As Integer
        Dim intPageSize As Integer
        Dim intCounter As Integer

        Dim strSearch As String = txtFilter.Text.ToString.Trim
        Dim strFieldFilter As String = ""

        With Conn
            If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
        End With

        If blnHaveFilter Then

            Select Case cmbFilter.SelectedIndex()

                Case Is = 0 'User LogIn
                    strFieldFilter = "user_id like '%" & ReplaceQuote(strSearch) & "%'"

                Case Is = 1 'ชื่อ-นามสกุล
                    strFieldFilter = "sname like '%" & ReplaceQuote(strSearch) & "%'"

            End Select

            strSqlCmdSelc = "SELECT * FROM v_head_user (NOLOCK)" _
                                         & " WHERE " & strFieldFilter _
                                         & " ORDER BY user_id"

        Else

            strSqlCmdSelc = "SELECT * FROM v_head_user (NOLOCK)" _
                                         & " ORDER BY user_id"

        End If

        intPageSize = 20

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
                    intPageSize = 20
                End If

                .PageSize = intPageSize
                intPageCount = .PageCount

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

                    strDateAdd = Mid(.Fields("cdate").Value.ToString, 1, 10)
                    strDateEdit = Mid(.Fields("edate").Value.ToString, 1, 10)
                    strDateLog = Mid(.Fields("log_date").Value.ToString, 1, 10)


                    dgvShoe.Rows.Add(
                                                                          IIf(.Fields("isexist").Value, My.Resources.stock_connect, My.Resources.stock_disconnect),
                                                                          IIf(.Fields("act_usr").Value.ToString.Trim = "A", My.Resources.admin, My.Resources.users),
                                                                          IIf(.Fields("isexist").Value, 1, 0),
                                                                          .Fields("usr_level").Value.ToString.Trim,
                                                                          .Fields("user_id").Value.ToString.Trim,
                                                                          .Fields("pass").Value.ToString.Trim,
                                                                          .Fields("sname").Value.ToString.Trim,
                                                                          .Fields("post").Value.ToString.Trim,
                                                                          .Fields("dept").Value.ToString.Trim,
                                                                           IIf(.Fields("sta_usr").Value, My.Resources.sign_deny, My.Resources.accept),
                                                                           .Fields("prmiss").Value.ToString.Trim,
                                                                           IIf(.Fields("sta_usr").Value, 1, 0),
                                                                            strDateLog,
                                                                           .Fields("log_time").Value.ToString.Trim,
                                                                           .Fields("com_ip").Value.ToString.Trim,
                                                                           strDateAdd,
                                                                           strDateEdit
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

    Private Sub dgvShoe_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShoe.CellDoubleClick

        Dim Conn As New ADODB.Connection
        Dim Rsd As New ADODB.Recordset

        Dim strSqlCmd As String
        Dim bytValue As Byte

        With Conn

            If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
        End With


        With dgvShoe

            If .Rows.Count > 0 Then

                Select Case .CurrentCell.ColumnIndex

                    Case Is = 0 'Column Online 

                        Select Case .Rows(.CurrentRow.Index).Cells(2).Value

                            Case Is = 1 'ยกเลิกการคำนวณ

                                .Rows(.CurrentRow.Index).Cells(2).Value = 0
                                .Rows(.CurrentRow.Index).Cells(0).Value = My.Resources.stock_disconnect
                                bytValue = 0

                            Case Is = 0 'คำนวณใหม่

                                .Rows(.CurrentRow.Index).Cells(2).Value = 1
                                .Rows(.CurrentRow.Index).Cells(0).Value = My.Resources.stock_connect
                                bytValue = 1

                        End Select

                        strSqlCmd = "UPDATE usermst SET isexist =" & bytValue.ToString _
                                                    & " WHERE user_id ='" & .Rows(.CurrentRow.Index).Cells(4).Value & "'"

                        Conn.Execute(strSqlCmd)

                    Case Is = 9 'สถานะจำกัดการเข้าสิทธิ

                        Select Case .Rows(.CurrentRow.Index).Cells(11).Value

                            Case Is = 1 'ปฏิเสธการเข้าใช้

                                .Rows(.CurrentRow.Index).Cells(9).Value = My.Resources.accept
                                .Rows(.CurrentRow.Index).Cells(10).Value = "อนุญาติ"
                                .Rows(.CurrentRow.Index).Cells(11).Value = 0
                                bytValue = 0

                            Case Is = 0 'ยอมรับการเข้าใช้

                                .Rows(.CurrentRow.Index).Cells(9).Value = My.Resources.sign_deny
                                .Rows(.CurrentRow.Index).Cells(10).Value = "ปฏิเสธ"
                                .Rows(.CurrentRow.Index).Cells(11).Value = 1
                                bytValue = 1

                        End Select

                        strSqlCmd = "UPDATE usermst SET sta_usr =" & bytValue.ToString _
                                                    & " WHERE user_id ='" & .Rows(.CurrentRow.Index).Cells(4).Value & "'"

                        Conn.Execute(strSqlCmd)

                    Case Else

                        ViewShoeData()

                End Select
            End If


        End With

        Conn.Close()
        Conn = Nothing

    End Sub

    Private Sub dgvShoe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvShoe.KeyDown

If e.KeyCode = Keys.Enter Then
    e.Handled = True
End If

End Sub

Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

        With tabCmd

            Select Case .SelectedIndex

                Case Is = 0 'เพิ่มข้อมูล

                    lblCmd.Text = "0"
                    With frmAeUser
                        .Show()
                        .Text = "เพิ่มข้อมูลผู้ใช้งาน"
                    End With

                    Me.Hide()
                    frmMainPro.Hide()

                Case Is = 1 'แก้ไขข้อมูล

                    If dgvShoe.Rows.Count > 0 Then

                        lblCmd.Text = "1"
                        With frmAeUser
                            .Show()
                            .Text = "แก้ไขข้อมูลผู้ใช้งาน"
                        End With

                        Me.Hide()
                        frmMainPro.Hide()

                    End If

                Case Is = 2 'มุมมองข้อมูล                            
                    ViewShoeData()

                Case Is = 3 'กรองข้อมูล

                    If dgvShoe.Rows.Count > 0 Then

                        With gpbFilter

                            .Top = 230
                            .Left = 210
                            .Width = 311
                            .Height = 125

                            .Visible = True

                            cmbFilter.SelectedItem = cmbFilter.Items(0)
                            txtFilter.Text =
                            dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(4).Value.ToString.Trim()

                            StateLockFind(False)
                            txtFilter.Focus()

                        End With
                    End If

                Case Is = 4 'ค้นหาข้อมูล

                    If dgvShoe.Rows.Count > 0 Then
                        With gpbSearch

                            .Top = 230
                            .Left = 210
                            .Width = 311
                            .Height = 125

                            .BringToFront()
                            .Visible = True

                            cmbType.SelectedItem = cmbType.Items(0)
                            txtSeek.Text =
                                             dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(4).Value.ToString.Trim()

                            StateLockFind(False)
                            txtSeek.Focus()

                        End With
                    End If

                Case Is = 5 'คัดลอกข้อมูล

                    If dgvShoe.Rows.Count > 0 Then
                        With gpbCopy

                            .Top = 230
                            .Left = 210
                            .Width = 311
                            .Height = 125

                            .BringToFront()
                            .Visible = True

                            lblSource.Text =
                                             dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(4).Value.ToString.Trim()

                            StateLockFind(False)
                            txtNew.Focus()

                        End With
                    End If

                Case Is = 6 'ลบข้อมูล                            
                    DeleteData()
                Case Is = 7 'ออก
                    Me.Close()
            End Select
        End With
    End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

  StateLockFind(True)
  gpbSearch.Visible = False

End Sub

Private Sub StateLockFind(ByVal Sta As Boolean)

    tabCmd.Enabled = Sta
    dgvShoe.Enabled = Sta
    tlsBarFmr.Enabled = Sta
    chkOpen.Enabled = Sta

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

Sub FindDocID()
Dim strSearch As String = txtSeek.Text.ToUpper.Trim

If strSearch <> "" Then

        Select Case cmbType.SelectedIndex()

                  Case Is = 0 'User LogIn
                          SearchData(4, strSearch)
                  Case Is = 1 'ชื่อ-นามสกุล
                          SearchData(6, strSearch)
                  Case Is = 2
                  Case Is = 3

        End Select

Else
    MsgBox("ไม่มีข้อมูลที่ต้องการค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
    txtSeek.Focus()
End If


End Sub


Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchTxt As String)

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset
Dim ClnRsd As New ADODB.Recordset

Dim intPageCount As Integer
Dim intPageSize As Integer


Dim strSqlCmdSelc As String = ""


Dim strSqlFind As String = ""
Dim i As Integer

       With Conn
              If .State Then .Close()
                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()
        End With

        strSqlCmdSelc = "SELECT * FROM v_head_user (NOLOCK)" _
                                    & " ORDER BY user_id"

        intPageSize = 20

        Rsd = New ADODB.Recordset
        With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                          If .RecordCount <> 0 Then

                                        ClnRsd = .Clone

                                         Select Case bytColNumber
                                                   Case Is = 4 'User LogIn
                                                          strSqlFind = "user_id "
                                                   Case Is = 6 'ชื่อ-นามสกุล
                                                           strSqlFind = "sname "

                                         End Select

                                         strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"
                                         '.Filter = ""
                                        ClnRsd.MoveFirst()
                                        ClnRsd.Find(strSqlFind)

                                        If ClnRsd.EOF Then
                                                MsgBox("ไม่มีข้อมูล : " & strSearchTxt & " ในระบบ" & vbNewLine _
                                                           & "โปรดระบุการค้นหาข้อมูลใหม่!", vbExclamation, "Not Found Data")

                                        Else


                                        If intPageSize > .RecordCount Then
                                                intPageSize = .RecordCount
                                        End If

                                        If intPageSize = 0 Then
                                                intPageSize = 20
                                        End If

                                        .PageSize = intPageSize
                                        intPageCount = .PageCount

                                        '---------------------------------------ค้นหาข้อมูล-------------------------------------------------------------
                                        .MoveFirst()
                                        .Find(strSqlFind)
                                         lblPage.Text = Str(.AbsolutePage)
                                        '-------------------------------------------------------------------------------------------------------------------

                                        InputDeptData()

                                                For i = 0 To dgvShoe.Rows.Count - 1
                                                        If InStr(UCase(dgvShoe.Rows(i).Cells(bytColNumber).Value), strSearchTxt.Trim.ToUpper) <> 0 Then
                                                                dgvShoe.CurrentCell = dgvShoe.Item(bytColNumber, i)
                                                                dgvShoe.Focus()
                                                                Exit For
                                                        End If
                                                Next i


                                        End If


                            ClnRsd.ActiveConnection = Nothing
                            ClnRsd.Close()
                            ClnRsd = Nothing

                     End If

            .ActiveConnection = Nothing
            .Close()

      End With

      Conn.Close()
      Conn = Nothing

StateLockFind(True)
gpbSearch.Visible = False

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

Private Sub cmbType_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbType.KeyPress

If e.KeyChar = Chr(13) Then
    txtSeek.Focus()
End If

End Sub

Private Sub dgvShoe_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvShoe.MouseMove

Dim objMousePosition As Point = dgvShoe.PointToClient(Control.MousePosition)
Dim objHitTestInfo As DataGridView.HitTestInfo
      objHitTestInfo = dgvShoe.HitTest(objMousePosition.X, objMousePosition.Y)

With dgvShoe

        Select Case objHitTestInfo.ColumnIndex
                  Case 0, 9
                         .Cursor = Cursors.Hand
                  Case Else
                        .Cursor = Cursors.Default
        End Select

End With
End Sub

Private Sub chkOpen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOpen.Click

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String
Dim strSqlCmd As String

Dim i As Integer
Dim bytValue As Byte

With Conn

        If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
         End With


With dgvShoe

            If .Rows.Count > 0 Then

                    strSqlCmdSelc = "SELECT user_id FROM v_head_user (NOLOCK)" _
                                          & " ORDER BY user_id"

                    Rsd = New ADODB.Recordset
                    Rsd.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                    Rsd.LockType = ADODB.LockTypeEnum.adLockOptimistic
                    Rsd.Open(strSqlCmdSelc, Conn, , , )

                    If chkOpen.CheckState = CheckState.Checked Then

                            For i = 0 To .Rows.Count - 1
                                .Item(2, i).Value = 1
                                .Item(0, i).Value = My.Resources.stock_connect
                             Next i

                            bytValue = 1

                    Else

                            For i = 0 To .Rows.Count - 1
                                .Item(2, i).Value = 0
                                .Item(0, i).Value = My.Resources.stock_disconnect
                             Next i

                            bytValue = 0


                    End If

                    Do While Not Rsd.EOF

                         strSqlCmd = "UPDATE usermst SET isexist =" & bytValue.ToString _
                                   & " WHERE user_id ='" & Rsd.Fields("user_id").Value.ToString.Trim & "'"

                         Conn.Execute(strSqlCmd)

                         Rsd.MoveNext()
                    Loop

                    Rsd.ActiveConnection = Nothing
                    Rsd.Close()


            End If

End With

Conn.Close()
Conn = Nothing

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

                strDept = .Rows(.CurrentRow.Index).Cells(4).Value.ToString.Trim
                strDeptName = .Rows(.CurrentRow.Index).Cells(6).Value.ToString.Trim

                btyConsider = MsgBox("User ID : " & strDept & vbNewLine _
                                                & "ชื่อ-นามสกุล : " & strDeptName & vbNewLine _
                                                & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then

                        Conn.BeginTrans()

                                '------------------------------------ลบตาราง usermst--------------------------------------------
                                strSqlCmd = "DELETE FROM usermst" _
                                               & " WHERE user_id ='" & strDept & "'"
                                Conn.Execute(strSqlCmd)

                                '------------------------------------ลบตาราง usertrn--------------------------------------------
                                strSqlCmd = "DELETE FROM usertrn" _
                                               & " WHERE user_id ='" & strDept & "'"
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

Private Sub ViewShoeData()

If dgvShoe.Rows.Count > 0 Then


        lblCmd.Text = "2"
        With frmAeUser
                .Show()
                .Text = "มุมมองข้อมูลผู้ใช้งาน"
        End With

        Me.Hide()
        frmMainpro.Hide()

End If

End Sub

Private Sub dgvShoe_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvShoe.KeyPress

If e.KeyChar = Chr(13) Then
    ViewShoeData()
End If

End Sub

Sub FilterData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String = ""
Dim strFieldFilter As String = ""

Dim blnHaveData As Boolean

Dim strSearch As String = ReplaceQuote(txtFilter.Text.ToUpper.Trim)

If strSearch <> "" Then


      With Conn
              If .State Then .Close()
                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()
      End With

                    Select Case cmbFilter.SelectedIndex()

                           Case Is = 0 'User LogIn
                                     strFieldFilter = "user_id like '%" & ReplaceQuote(strSearch) & "%'"
                           Case Is = 1 'ชื่อ-นามสกุล
                                     strFieldFilter = "sname like '%" & ReplaceQuote(strSearch) & "%'"
                    End Select

                    strSqlCmdSelc = "SELECT * FROM v_head_user (NOLOCK)" _
                                         & " WHERE " & strFieldFilter _
                                         & " ORDER BY user_id"

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

Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
  FilterData()
End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click

If blnHaveFilter Then
  blnHaveFilter = False
  InputDeptData()
End If

  StateLockFind(True)
  gpbFilter.Visible = False
End Sub

Private Sub btnCopyQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyQuit.Click

  StateLockFind(True)
  gpbCopy.Visible = False

End Sub

Private Sub txtNew_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNew.GotFocus
    txtNew.SelectAll()
End Sub

Private Sub txtNew_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNew.KeyPress

Dim strNewUser As String = txtNew.Text.ToUpper.ToString.Trim

 If e.KeyChar = Chr(13) Then

        If strNewUser <> "" Then
             CheckUsrDuplicate()
        Else
              MsgBox("โปรดระบุ New User LogIn ก่อนคัดลอกข้อมูล!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input New User!")
              txtNew.Focus()
        End If

  End If

End Sub

Private Sub CheckUsrDuplicate()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String = ""
Dim blnDup As Boolean

Dim strNewUser As String = ReplaceQuote(txtNew.Text.ToUpper.ToString.Trim)

               With Conn
                        If .State Then .Close()
                            .ConnectionString = strConnAdodb
                            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                            .ConnectionTimeout = 90
                            .Open()
                End With


                strSqlCmdSelc = "SELECT * FROM usermst (NOLOCK)" _
                                      & " WHERE user_id ='" & strNewUser & "'"


                Rsd = New ADODB.Recordset
                With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                          If .RecordCount <> 0 Then
                             blnDup = True
                          Else
                             blnDup = False
                          End If

                         .ActiveConnection = Nothing
                         .Close()

                End With
                Rsd = Nothing

                Conn.Close()
                Conn = Nothing

                 If blnDup Then

                     MsgBox("New User LogIn : " & txtNew.Text.ToString.Trim & vbNewLine _
                                 & "มีอยู่แล้ว โปรดระบุข้อมูลใหม่!!" _
                              , MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Duplicate User Data!")
                     txtNew.Focus()

                 Else

                    UserCopy(lblSource.Text, strNewUser)
                    InputDeptData()
                    StateLockFind(True)
                    gpbCopy.Visible = False
                    SearchData(4, strNewUser)

                 End If

End Sub

Private Sub UserCopy(ByVal strOldUser As String, ByVal strNewUser As String)
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

            '------------------------------------------------บันทึกข้อมูลในตาราง usermst-----------------------------------------------
             strSqlCmd = "INSERT INTO usermst" _
                            & " SELECT cdate,edate,act_usr,sname,post,dept,user_id='" & strNewUser & "'" _
                            & ",pass,log_date,log_time,com_ip,isexist=0,sta_usr,pic_sign" _
                            & "  FROM usermst" _
                            & " WHERE user_id ='" & strOldUser & "'"

             Conn.Execute(strSqlCmd)

            '------------------------------------------------บันทึกข้อมูลในตาราง usertrn-----------------------------------------------

             strSqlCmd = "INSERT INTO usertrn" _
                            & " SELECT user_id='" & strNewUser & "'" _
                            & ",file_icon,open_cnt,last_date,last_time" _
                            & ",act_open,act_view,act_add,act_edit,act_delete,act_copy,act_print,act_other" _
                            & ",ps00,ps01,ps02,ps03,ps04,ps05,ps06,ps07,ps08,ps09,ps10,dep_app" _
                            & " FROM usertrn" _
                            & " WHERE user_id ='" & strOldUser & "'"

             Conn.Execute(strSqlCmd)

             Conn.CommitTrans()

             Conn.Close()
             Conn = Nothing


End Sub

Private Sub btnCopyOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyOK.Click
 Dim strNewUser As String = txtNew.Text.ToUpper.ToString.Trim

        If strNewUser <> "" Then
             CheckUsrDuplicate()
        Else
              MsgBox("โปรดระบุ New User LogIn ก่อนคัดลอกข้อมูล!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input New User!")
              txtNew.Focus()
        End If

End Sub

End Class