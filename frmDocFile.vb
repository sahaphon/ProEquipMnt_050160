Imports ADODB

Public Class frmDocFile

Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean

Private Sub frmDocFile_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
Dim strSearch As String

    If FormCount("frmAeDocFile") > 0 Then

        With frmAeDocFile

               strSearch = .lblComplete.Text

                                If strSearch <> "" Then
                                        SearchData(0, strSearch)
                                End If

              .Close()

              Me.Height = Int(lblHeight.Text)
              Me.Width = Int(lblWidth.Text)

              Me.Top = Int(lblTop.Text)
              Me.Left = Int(lblLeft.Text)

        End With

    End If

End Sub

Private Sub frmDocFile_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

    lblHeight.Text = Me.Height.ToString.Trim
    lblWidth.Text = Me.Width.ToString.Trim

    lblTop.Text = Me.Top.ToString.Trim
    lblLeft.Text = Me.Left.ToString.Trim

End Sub

Private Sub frmDocFile_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    Me.Dispose()
End Sub

Private Sub frmDocFile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Me.WindowState = FormWindowState.Maximized
  StdDateTimeThai()
  tlsBarFmr.Cursor = Cursors.Hand

  PreGroupType()
  InputDeptData()
  tabCmd.Focus()

End Sub

Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click
With tabCmd

          Select Case .SelectedIndex
                    Case Is = 0 'เพิ่มข้อมูล

                            lblCmd.Text = "0"
                            With frmAeDocFile
                                    .Show()
                                    .Text = "เพิ่มข้อมูลแฟ้มระบบ"
                            End With

                            Me.Hide()
                            frmMainpro.Hide()

                    Case Is = 1 'แก้ไขข้อมูล

                            If dgvShoe.Rows.Count > 0 Then

                                    lblCmd.Text = "1"
                                    With frmAeDocFile
                                            .Show()
                                            .Text = "แก้ไขข้อมูลแฟ้มระบบ"
                                    End With

                                    Me.Hide()
                                    frmMainpro.Hide()

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
                                            .Width = 311
                                            .Height = 125

                                            .Visible = True

                                              cmbType.SelectedItem = cmbType.Items(0)
                                             txtSeek.Text = _
                                             dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim()

                                              StateLockFind(False)
                                              txtSeek.Focus()

                                 End With
                            End If

                    Case Is = 5 'ลบข้อมูล
                            DeleteData()
                    Case Is = 6 'ออก
                            Me.Close()
          End Select
End With

End Sub

Private Sub InputDeptData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset


Dim strSqlCmdSelc As String = ""
Dim strDateAdd As String = ""
Dim strDateEdit As String = ""

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
                             Case Is = 0 'รหัสแฟ้ม
                                     strFieldFilter = "file_icon like '%" & ReplaceQuote(strSearch) & "%'"
                             Case Is = 1 'ชื่อแฟ้มระบบงาน
                                     strFieldFilter = "file_name like '%" & ReplaceQuote(strSearch) & "%'"
                    End Select

                    strSqlCmdSelc = "SELECT * FROM filemst (NOLOCK)" _
                                          & " WHERE " & strFieldFilter _
                                          & " ORDER BY file_icon"


              Else

                    strSqlCmdSelc = "SELECT * FROM filemst (NOLOCK)" _
                                         & " ORDER BY file_icon"

              End If

              intPageSize = 23

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
                                        intPageSize = 23
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

                                            strDateAdd = Mid(.Fields("pre_date").Value.ToString, 1, 10)
                                            strDateEdit = Mid(.Fields("last_date").Value.ToString, 1, 10)


                                            dgvShoe.Rows.Add( _
                                                                          .Fields("file_icon").Value.ToString.Trim, _
                                                                          .Fields("file_name").Value.ToString.Trim, _
                                                                           strDateAdd, .Fields("pre_by").Value.ToString.Trim, _
                                                                           strDateEdit, .Fields("last_by").Value.ToString.Trim _
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

Private Sub dgvShoe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvShoe.KeyDown

If e.KeyCode = Keys.Enter Then
    e.Handled = True
End If

End Sub

Private Function FormCount(ByVal frmName As String) As Long
Dim frm As Form

    For Each frm In My.Application.OpenForms
                If frm Is My.Forms.frmAeDocFile Then
                        FormCount = FormCount + 1
                End If
    Next
End Function

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


        strSqlCmdSelc = "SELECT * FROM filemst (NOLOCK)" _
                                    & " ORDER BY file_icon"

        intPageSize = 23

        Rsd = New ADODB.Recordset
        With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                          If .RecordCount <> 0 Then

                                        ClnRsd = .Clone

                                         Select Case bytColNumber
                                                   Case Is = 0 'รหัสแฟ้ม
                                                          strSqlFind = "file_icon "
                                                   Case Is = 1 'ชื่อแฟ้มระบบงาน
                                                           strSqlFind = "file_name "
                                                   Case Is = 2
                                                   Case Is = 3

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
                                                intPageSize = 23
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
                strDeptName = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim

                btyConsider = MsgBox("รหัสแฟ้ม : " & strDept & vbNewLine _
                                                & "ชื่อแฟ้มระบบงาน : " & strDeptName & vbNewLine _
                                                & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then

                        Conn.BeginTrans()

                                '------------------------------------ลบตาราง filemst--------------------------------------------
                                strSqlCmd = "DELETE FROM filemst" _
                                               & " WHERE file_icon ='" & strDept & "'"
                                Conn.Execute(strSqlCmd)
                                '------------------------------------ลบตาราง usertrn--------------------------------------------
                                strSqlCmd = "DELETE FROM usertrn" _
                                               & " WHERE file_icon ='" & strDept & "'"
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
        With frmAeDocFile
                .Show()
                .Text = "มุมมองข้อมูลแฟ้มระบบงาน"
        End With

        Me.Hide()
        frmMainpro.Hide()

End If

End Sub

Private Sub dgvShoe_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShoe.CellContentClick

End Sub

Private Sub dgvShoe_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvShoe.KeyPress

If e.KeyChar = Chr(13) Then
    ViewShoeData()
End If

End Sub

Private Sub dgvShoe_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShoe.CellDoubleClick
    ViewShoeData()
End Sub

Private Sub PreGroupType()

Dim strGpTopic(1) As String
Dim i As Byte

      strGpTopic(0) = "รหัสแฟ้ม"
      strGpTopic(1) = "ชื่อแฟ้ม"


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

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

  StateLockFind(True)
  gpbSearch.Visible = False

End Sub

Private Sub StateLockFind(ByVal Sta As Boolean)

    tabCmd.Enabled = Sta
    dgvShoe.Enabled = Sta
    tlsBarFmr.Enabled = Sta

End Sub

Private Sub txtSeek_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSeek.GotFocus
    txtSeek.SelectAll()
End Sub

Private Sub txtSeek_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSeek.KeyPress

If e.KeyChar = Chr(13) Then
    FindDocID()
End If

End Sub

Sub FindDocID()
Dim strSearch As String = txtSeek.Text.ToUpper.Trim

If strSearch <> "" Then

        Select Case cmbType.SelectedIndex()

                  Case Is = 0 'รหัสหน่วยงาน
                          SearchData(0, strSearch)
                  Case Is = 1 'ชื่อหน่วยงาน
                          SearchData(1, strSearch)
                  Case Is = 2 'รหัสหน่วยงาน                          
                  Case Is = 3 'ชื่อหน่วยงาน

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

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  FindDocID()
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
                             Case Is = 0 'รหัสแฟ้ม
                                     strFieldFilter = "file_icon like '%" & ReplaceQuote(strSearch) & "%'"
                             Case Is = 1 'ชื่อแฟ้มระบบงาน
                                     strFieldFilter = "file_name like '%" & ReplaceQuote(strSearch) & "%'"
                    End Select

                    strSqlCmdSelc = "SELECT * FROM filemst (NOLOCK)" _
                                          & " WHERE " & strFieldFilter _
                                          & " ORDER BY file_icon"

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

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click
  If blnHaveFilter Then
        blnHaveFilter = False
        InputDeptData()
  End If

    StateLockFind(True)
    gpbFilter.Visible = False
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

End Class