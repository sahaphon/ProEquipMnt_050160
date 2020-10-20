Imports ADODB

Public Class frmAeUser

Protected Overrides ReadOnly Property CreateParams() As CreateParams '��ͧ�ѹ��ûԴ������� Close Button
Get
Dim cp As CreateParams = MyBase.CreateParams
Const CS_DBLCLKS As Int32 = &H8
Const CS_NOCLOSE As Int32 = &H200
cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
Return cp
End Get
End Property

Private Sub frmAeUser_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    Me.Dispose()
End Sub


Private Sub frmAeUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

Me.WindowState = FormWindowState.Maximized
StdDateTimeThai()

PreGroupType()
PreGroupLevel()

Select Case frmUserPermit.lblCmd.Text.ToString

          Case Is = "0" '����������
                  ShowFileDefault()
                  cmbLevel.Focus()
          Case Is = "1" '��䢢�����
                  LockEditData()
                  cmbLevel.Focus()
          Case Is = "2" '����ͧ������
                  LockEditData()
                  btnSaveData.Enabled = False
                  cmbLevel.Focus()

 End Select

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click

  Me.Hide()
  frmMainPro.Show()
  frmUserPermit.Show()

End Sub

Private Sub ShowFileDefault()

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

         strSqlCmdSelc = "SELECT * FROM filemst (NOLOCK)" _
                              & " ORDER BY file_icon"

         Rsd = New ADODB.Recordset
         With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                         dgvPs.Rows.Clear()

                          If .RecordCount <> 0 Then


                                Do While Not .EOF()

                                        dgvPs.Rows.Add( _
                                                                   My.Resources.open_folder00, _
                                                                  .Fields("file_icon").Value.ToString.Trim, _
                                                                  .Fields("file_name").Value.ToString.Trim, _
                                                                  .Fields("file_icon").Value.ToString.Trim, _
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

Private Sub dgvPs_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPs.CellDoubleClick

With dgvPs

        If .Rows.Count > 0 Then

                Select Case .CurrentCell.ColumnIndex

                          Case Is = 16 'ʶҹ� Open

                                    If .Rows(.CurrentRow.Index).Cells(8).Value = 0 Then
                                         .Item(8, .CurrentRow.Index).Value = 1
                                         .Item(16, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(8, .CurrentRow.Index).Value = 0
                                         .Item(16, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                          Case Is = 17 'ʶҹ� View

                                    If .Rows(.CurrentRow.Index).Cells(9).Value = 0 Then
                                         .Item(9, .CurrentRow.Index).Value = 1
                                         .Item(17, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(9, .CurrentRow.Index).Value = 0
                                         .Item(17, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                         Case Is = 18 'ʶҹ� Add

                                    If .Rows(.CurrentRow.Index).Cells(10).Value = 0 Then
                                         .Item(10, .CurrentRow.Index).Value = 1
                                         .Item(18, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(10, .CurrentRow.Index).Value = 0
                                         .Item(18, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 19 'ʶҹ� Edit

                                    If .Rows(.CurrentRow.Index).Cells(11).Value = 0 Then
                                         .Item(11, .CurrentRow.Index).Value = 1
                                         .Item(19, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(11, .CurrentRow.Index).Value = 0
                                         .Item(19, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 20 'ʶҹ� Del

                                    If .Rows(.CurrentRow.Index).Cells(12).Value = 0 Then
                                         .Item(12, .CurrentRow.Index).Value = 1
                                         .Item(20, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(12, .CurrentRow.Index).Value = 0
                                         .Item(20, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 21 'ʶҹ� Copy

                                    If .Rows(.CurrentRow.Index).Cells(13).Value = 0 Then
                                         .Item(13, .CurrentRow.Index).Value = 1
                                         .Item(21, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(13, .CurrentRow.Index).Value = 0
                                         .Item(21, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 22 'ʶҹ� Print

                                    If .Rows(.CurrentRow.Index).Cells(14).Value = 0 Then
                                         .Item(14, .CurrentRow.Index).Value = 1
                                         .Item(22, .CurrentRow.Index).Value = My.Resources.unlock_green
                                    Else
                                         .Item(14, .CurrentRow.Index).Value = 0
                                         .Item(22, .CurrentRow.Index).Value = My.Resources.lock_red
                                    End If

                        Case Is = 23 'ʶҹ� Other

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

Private Sub dgvPs_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvPs.KeyDown

If e.KeyCode = Keys.Enter Then
    e.Handled = True
End If

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

Private Sub CheckStatusCheckBox()
Dim i, x, z As Integer

With dgvPs

            If .Rows.Count > 0 Then

                   '------------------------------------------------ �ٻ Perimission --------------------------------------------

                    For i = 8 To 15 'column

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

Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

  If dgvPs.Rows.Count > 0 Then

      With gpbSearch

                 .Top = 300
                 .Left = 350
                 .Width = 311
                 .Height = 125

                 .Visible = True

                 cmbType.SelectedItem = cmbType.Items(0)
                 txtSeek.Text = _
                 dgvPs.Rows(dgvPs.CurrentRow.Index).Cells(1).Value.ToString.Trim()

                 StateLockFind(False)
                 txtSeek.Focus()

       End With

    End If

End Sub

Private Sub StateLockFind(ByVal Sta As Boolean)

    gpbWc.Enabled = Sta
    gpbSub.Enabled = Sta
    btnSaveData.Enabled = Sta
    btnSearch.Enabled = Sta

End Sub

Private Sub PreGroupType()

Dim strGpTopic(1) As String
Dim i As Byte

      strGpTopic(0) = "��������к�"
      strGpTopic(1) = "��������к�"


      With cmbType

              For i = 0 To 1
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(0)

      End With

End Sub

Private Sub PreGroupLevel()

Dim strGpTopic(1) As String
Dim i As Byte

      strGpTopic(0) = "A"
      strGpTopic(1) = "U"


      With cmbLevel

              For i = 0 To 1
                 .Items.Add(strGpTopic(i))
              Next i

              .SelectedItem = .Items(1)

      End With

End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
    StateLockFind(True)
    gpbSearch.Visible = False
End Sub

Private Sub FindData()

Dim i, x As Integer
Dim z As Boolean

Dim strSearchTxt As String


    Select Case cmbType.SelectedIndex()
              Case Is = 0 '�������
                       x = 1
              Case Is = 1 '��������к��ҹ
                       x = 2
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
        MsgBox("������ : " & strSearchTxt & " �����������к�!!" & vbNewLine _
                     & "�ô�кآ���������", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Not Found Data!!")
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


Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  FindData()
End Sub

Private Sub cmbLevel_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbLevel.KeyPress

If e.KeyChar = Chr(13) Then
    txtSurName.Focus()
End If

End Sub

Private Sub cmbLevel_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbLevel.LostFocus

With cmbLevel
         .Text = .Text.ToString.Trim.ToUpper
End With

Select Case cmbLevel.Text.ToString.Trim
          Case Is = "A"
                  lblDesc.Text = "Adminstrator"
          Case Is = "U"
                  lblDesc.Text = "User"
End Select
End Sub

Private Sub txtSurName_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurName.GotFocus
    txtSurName.SelectAll()
End Sub

Private Sub txtSurName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSurName.KeyDown
Dim intChkPoint As Integer

    With txtSurName

            Select Case e.KeyCode
                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���
                               If .SelectionStart = 0 Then
                               End If
                      Case Is = 38 '�����١�â��                                
                               cmbLevel.Focus()
                      Case Is = 39 '�����١�â��
                                If .SelectionLength = .Text.Trim.Length Then
                                    txtPost.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtPost.Focus()
                                    End If
                                End If
                      Case Is = 40 ''�����١��ŧ
                              txtDept.Focus()
                      Case Is = 113 '���� F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtSurName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSurName.KeyPress

If e.KeyChar = Chr(13) Then
    txtPost.Focus()
End If

End Sub

Private Sub txtSurName_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurName.LostFocus

With txtSurName
     .Text = .Text.ToString.Trim.ToUpper
End With

End Sub

Private Sub txtPost_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPost.GotFocus
    txtPost.SelectAll()
End Sub

Private Sub txtPost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPost.KeyDown
Dim intChkPoint As Integer

    With txtPost

            Select Case e.KeyCode
                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���
                               If .SelectionStart = 0 Then
                                    txtSurName.Focus()
                               End If
                      Case Is = 38 '�����١�â��                                                               
                      Case Is = 39 '�����١�â��
                                If .SelectionLength = .Text.Trim.Length Then
                                    txtDept.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtDept.Focus()
                                    End If
                                End If
                      Case Is = 40 ''�����١��ŧ                              
                      Case Is = 113 '���� F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub txtPost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPost.KeyPress

If e.KeyChar = Chr(13) Then
    txtDept.Focus()
End If

End Sub

Private Sub txtPost_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPost.LostFocus

With txtPost
     .Text = .Text.ToString.Trim.ToUpper
End With

End Sub

Private Sub txtDept_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDept.GotFocus
    txtDept.SelectAll()
End Sub

Private Sub txtDept_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDept.KeyDown
Dim intChkPoint As Integer

    With txtDept

            Select Case e.KeyCode
                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���
                               If .SelectionStart = 0 Then
                                    txtPost.Focus()
                               End If
                      Case Is = 38 '�����١�â��                                                               
                                    txtSurName.Focus()
                      Case Is = 39 '�����١�â��
                                If .SelectionLength = .Text.Trim.Length Then
                                    txtUser.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtUser.Focus()
                                    End If
                                End If
                      Case Is = 40 ''�����١��ŧ    
                              txtPass.Focus()
                      Case Is = 113 '���� F2
                                .SelectionStart = .Text.Trim.Length

            End Select
    End With
End Sub

Private Sub txtDept_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDept.KeyPress

    If e.KeyChar = Chr(13) Then
        txtUser.Focus()
    End If

End Sub

Private Sub txtDept_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDept.LostFocus
  With txtDept
       .Text = .Text.ToString.Trim.ToUpper
  End With
End Sub

Private Sub txtUser_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUser.GotFocus
 txtUser.SelectAll()
End Sub

Private Sub txtUser_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUser.KeyDown
Dim intChkPoint As Integer

    With txtUser

            Select Case e.KeyCode
                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���
                               If .SelectionStart = 0 Then
                                    txtDept.Focus()
                               End If
                      Case Is = 38 '�����١�â��                                                               
                                    txtPost.Focus()
                      Case Is = 39 '�����١�â��
                                If .SelectionLength = .Text.Trim.Length Then
                                    txtPass.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtPass.Focus()
                                    End If
                                End If
                      Case Is = 40 ''�����١��ŧ                              
                      Case Is = 113 '���� F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub txtUser_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtUser.KeyPress

If e.KeyChar = Chr(13) Then
        txtPass.Focus()
End If

End Sub

Private Sub txtUser_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUser.LostFocus

With txtUser
     .Text = .Text.ToString.Trim.ToUpper
End With

End Sub

Private Sub txtPass_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPass.GotFocus
    txtPass.SelectAll()
End Sub

Private Sub txtPass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPass.KeyDown
Dim intChkPoint As Integer

    With txtPass

            Select Case e.KeyCode
                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���

                               If .SelectionStart = 0 Then
                                    txtUser.Focus()
                               End If

                      Case Is = 38 '�����١�â��                                                               
                                    txtPost.Focus()
                      Case Is = 39 '�����١�â��

                                If .SelectionLength = .Text.Trim.Length Then
                                    txtPass.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtPass.Focus()
                                    End If
                                End If

                      Case Is = 40 ''�����١��ŧ                              
                      Case Is = 113 '���� F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtPass_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPass.KeyPress

If e.KeyChar = Chr(13) Then
        txtPost.Focus()
End If

End Sub

Private Sub txtPass_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPass.LostFocus

With txtPass
     .Text = .Text.ToString.Trim.ToUpper
End With

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBeforeSave()
End Sub

Private Sub CheckDataBeforeSave()

Dim bytConSave As Byte

    If cmbLevel.Text.Trim.ToString <> "" Then 'Level

            If txtSurName.Text.Trim.ToString <> "" Then '���͹��ʡ��

                    If txtUser.Text.Trim.ToString <> "" Then 'User Id

                            If txtPass.Text.Trim.ToString <> "" Then 'PassWord


                                        bytConSave = MsgBox("�س��ͧ��úѹ�֡���������������!" _
                                                                        , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                        If bytConSave = 6 Then
                                                    Select Case Me.Text

                                                           Case Is = "���������ż����ҹ"
                                                                      SaveNewRecord()
                                                           Case Is = "��䢢����ż����ҹ"
                                                                      SaveEditRecord()
                                                     End Select
                                        Else
                                            txtSurName.Focus()
                                        End If

                            Else

                                MsgBox("�ô�к� PassWord ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                txtPass.Focus()

                            End If

                    Else

                        MsgBox("�ô�к� User Id ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                        txtUser.Focus()

                    End If


            Else


                MsgBox("�ô�кت���-���ʡ�� ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                txtSurName.Focus()

            End If

    Else

        MsgBox("�ô�к��дѺ������� ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
        cmbLevel.Focus()

    End If

End Sub

Private Sub SaveNewRecord()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlCmdSelc As String
  Dim strSqlCmd As String

  Dim strUser As String = txtUser.Text.ToString.Trim
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

        strSqlCmdSelc = "SELECT user_id FROM usermst (NOLOCK)" _
                              & " WHERE user_id ='" & ReplaceQuote(strUser) & "'"

         Rsd = New ADODB.Recordset
         With Rsd

                 .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                 .LockType = ADODB.LockTypeEnum.adLockOptimistic
                 .Open(strSqlCmdSelc, Conn, , , )

                  If .RecordCount <> 0 Then

                          MsgBox("������ User Id : " & strUser & " ����������!" & vbNewLine _
                                      & "�ô�кآ���������", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Duplication Data!!")
                          txtUser.Focus()
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

                    '------------------------------------------------���� User ID -----------------------------------------------------------------                                
                    strSqlCmd = "INSERT INTO usermst " _
                                   & "(act_usr,sname,post,dept,user_id,pass,cdate,isexist,sta_usr)" _
                                   & " VALUES (" _
                                   & "'" & cmbLevel.Text.ToString.Trim & "'" _
                                   & ",'" & ReplaceQuote(txtSurName.Text.ToString.Trim) & "'" _
                                   & ",'" & ReplaceQuote(txtPost.Text.ToString.Trim) & "'" _
                                   & ",'" & ReplaceQuote(txtDept.Text.ToString.Trim) & "'" _
                                   & ",'" & ReplaceQuote(txtUser.Text.ToString.Trim) & "'" _
                                   & ",'" & ReplaceQuote(txtPass.Text.ToString.Trim) & "'" _
                                   & ",'" & strDate & "'" _
                                   & "," & "0" _
                                   & "," & "0" _
                                   & ")"
                    Conn.Execute(strSqlCmd)

                    '------------------------------------------------�����Է�� User �����к� -----------------------------------------------------------------                                
                    With dgvPs

                            z = .Rows.Count

                            If z > 0 Then

                                For i = 0 To .Rows.Count - 1

                                    strSqlCmd = "INSERT INTO usertrn " _
                                                    & "(user_id,file_icon,open_cnt,last_date,last_time," _
                                                    & "act_open,act_view,act_add,act_edit,act_delete,act_copy,act_print,act_other" _
                                                    & ")" _
                                                    & " VALUES (" _
                                                    & "'" & ReplaceQuote(txtUser.Text.ToString.Trim) & "'" _
                                                    & ",'" & .Rows(i).Cells(1).Value & "'" _
                                                    & "," & 0 _
                                                    & ",'" & "1899-01-01" & "'" _
                                                    & ",'" & "00:00:00" & "'" _
                                                    & "," & IIf(.Rows(i).Cells(8).Value.ToString.Trim <> "", .Rows(i).Cells(8).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(9).Value.ToString.Trim <> "", .Rows(i).Cells(9).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(10).Value.ToString.Trim <> "", .Rows(i).Cells(10).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(11).Value.ToString.Trim <> "", .Rows(i).Cells(11).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(12).Value.ToString.Trim <> "", .Rows(i).Cells(12).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(13).Value.ToString.Trim <> "", .Rows(i).Cells(13).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(14).Value.ToString.Trim <> "", .Rows(i).Cells(14).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(15).Value.ToString.Trim <> "", .Rows(i).Cells(15).Value.ToString.Trim, "0") _
                                                    & ")"
                                    Conn.Execute(strSqlCmd)

                                Next i

                            End If

                    End With


                    Conn.CommitTrans()

                    lblComplete.Text = strUser  '�觺͡��Һѹ�֡�����������

                    Me.Hide()
                    frmMainpro.Show()
                    frmUserPermit.Show()


          End If

          Conn.Close()
          Conn = Nothing

End Sub

Private Sub LockEditData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String
Dim strDateEdit As String
Dim strDocFile As String

        strDocFile = frmUserPermit.dgvShoe.Rows(frmUserPermit.dgvShoe.CurrentRow.Index).Cells(4).Value.ToString.Trim
        txtUser.Text = strDocFile

        txtUser.ReadOnly = True

        With Conn

                If .State Then .Close()
                   .ConnectionString = strConnAdodb
                   .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                   .ConnectionTimeout = 90
                   .Open()

         End With

         strSqlCmdSelc = "SELECT * FROM v_usr_permit (NOLOCK)" _
                              & " WHERE user_id ='" & strDocFile & "'" _
                              & " ORDER BY file_icon"

         Rsd = New ADODB.Recordset
         With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                         dgvPs.Rows.Clear()

                         If .RecordCount <> 0 Then

                                cmbLevel.Text = .Fields("act_usr").Value.ToString.Trim
                                lblDesc.Text = IIf(.Fields("act_usr").Value.ToString.Trim = "A", "Administrator", "User")
                                txtSurName.Text = .Fields("sname").Value.ToString.Trim
                                txtPost.Text = .Fields("post").Value.ToString.Trim
                                txtDept.Text = .Fields("dept").Value.ToString.Trim()
                                txtPass.Text = .Fields("pass").Value.ToString.Trim()

                                Do While Not .EOF()


                                        strDateEdit = Mid(.Fields("last_date").Value.ToString, 1, 10)

                                        dgvPs.Rows.Add( _
                                                                   My.Resources.open_folder00, _
                                                                  .Fields("file_icon").Value.ToString.Trim, _
                                                                  .Fields("file_name").Value.ToString.Trim, _
                                                                   "", _
                                                                   "", _
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

Private Sub SaveEditRecord()

Dim Conn As New ADODB.Connection

Dim strSqlCmd As String
Dim strUser As String = txtUser.Text.ToString.Trim

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


                    '------------------------------------------------���� User ID---------------------------------------------------------

                      strDate = datSave.Date.ToString("yyyy-MM-dd")
                      strDate = SaveChangeEngYear(strDate)

                      strSqlCmd = "UPDATE usermst SET act_usr ='" & cmbLevel.Text.ToString.Trim & "'" _
                                   & "," & "sname ='" & ReplaceQuote(txtSurName.Text.ToString.Trim) & "'" _
                                   & "," & "post ='" & ReplaceQuote(txtPost.Text.ToString.Trim) & "'" _
                                   & "," & "dept ='" & ReplaceQuote(txtDept.Text.ToString.Trim) & "'" _
                                   & "," & "pass ='" & ReplaceQuote(txtPass.Text.ToString.Trim) & "'" _
                                   & "," & "edate ='" & strDate & "'" _
                                   & " WHERE user_id ='" & txtUser.Text.ToString.Trim & "'"
                     Conn.Execute(strSqlCmd)

                    '------------------------------------------------ź������㹵��ҧ usertrn----------------------------------------------------
                     strSqlCmd = "Delete FROM usertrn" _
                                    & " WHERE user_id ='" & txtUser.Text.ToString.Trim & "'"
                     Conn.Execute(strSqlCmd)


                    '------------------------------------------------������������к� �����User-----------------------------------------------------------------                                
                    With dgvPs

                            z = .Rows.Count

                            If z > 0 Then

                                For i = 0 To .Rows.Count - 1

                                    If .Rows(i).Cells(6).Value.ToString.Trim <> "" Then
                                            strDateEntry = Mid(.Rows(i).Cells(6).Value.ToString, 7, 4) & "-" _
                                                               & Mid(.Rows(i).Cells(6).Value.ToString, 4, 2) & "-" _
                                                               & Mid(.Rows(i).Cells(6).Value.ToString, 1, 2)

                                            strDateEntry = SaveChangeEngYear(strDateEntry)
                                    Else
                                            strDateEntry = ""
                                    End If

                                    strSqlCmd = "INSERT INTO usertrn " _
                                                    & "(user_id,file_icon," _
                                                    & "act_open,act_view,act_add,act_edit,act_delete,act_copy,act_print,act_other," _
                                                    & "open_cnt,last_date,last_time" _
                                                    & ")" _
                                                    & " VALUES (" _
                                                    & "'" & ReplaceQuote(txtUser.Text.ToString.Trim) & "'" _
                                                    & ",'" & .Rows(i).Cells(1).Value & "'" _
                                                    & "," & IIf(.Rows(i).Cells(8).Value.ToString.Trim <> "", .Rows(i).Cells(8).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(9).Value.ToString.Trim <> "", .Rows(i).Cells(9).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(10).Value.ToString.Trim <> "", .Rows(i).Cells(10).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(11).Value.ToString.Trim <> "", .Rows(i).Cells(11).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(12).Value.ToString.Trim <> "", .Rows(i).Cells(12).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(13).Value.ToString.Trim <> "", .Rows(i).Cells(13).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(14).Value.ToString.Trim <> "", .Rows(i).Cells(14).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(15).Value.ToString.Trim <> "", .Rows(i).Cells(15).Value.ToString.Trim, "0") _
                                                    & "," & IIf(.Rows(i).Cells(5).Value.ToString.Trim <> "", .Rows(i).Cells(5).Value.ToString.Trim, "0") _
                                                    & ",'" & strDateEntry & "'" _
                                                    & ",'" & .Rows(i).Cells(7).Value.ToString & "'" _
                                                    & ")"

                                    Conn.Execute(strSqlCmd)

                                Next i

                            End If

                    End With

                    Conn.CommitTrans()

                    lblComplete.Text = strUser  '�觺͡��Һѹ�֡�����������
                    Me.Hide()
                    frmMainpro.Show()
                    frmUserPermit.Show()


          Conn.Close()
          Conn = Nothing

End Sub

Private Sub cmbGp_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

If e.KeyChar = Chr(13) Then
    btnSaveData.Focus()
End If

End Sub

End Class