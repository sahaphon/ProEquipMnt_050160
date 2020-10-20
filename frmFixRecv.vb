Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream
Imports System.Windows.Forms.DataGridView

Public Class frmFixRecv

Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean    '�óա�ͧ������
Dim IsShowSeek As Boolean

Dim dubNumberStart As Double   '�١��˹� = 1
Dim dubNumberEnd As Double     '�١��˹� = 2100

Dim strSqlFindData As String
Dim strDocCode As String = "F7"

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim dsTn As New DataSet

Private Sub frmFixRecv_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

     If FormCount("frmAeFixRecv") > 0 Then

        With frmAeFixRecv

               strSearch = .lblComplete.Text          '���ʫ���

                If strSearch <> "" Then
                   SearchData(0, strSearch)
                End If

                .Close()

        End With

     End If
     Timer1.Enabled = True          '������ê�����ŷء 1 �ҷ�

    Me.Height = Int(lblHeight.Text)
    Me.Width = Int(lblWidth.Text)

    Me.Top = Int(lblTop.Text)
    Me.Left = Int(lblLeft.Text)

End Sub

Private Function FormCount(ByVal frmName As String) As Long
  Dim frm As Form

    For Each frm In My.Application.OpenForms

         If frm Is My.Forms.frmAeFixRecv Then
            FormCount = FormCount + 1
         End If
    Next

End Function

Private Sub frmFixRecv_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

  lblHeight.Text = Me.Height.ToString
  lblWidth.Text = Me.Width.ToString

  lblTop.Text = Me.Top.ToString
  lblLeft.Text = Me.Left.ToString

End Sub

Private Sub frmFixRecv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 Me.Dispose()
End Sub

Private Sub frmFixRecv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

 Me.WindowState = FormWindowState.Maximized     '���¢�Ҵ���˹�Ҩ�
     StdDateTimeThai()                           '���¡ �Ѻ�ٷչ StdDateTimeThai
     tlsBarFmr.Cursor = Cursors.Hand             '����������ç Toolstripbar ���ٻ���

     dubNumberStart = 1                          '�������á� Recordset = 1
     dubNumberEnd = 2100                         '�������á� Recordset = 2100

     PreGroupType()

     InputData()
     tabCmd.Focus()

End Sub

Private Sub PreGroupType()
  Dim strGpTopic(4) As String
  Dim i As Byte

      strGpTopic(0) = "�����觫���"
      strGpTopic(1) = "�����ػ�ó�"
      strGpTopic(2) = "��������´�ػ�ó�"
      strGpTopic(3) = "������ػ�ó�"
      strGpTopic(4) = "ʶҹ��觫���"

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

Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchtxt As String)

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
                                 strSqlFind = "fix_id "
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 2
                                 strSqlFind = "eqp_id "
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 3
                                 strSqlFind = "eqp_name"
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 4
                                 strSqlFind = "desc_thai "
                                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

                          Case Is = 5
                                  strSqlFind = "fix_sta"
                                  strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

           End Select

           strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                        & " WHERE " & strSqlFind _
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

                            '---------------------------------------���Ң�����-------------------------------------------------------------

                            .MoveFirst()
                            .Find(strSqlFind)
                            lblPage.Text = Str(.AbsolutePage)

                            '-------------------------------------------------------------------------------------------------------------

                            If .Fields("RowNumber").Value >= 2100 Then

                               dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                               dubNumberEnd = .Fields("RowNumber").Value + 1000

                            Else

                               dubNumberStart = 1
                               dubNumberEnd = 2100

                            End If

                                strSqlFindData = strSqlFind

                                InputData()

                                       For i = 0 To dgvFix.Rows.Count - 1

                                            If InStr(UCase(dgvFix.Rows(i).Cells(2).Value), strSearchtxt.Trim.ToUpper) <> 0 Then
                                               dgvFix.CurrentCell = dgvFix.Item(2, i)
                                               dgvFix.Focus()
                                               Exit For

                                            End If

                                       Next i

                Else

                     MsgBox("����բ����� : " & strSearchtxt & " ��к�" & vbNewLine _
                                        & "�ô�кء�ä��Ң���������!", vbExclamation, "Not Found Data")

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

Private Sub StateLockFind(ByVal sta As Boolean)
  tabCmd.Enabled = sta
  dgvFix.Enabled = sta
  tlsBarFmr.Enabled = sta

End Sub

Private Sub InputData()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String = ""
 Dim strDateAdd As String = ""
 Dim strDateEdit As String = ""

 Dim strInDate As String = ""

 Dim intPageCount As Integer          '�ӹǹ˹�ҷ�����
 Dim intPageSize As Integer           '�ӹǹ��¡��� 1 ˹��
 Dim intCounter As Integer

 Dim strSearch As String = txtFilter.Text.ToString.Trim      '��ͧ������
 Dim strFieldFilter As String = ""

 Dim dteComputer As Date = Now()
 Dim imgStaFix As Image               '�ٻʶҹ��觫���

 Dim strDateFilter As String = ""
 Dim strYearCnvt As String = ""

       With Conn

            If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

          If blnHaveFilter Then          '�ó����͡ ��ͧ������

                    Select Case cmbFilter.SelectedIndex()

                               Case Is = 0
                                     strFieldFilter = "fix_id like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 1
                                      strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 2
                                      strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 3
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                               Case Is = 4
                                      strFieldFilter = "fix_sta like '%" & ReplaceQuote(strSearch) & "%'"

                    End Select

                    strSqlCmdSelc = "SELECT  * FROM v_fixeqptrn (NOLOCK)" _
                                                   & " WHERE " & strFieldFilter _
                                                   & " AND fix_sta = '2'" _
                                                   & " ORDER BY fix_id"
        Else

              strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                          & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                          & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                          & " AND fix_sta = '2'" _
                                          & " ORDER BY fix_id"

              End If

              intPageSize = 30   '����á�˹���Ҵ��д��

              Rsd = New ADODB.Recordset
              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )



                          If .RecordCount <> 0 Then

                                    If intPageSize > .RecordCount Then    '��Ҩӹǹ��¡��� 1 page(30) > �ӹǹ�ä���촷�� qurey ��
                                        intPageSize = .RecordCount
                                    End If

                                    If intPageSize = 0 Then
                                        intPageSize = 30
                                    End If

                                    .PageSize = intPageSize        '.PageSize ���˹��������˹�Ҩ�����ա����¡�� 㹡���ʴ���
                                     intPageCount = .PageCount     '.PageCount �Ѻ�ӹǹ˹�ҷ����� �����ҡ��á�˹���Ҵ�ͧ˹��


                                    '--------------------------�ó��ա�ä���-----------------------------------

                                     If strSqlFindData <> "" Then

                                            .MoveFirst()
                                            .Find(strSqlFindData)

                                             If Not .EOF Then
                                                lblPage.Text = Str(.AbsolutePage)    '.AbsolutePage ����ҧ�ԧ��ѧ˹�ҷ���ͧ���
                                             End If

                                            strSqlFindData = ""

                                     End If


                                    '---------- ��˹����� � tlsBarFmr ----------------------------------------

                                    If Int(lblPage.Text.ToString) > intPageCount Then
                                        lblPage.Text = intPageCount.ToString
                                    End If

                                    txtPage.Text = lblPage.Text.ToString
                                    intBkPageCount = .PageCount
                                    lblPageAll.Text = "/ " & .PageCount.ToString
                                    .AbsolutePage = Int(lblPage.Text.ToString)

                                    dgvFix.Rows.Clear()            '������ Gridview ��͹ Inputdata

                                    intCounter = 0

                                    Do While Not .EOF

                                 '--------------------------------------- ʶҹ��觫��� --------------------------------------------------

                                            Select Case .Fields("fix_sta").Value.ToString.Trim

                                                   Case Is = "1"     '�觫���
                                                        imgStaFix = My.Resources.sign_deny

                                                   Case Is = "2"     '�Ѻ�׹�觫���
                                                        imgStaFix = My.Resources.accept

                                                   Case Else         '����
                                                        imgStaFix = My.Resources.blank

                                            End Select

                                            dgvFix.Rows.Add( _
                                                                imgStaFix, _
                                                                .Fields("fix_desc").Value.ToString.Trim, _
                                                                .Fields("fix_id").Value.ToString.Trim, _
                                                                .Fields("eqp_id").Value.ToString.Trim, _
                                                                "#" & .Fields("size_id").Value.ToString.Trim, _
                                                                .Fields("issue").Value.ToString.Trim, _
                                                                .Fields("fix_issue").Value.ToString.Trim, _
                                                                .Fields("fix_amount").Value.ToString.Trim, _
                                                                Format(.Fields("fix_price").Value, "##,##0.00"), _
                                                                 Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                                                 Mid(.Fields("due_date").Value.ToString.Trim, 1, 10), _
                                                                 Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("recv_by").Value.ToString.Trim, _
                                                                 Mid(.Fields("pre_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("pre_by").Value.ToString.Trim, _
                                                                 Mid(.Fields("last_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("last_by").Value.ToString.Trim _
                                                             )

                                              intCounter = intCounter + 1

                                              If intCounter = intPageSize Then
                                                    Exit Do
                                              End If

                                         .MoveNext()            '����价������¹����
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

Private Sub InputGpbRecv()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String = ""
 Dim strDateAdd As String = ""
 Dim strDateEdit As String = ""

 Dim strInDate As String = ""

 Dim dteComputer As Date = Now()
 Dim imgStaFix As Image               '�ٻʶҹ��觫���

 Dim strDateFilter As String = ""
 Dim strYearCnvt As String = ""
 Dim ChkboxSta As Boolean

       With Conn

            If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

              strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                          & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                          & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                          & " AND fix_sta = '1'" _
                                          & " ORDER BY fix_id"

              Rsd = New ADODB.Recordset
              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )



                          If .RecordCount <> 0 Then

                                    dgvShow.Rows.Clear()      '������ datagrid

                                    Do While Not .EOF

                                    '--------------------------- ʶҹ��觫��� ------------------------------

                                            Select Case .Fields("fix_sta").Value.ToString.Trim

                                                   Case Is = "1"     '�觫���
                                                        imgStaFix = My.Resources.sign_deny
                                                        ChkboxSta = False     '

                                                   Case Is = "2"     '�Ѻ�׹�觫���
                                                        imgStaFix = My.Resources.accept
                                                        ChkboxSta = True     '��� chekbox �١���͡

                                                   Case Else         '����
                                                        imgStaFix = My.Resources.blank

                                            End Select

                                              dgvShow.Rows.Add( _
                                                                ChkboxSta, _
                                                                imgStaFix, _
                                                                .Fields("fix_desc").Value.ToString.Trim, _
                                                                .Fields("fix_id").Value.ToString.Trim, _
                                                                .Fields("eqp_id").Value.ToString.Trim, _
                                                                .Fields("eqp_name").Value.ToString.Trim, _
                                                                "#" & .Fields("size_id").Value.ToString.Trim, _
                                                                .Fields("sup_name").Value.ToString.Trim, _
                                                                 Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                                                .Fields("recv_by").Value.ToString.Trim _
                                                              )

                                         .MoveNext()            '����价������¹����
                                     Loop

                            End If

                   .Close()

              End With

              Rsd = Nothing

  Conn.Close()
  Conn = Nothing
End Sub

Private Function chkfxtranc() As Boolean    '������ fixeqptrn fix_sta = 1  �������
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlCmdSelc As String = ""

       With Conn

            If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

              strSqlCmdSelc = "SELECT * FROM fixeqptrn (NOLOCK)" _
                                          & " WHERE fix_sta = '1'"

              Rsd = New ADODB.Recordset

              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                        If .RecordCount <> 0 Then
                           Return True

                        Else
                           Return False

                        End If

             .ActiveConnection = Nothing
             .Close()
             End With

  Conn.Close()
  Conn = Nothing
End Function

Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

 Dim btnReturn As Boolean
 Dim strFixSta As String = ""           '��ʶҹ��觫���

     With tabCmd

         Select Case tabCmd.SelectedIndex

                Case Is = 0  '�Ѻ����觫���

                     If chkfxtranc() Then       '������ fixeqptrn fix_sta = 1  �������

                           IsShowSeek = Not IsShowSeek
                           If IsShowSeek Then

                               With gpbfxrecv             'groupbox �Ѻ����ػ�ó�

                                   .Visible = True
                                   .Left = 285
                                   .Top = 200
                                   .Height = 347
                                   .Width = 795

                                   dgvShow.Rows.Clear()
                                   InputGpbRecv()          '����Ң�������觫���
                                   StateLockFindDept(False)

                               End With

                            Else
                                 StateLockFindDept(True)

                            End If

                     End If


                Case Is = 1  '��䢢�����

                    If dgvFix.Rows.Count <> 0 Then

                         btnReturn = CheckUserEntry(strDocCode, "act_edit")
                         If btnReturn Then

                            ClearTmpTableUser("tmp_fixeqptrn")
                            lblCmd.Text = "1"                     '���͡�˹�����繡�����

                            With frmAeFixRecv
                                 .Show()
                                 .Text = "��䢢�����"

                            End With

                            Me.Hide()
                            frmMainPro.Hide()

                         Else
                            MsnAdmin()
                         End If

                    End If


                Case Is = 2    '����ͧ

                   If dgvFix.Rows.Count <> 0 Then

                                btnReturn = CheckUserEntry(strDocCode, "act_view")
                                If btnReturn Then
                                   ViewShoeData()

                                Else
                                    MsnAdmin()

                                End If

                   End If


                Case Is = 3   '��ͧ������

                     If dgvFix.Rows.Count <> 0 Then

                        With gpbFilter

                        .Top = 230
                        .Left = 210
                        .Width = 348
                        .Height = 125

                        .Visible = True

                        cmbFilter.Text = cmbFilter.Items(0)
                        txtFilter.Text = _
                                   dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(2).Value.ToString.Trim

                        StateLockFind(False)
                        txtFilter.Focus()

                        End With

                     End If


                Case Is = 4   '���Ң�����

                     If dgvFix.Rows.Count <> 0 Then

                        With gpbSearch

                             .Top = 230
                             .Left = 210
                             .Width = 348
                             .Height = 125

                             .Visible = True

                             cmbType.Text = cmbType.Items(0)
                             txtSeek.Text = _
                                     dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(2).Value.ToString.Trim

                             StateLockFind(False)
                             txtSeek.Focus()

                        End With

                     End If


                Case Is = 6            '����������

                    If dgvFix.Rows.Count > 0 Then

                        ClearTmpTableUser("tmp_fixeqptrn")

                        frmMainPro.lblRptCentral.Text = "G"     ' �觤���������ÿ���� MainPro 

                        '------------------------- �觤��������� lblRptDesc �ͧ����� MainPro ���� Userid �Ѻ Eqpid ----------------------------- 

                        frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim & "'"

                        frmRptCentral.Show()

                        StateLockFind(True)
                        frmMainPro.Hide()

                   Else
                        MsgBox("�ô�кآ����š�͹�����", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Equipment Empty!!")

                   End If

               Case Is = 5           '��鹿٢�����
                    InputData()

               Case Is = 7           'ź������

                      btnReturn = CheckUserEntry(strDocCode, "act_delete")
                            If btnReturn Then
                                DeleteData()
                            Else
                                MsnAdmin()
                            End If

               Case Is = 8           '�͡
                    Me.Close()

         End Select

    End With
End Sub

Private Sub dgvShow_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShow.CellClick
  On Error Resume Next
     With dgvShow

          If .Rows.Count <> 0 Then

               Select Case .CurrentCell.ColumnIndex

                      Case Is = 0

                          If dgvShow.Rows(e.RowIndex).Cells("recv").Value = False Then          '��Ǩ�ͺ��� chekbox�١���͡��������������� (�ó��ѧ���١���͡)

                                If Convert.ToBoolean(dgvShow.Rows(e.RowIndex).Cells("recv").Value) = True Then      '0 = false 1 = true

                                   dgvShow.Rows(e.RowIndex).Cells("recv").Value = False

                                Else

                                    With frmAeFixRecv
                                         .Show()
                                         lblCmd.Text = "0"    '�觺͡����繡���Ѻ���

                                    End With
                                    Me.Hide()
                                    frmMainPro.Hide()

                                End If

                           StateLockFindDept(True)
                           IsShowSeek = False
                           gpbfxrecv.Visible = False


                           End If

                End Select

          End If

     End With

End Sub

Private Sub UpdateData()                         ' UPDATE ʶҹ� fix_sta �ҡ 1='�觫���' �� 2='�Ѻ�׹�觫���'
 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim strDate As Date = Now()
 Dim strFixID As String
 Dim strEqpID As String
 Dim strSize As String

     strFixID = dgvFix.Rows(dgvFix.CurrentCell.ColumnIndex).Cells(2).Value.ToString
     strEqpID = dgvFix.Rows(dgvFix.CurrentCell.ColumnIndex).Cells(3).Value.ToString
     strSize = dgvFix.Rows(dgvFix.CurrentCell.ColumnIndex).Cells(4).Value.ToString

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

                 strSqlCmd = "UPDATE fixeqpmst SET fix_sta = '2'" _
                                & " WHERE fix_id = '" & strFixID & "'" _
                                & " AND eqp_id = '" & strEqpID & "'"


                 Conn.Execute(strSqlCmd)

                 strSqlCmd = "UPDATE fixeqptrn SET fix_sta = '2'" _
                                & " WHERE fix_id = '" & strFixID & "'" _
                                & " AND eqp_id = '" & strEqpID & "'" _
                                & " AND size_id = '" & strSize & "'"

                 Conn.Execute(strSqlCmd)


 Conn.Close()
 Conn = Nothing

End Sub

Private Sub dgvFix_CellMouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvFix.CellMouseDoubleClick
 ViewShoeData()
End Sub

Private Sub dgvFix_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvFix.KeyDown

 If e.KeyCode = Keys.Enter Then
     e.Handled = True
  End If

End Sub

Private Sub dgvFix_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvFix.KeyPress

  Dim blnReturn As Boolean
      If e.KeyChar = Chr(13) Then

         blnReturn = CheckUserEntry(strDocCode, "act_view")
         If blnReturn Then
            ViewShoeData()
         End If

      End If

End Sub

Private Sub ViewShoeData()

  If dgvFix.Rows.Count <> 0 Then

     ClearTmpTableUser("tmp_fixeqptrn")
     lblCmd.Text = "2"

     With frmAeFixRecv
          .Show()
          .Text = "����ͧ������"

     End With

     Me.Hide()
     frmMainPro.Hide()

  Else
     MsnAdmin()
  End If

End Sub

Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
  lblPage.Text = "1"
  InputData()
End Sub

Private Sub btnPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPre.Click
  If Int(lblPage.Text) > 1 Then
     lblPage.Text = Str(Int(lblPage.Text) - 1)
     InputData()
  End If
End Sub

Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
 lblPage.Text = Str(Int(lblPage.Text) + 1)
 InputData()
End Sub

Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
 lblPage.Text = Str(intBkPageCount)
 InputData()
End Sub

Private Sub txtPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPage.KeyPress

 If e.KeyChar = Chr(13) Then
   dgvFix.Focus()
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
                                    InputData()

                            Else

                                    lblPage.Text = intMovePage.ToString.Trim
                                    txtPage.Text = lblPage.Text
                                    InputData()

                            End If
                    Else

                        If intMovePage > 0 Then
                            lblPage.Text = intMovePage.ToString.Trim
                            txtPage.Text = lblPage.Text
                        Else
                            lblPage.Text = "1"
                            txtPage.Text = lblPage.Text
                        End If

                        InputData()

                    End If

                Catch ex As Exception
                    txtPage.Text = lblPage.Text
                End Try
End Sub

Private Sub dgvFix_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFix.RowsAdded
  dgvFix.Rows(e.RowIndex).Height = 27
End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  SearchDT()
End Sub

Private Sub SearchDT()                                        '�����͡���
 Dim strSearch As String = txtSeek.Text.ToUpper.Trim

 If strSearch <> "" Then

           Select Case cmbType.SelectedIndex()

                  Case Is = 0 '�����觫���
                          SearchData(0, strSearch)             '�觵����͹� ,Text ��� �Ѻ�ٷչ SearchData

                  Case Is = 1 '�����ػ�ó�
                          SearchData(2, strSearch)

                  Case Is = 2 '��������´�ػ�ó�
                          SearchData(3, strSearch)

                  Case Is = 3 '������ػ�ó�
                          SearchData(4, strSearch)

                  Case Is = 4 'ʶҹ��觫���
                          SearchData(5, strSearch)

          End Select

 Else
       MsgBox("�ô��͡���������ͤ���!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
       txtSeek.Focus()

 End If
End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
 StateLockFind(True)
 gpbSearch.Visible = False
End Sub

Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
  FindDocID()
End Sub

Private Sub FindDocID()     '�����͡���

 Dim strSearch As String = txtFilter.Text.ToUpper.Trim

     If strSearch <> "" Then

           Select Case cmbFilter.SelectedIndex()

                  Case Is = 0 '�����觫���
                          SearchData(0, strSearch)     '�觵����͹� ,Text ��� �Ѻ�ٷչ SearchData

                  Case Is = 1 '�����ػ�ó�
                          SearchData(2, strSearch)

                  Case Is = 2 '��������´�ػ�ó�
                          SearchData(3, strSearch)

                  Case Is = 3 '������ػ�ó�
                          SearchData(4, strSearch)

                  Case Is = 4 'ʶҹ��觫���
                          SearchData(5, strSearch)

          End Select

    Else
         MsgBox("�ô��͡���������ͤ���!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
         txtSeek.Focus()

    End If

End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click
   StateLockFind(True)
   gpbFilter.Visible = False
End Sub

Private Sub DeleteData()
 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim btyConsider As Byte
 Dim strFixID As String
 Dim strEqpID As String
 Dim strSize As String
 Dim strPrice As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

     With dgvFix

          If .Rows.Count <> 0 Then

             strFixID = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim
             strEqpID = .Rows(.CurrentRow.Index).Cells(3).Value.ToString.Trim
             strSize = .Rows(.CurrentRow.Index).Cells(4).Value.ToString.Trim
             strPrice = .Rows(.CurrentRow.Index).Cells(8).Value.ToString.Trim
             strSize = Mid(strSize, 2)          '�Ѵ # �͡

             btyConsider = MsgBox("�����ػ�ó�: " & strEqpID & vbNewLine _
                                               & "Size : " & strSize & vbNewLine _
                                               & "�س��ͧ���ź���������!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "�׹�ѹ���ź������")

                If btyConsider = 6 Then

                       If chkFixData(strFixID) Then   '��Ң������ fixeqptrn ����§ 1 �ä����

                          Conn.BeginTrans()

                          '---------------------------- ź���ҧ fixeqptrn --------------------------------------------

                          strSqlCmd = "DELETE FROM fixeqptrn" _
                                               & " WHERE fix_id ='" & strFixID & "'" _
                                               & " AND size_id = '" & strSize & "'"

                          Conn.Execute(strSqlCmd)
                          Conn.CommitTrans()

                         .Rows.RemoveAt(.CurrentRow.Index)


                         '------------------------------------ ź���ҧ fixeqpmst ----------------------------------------

                          strSqlCmd = "DELETE FROM fixeqpmst" _
                                                 & " WHERE fix_id ='" & strFixID & "'"

                          Conn.Execute(strSqlCmd)

                          InputData()  '�Ѿഷ������� datagrid


                       Else

                          Conn.BeginTrans()

                          ChangFixPrice()  'update fix_price �ء���������ź ������

                          '---------------------------- ź���ҧ fixeqptrn --------------------------------------------

                          strSqlCmd = "DELETE FROM fixeqptrn" _
                                               & " WHERE fix_id ='" & strFixID & "'" _
                                               & " AND size_id = '" & strSize & "'"

                          Conn.Execute(strSqlCmd)
                          Conn.CommitTrans()

                         .Rows.RemoveAt(.CurrentRow.Index)
                          InputData()  '�Ѿഷ������� datagrid

                       End If

                End If

          End If
          .Focus()
     End With

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub ChangFixPrice()  'update fix_price �ء���������ź ������
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlSelc As String
 Dim strSqlcmd As String
 Dim fxprice As Double
 Dim Sumprice As Double

 Dim strFixID As String
 Dim strEqpID As String
 Dim strSize As String


     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

     With dgvFix
          strFixID = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim
          strEqpID = .Rows(.CurrentRow.Index).Cells(3).Value.ToString.Trim
          strSize = .Rows(.CurrentRow.Index).Cells(4).Value.ToString.Trim
          fxprice = .Rows(.CurrentRow.Index).Cells(8).Value

     End With

            strSqlSelc = " SELECT * FROM fixeqpmst" _
                                 & " WHERE fix_id = '" & strFixID & "'"

             Rsd = New ADODB.Recordset
             With Rsd

                     .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                     .LockType = ADODB.LockTypeEnum.adLockOptimistic
                     .Open(strSqlSelc, Conn, , , )

                     If .RecordCount <> 0 Then

                        Sumprice = .Fields("fix_price").Value
                        Sumprice = Sumprice - fxprice

                        strSqlcmd = "UPDATE fixeqpmst SET fix_price = '" & Sumprice & "'" _
                                                  & "WHERE fix_id = '" & strFixID & "'"

                        Conn.Execute(strSqlcmd)
                     End If

             .ActiveConnection = Nothing
             .Close()
             End With

    Conn.Close()
    Conn = Nothing
End Sub

Private Function chkFixData(ByVal txtFixid As String) As Boolean        '�礢������ fixeqptrn �������� �ä�����ش�����������
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlSelc As String

       With Conn

           If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

              strSqlSelc = "SELECT * FROM fixeqptrn (NOLOCK)" _
                                        & " WHERE fix_id = '" & txtFixid & "'"


              Rsd = New ADODB.Recordset

              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlSelc, Conn, , , )

                        If .RecordCount = 1 Then          '����� record �ش���� ��͹ź
                            Return True

                        Else
                            Return False

                        End If

             .ActiveConnection = Nothing
             .Close()
             End With

 Conn.Close()
 Conn = Nothing

End Function

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  InputData()
End Sub

Private Sub lblGpbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblGpbClose.Click
  With gpbfxrecv

       .Visible = False
        IsShowSeek = False

  End With

  StateLockFindDept(True)

End Sub

Private Sub StateLockFindDept(ByVal sta As Boolean)
  tabCmd.Enabled = sta
  dgvFix.Enabled = sta
  tlsBarFmr.Enabled = sta
End Sub

Private Sub txtSeek_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSeek.LostFocus
 txtSeek.Text = txtSeek.Text.ToUpper.Trim
End Sub

End Class