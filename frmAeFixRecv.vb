Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeFixRecv
 Dim IsShowSeek As Boolean        '�������ʴ�ʶҹ� gpbSeek
 Dim strDateDefault As String     '���������Ѻ�ѹ�������

Protected Overrides ReadOnly Property CreateParams() As CreateParams       '��ͧ�ѹ��ûԴ������� Close Button(�����ҡ�ҷ)
    Get
         Dim cp As CreateParams = MyBase.CreateParams
         Const CS_DBLCLKS As Int32 = &H8
         Const CS_NOCLOSE As Int32 = &H200
         cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
         Return cp
    End Get

End Property

Private Sub frmAeFixRecv_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearTmpTable(0, "")        'ź������ Table tmp_fixeqptrn
  frmFixRecv.lblCmd.Text = "0"  '������ʶҹ�
  Me.Dispose()                '����¿���� �׹˹��¤�����
End Sub

Private Sub frmAeFixRecv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Dim dtComputer As Date = Now       '������纤���ѹ���Ѩ�غѹ
  Dim strCurrentDate As String       '�纤��ʵ�ԧ�ѹ���Ѩ�غѹ

      Me.WindowState = FormWindowState.Maximized  '��������������˹�Ҩ�
      StdDateTimeThai()        '�Ѻ�ٷչ��ŧ�ѹ�����Ǵ�.�� ����� Module
      strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

      PrePartSeek()
      ClearDataGpbHead()

      Select Case frmFixRecv.lblCmd.Text.ToString

             Case Is = "0"   '�Ѻ����觫���
                  LockEditData()
                  CallData()            ' ���¡��������������´�觫���

             Case Is = "1"   '���
                  LockEditData1()
                  CallData1()          ' ���¡��������������´�觫���

             Case Is = "2"   '����ͧ
                  LockEditData1()
                  CallData1()          ' ���¡��������������´�觫���
                  btnSaveData.Enabled = False

       End Select

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  On Error Resume Next    '�ҡ�Դ error ������ѧ�зӧҹ���������ʹ� error ����Դ���
  Dim strCode As String

        If MessageBox.Show("��ͧ����͡�ҡ����� �������", "��س��׹�ѹ�͡�ҡ�����", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            With frmFixRecv.dgvFix
                If .Rows.Count > 0 Then   '����բ������ Grid
                    strCode = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim          '���strCode = ��������ǻѨ�غѹ Cell �á
                    lblComplete.Text = strCode  '��� label �ʴ�������� Cell �Ѩ�غѹ   
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

     strEqptype(0) = "���촩մ EVA INJECTION"
     strEqptype(1) = "���촩մ PVC INJECTION"
     strEqptype(2) = "������ʹ PU"
     strEqptype(3) = "����ἧ�Ѵ���˹ѧ˹��,���"
     strEqptype(4) = "�մ�Ѵ"
     strEqptype(5) = "���͡ʡ�չ"
     strEqptype(6) = "���͡����"

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
 Dim strSqlSelc As String = ""                          ' ��ʵ�ԧ sql select

 Dim strSqlCmd As String = ""                           ' ��ʵ�ԧ Command
 Dim blnHavedata As Boolean                             ' �纤�ҵ����� ����Ѻ������բ������������

 Dim strPart As String = ""
 Dim strFixid As String
 Dim strSize As String
 Dim strGpType As String = ""

     strFixid = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(3).Value.ToString
     strSize = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(6).Value.ToString
     strSize = Mid(strSize, 2, 5)        '�� # �͡.
    
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

            '------------------------------- ���������ŧ㹵��ҧ tmp_fixeqptrn ----------------------------

             strSqlSelc = "INSERT INTO tmp_fixeqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE fix_id = '" & strFixid & "'"

             Conn.Execute(strSqlSelc)

             blnHavedata = True                '�觺͡����բ�����
             lockGpbHead(False)                '��ͤ Groupbox Head

         Else

             blnHavedata = False

         End If
         .ActiveConnection = Nothing           '��� ReccordSet
         .Close()

     End With
     Rsd = Nothing
  Conn.Close()
  Conn = Nothing

End Sub

Private Sub LockEditData1()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String = ""                          ' ��ʵ�ԧ sql select

 Dim strSqlCmd As String = ""                           ' ��ʵ�ԧ Command
 Dim blnHavedata As Boolean                             ' �纤�ҵ����� ����Ѻ������բ������������

 Dim strPart As String = ""
 Dim strFixid As String
 Dim strSize As String
 Dim strGpType As String = ""

     strFixid = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(2).Value.ToString
     strSize = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(4).Value.ToString
     strSize = Mid(strSize, 2)        '�� # �͡.
     
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

            '------------------------------- ���������ŧ㹵��ҧ tmp_fixeqptrn ----------------------------

             strSqlSelc = "INSERT INTO tmp_fixeqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE fix_id = '" & strFixid & "'" _
                                  & " AND size_id = '" & strSize & "'"

             Conn.Execute(strSqlSelc)

             blnHavedata = True                '�觺͡����բ�����
             lockGpbHead(False)                '��ͤ Groupbox Head

         Else
               MsgBox("����բ����,")
               blnHavedata = False

         End If
         .ActiveConnection = Nothing           '��� ReccordSet
         .Close()

     End With
     Rsd = Nothing
     Conn.Close()
     Conn = Nothing

End Sub

Private Function checkGroup(ByVal txtGroup As String) As String    '�ѧ�����ŧ  GroupSize
Dim GroupDesc As String = ""

    Select Case txtGroup

           Case Is = "A"
                 GroupDesc = "���촩մ EVA INJECTION"

           Case Is = "B"
                 GroupDesc = "���촩մ PVC INJECTION"

           Case Is = "C"
                 GroupDesc = "������ʹ PU"

           Case Is = "D"
                 GroupDesc = "����ἧ�Ѵ���˹ѧ˹��,���"

           Case Is = "E"
                 GroupDesc = "�մ�Ѵ"

           Case Is = "F"
                 GroupDesc = "���͡ʡ�չ"

           Case Is = "G"
                 GroupDesc = "���͡����"

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

             strEqpid = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(4).Value.ToString.Trim     '�����ػ�ó�
             strSize = frmFixRecv.dgvShow.Rows(frmFixRecv.dgvShow.CurrentRow.Index).Cells(6).Value.ToString.Trim     '�� Size
             strSize = Mid(strSize, 2)       '�Ѵ # ˹�� size �͡

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
                     lockFixdetail(False)          '��ͤ��������´�觫���

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


             strEqpid = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(3).Value.ToString.Trim     '�����ػ�ó�
             strSize = frmFixRecv.dgvFix.Rows(frmFixRecv.dgvFix.CurrentRow.Index).Cells(4).Value.ToString.Trim     '�� Size
             strSize = Mid(strSize, 2)       '�Ѵ # ˹�� size �͡


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
                     lockFixdetail(False)          '��ͤ��������´�觫���

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

                   'SPrice()          '�ѧ���ѹ�Ҽ������ҫ����ػ�ó�
                   SaveEditdata()    '�Ѿഷ������

                Else
                   MsgBox("�ô�к���������´��ë���  " _
                                 & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                   txtFixDetail.Focus()

                End If


        Else
            MsgBox("�ô�кؼ���Ѻ���  " _
                             & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
            txtRecvBy.Focus()

        End If

  Else
      MsgBox("�ô�к��ѹ����Ѻ���  " _
                            & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
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

                Select Case strMerg.ToString.Length    ' Check �������ʵ�ԧ
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

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                Case Is = 38 '�١�â��
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtRecvBy.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtRecvBy.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtFixDetail.Focus()
                Case Is = 113 '���� F2
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

            Case Is = 35 '���� End 

            Case Is = 36 '���� Home

            Case Is = 37 '�١�ë���

                 If .SelectionStart = 0 Then
                    txtRecvDate.Focus()
                 End If

            Case Is = 38 '�����١�â��

            Case Is = 39   '�����١�â��
                 If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                    txtRecvAmt.Focus()

                 Else

                    intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtRecvAmt.Focus()

                        End If
                 End If

            Case Is = 40 '����ŧ
                      txtFixDetail.Focus()
            Case Is = 113 '���� F2
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
                            If InStr("0123456789.", strTmp) > 0 Then    '����ʵ�ԧ
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

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

                    If .SelectionStart = 0 Then
                        txtRecvBy.Focus()
                    End If

                Case Is = 38 '�����١�â��  

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtFxPrice.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkpoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtFxPrice.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtFixDetail.Focus()
                Case Is = 113 '���� F2
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
                mskRecvAmt.Text = ""     '������ mskSizeQty
                z = CDbl(strMerg)        '�ŧ Type dbl
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

                Select Case strMerge.IndexOf(".") '�ҵ��˹觷�辺�繤����á

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
              Case Is = 35 '���� End 
              Case Is = 36 '���� Home
              Case Is = 37 '�١�ë���
                        If .SelectionStart = 0 Then
                            txtRecvAmt.Focus()
                        End If
              Case Is = 38   '�����١�â��

              Case Is = 39   '�����١�â��
                        If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                            txtFixDetail.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                            If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                               txtFixDetail.Focus()
                            End If
                        End If
             Case Is = 40 '����ŧ
                      txtFixDetail.Focus()
             Case Is = 113 '���� F2
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

            Case Is = 35 '���� End 

            Case Is = 36 '���� Home

            Case Is = 37 '�١�ë���
                 If .SelectionStart = 0 Then
                    txtFxPrice.Focus()
                 End If

            Case Is = 38 '�����١�â��
                 txtRecvBy.Focus()

            Case Is = 39   '�����١�â��
                 If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                    btnSaveData.Focus()
                 Else

                    intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           btnSaveData.Focus()

                        End If
                 End If

            Case Is = 40 '����ŧ

            Case Is = 113 '���� F2
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

               'Conn.BeginTrans()      '�ش������� Transection

               strDate = DateSave.Date.ToString("yyyy-MM-dd")
               strDate = SaveChangeEngYear(strDate)

              '------------------------- �ѹ����͡��� ----------------------------------------------------

               strDocDate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocDate = SaveChangeEngYear(strDocDate)

               '---------------------------------------- Ǵ�.�Ѻ�׹�觫��� -----------------------------------

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


                     '--------------------------------------- ź������㹵��ҧ eqptrn -----------------------------------

                     strSqlcmd = "DELETE FROM fixeqptrn" _
                                           & " WHERE fix_id ='" & lblFix_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)

                     '--------------------------------------- ����������㹵��ҧ fixEqptrn -------------------------------

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
                     'Conn.CommitTrans()                   '��� Commit transection

                     lblComplete.Text = txtSize.Text.ToString.Trim  '�觺͡��Һѹ�֡�����������
                     Me.Hide()

                           If ChkFixStatus() Then         '��Ǩ�ͺ����Ѻ����觫����ú�������
                              SaveEditFixmst()            '�Ѿഷ������ table fixeqpmst

                           End If

                     frmMainPro.Show()
                     frmFixRecv.Show()


  Conn.Close()
  Conn = Nothing

End Sub

Private Function ChkFixStatus() As Boolean     '��Ǩ�ͺʶҹС���觫���
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

      '------------------- ����Ѻ����觫��� �ú�ء size -----------------------

      strSqlSelc = "SELECT * " _
                        & " FROM tmp_fixeqptrn (NOLOCK)" _
                        & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "' " _
                        & " AND fix_sta = '1'"

      Rsd = New ADODB.Recordset

      With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(strSqlSelc, Conn, , , )

           If .RecordCount <> 0 Then         '����ѧ�Ѻ������ú (fix_sta = 1)
              Return False

           Else                              '�ó��Ѻ��� �ú��ǹ
              Return True

           End If

      End With
      Rsd = Nothing
      Rsd.Close()


  Conn = Nothing
  Conn.Close()

End Function

Private Sub SaveEditFixmst()    '�Ѿഷʶҹ��Ѻ����觫��� tb_fixeqpmst

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

            Conn.BeginTrans()      '�ش������� Transection

            strDate = DateSave.Date.ToString("yyyy-MM-dd")
            strDate = SaveChangeEngYear(strDate)

            '------------------------- �ѹ����͡��� ----------------------------------------------------

            strDocDate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
            strDocDate = SaveChangeEngYear(strDocDate)

            '---------------------------------------- Ǵ�.�Ѻ�׹�觫��� -----------------------------------

            strRecvDate = Mid(txtRecvDate.Text.ToString.Trim, 7, 4) & "-" _
                            & Mid(txtRecvDate.Text.ToString.Trim, 4, 2) & "-" _
                            & Mid(txtRecvDate.Text.ToString.Trim, 1, 2)
            strRecvDate = SaveChangeEngYear(strRecvDate)


            Select Case cmbType.Text

                   Case Is = "���촩մ EVA INJECTION"
                           strGpType = "A"

                   Case Is = "���촩մ PVC INJECTION"
                           strGpType = "B"

                   Case Is = "������ʹ PU"
                           strGpType = "C"

                   Case Is = "����ἧ�Ѵ���˹ѧ˹��,���"
                           strGpType = "D"

                   Case Is = "�մ�Ѵ"
                           strGpType = "E"

                   Case Is = "���͡ʡ�չ"
                           strGpType = "F"

                   Case Is = "���͡����"
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

                     '--------------- �Ѿഷʶҹ� eqpmst ��  2(�Ѻ��Ѻ�觫���)-------------------------------------------------

                     strSqlcmd = " UPDATE eqpmst SET fix_sta = '2'" _
                                            & "," & "last_date = '" & strDate & "'" _
                                            & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)

                     '---------------- �Ѿഷʶҹ���觫��� eqptrn fix_sta = 2 (�Ѻ��Ѻ�觫���) ------------------------------------------

                     strSqlcmd = " UPDATE eqptrn SET fix_sta = '2'" _
                                            & "," & "last_date = '" & strDate & "'" _
                                            & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'" _
                                            & " AND size_id = '" & txtSize.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlcmd)

                     Conn.CommitTrans()               '��� Commit transection


   Conn.Close()
   Conn = Nothing
End Sub

Private Function SPrice() As Double        '�ѧ���ѹ�Ҽ������ҫ����ػ�ó�
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