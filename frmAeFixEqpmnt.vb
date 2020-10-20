Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeFixEqpmnt
 Dim IsShowSeek As Boolean        '�������ʴ�ʶҹ� gpbSeek
 Dim strDateDefault As String     '���������Ѻ�ѹ�������
 Dim IsShowSearch As Boolean

 Dim staAction As String

Protected Overrides ReadOnly Property CreateParams() As CreateParams       '��ͧ�ѹ��ûԴ������� Close Button(�����ҡ�ҷ)
   Get
       Dim cp As CreateParams = MyBase.CreateParams
           Const CS_DBLCLKS As Int32 = &H8
           Const CS_NOCLOSE As Int32 = &H200
           cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
           Return cp
   End Get
End Property

Private Sub frmAeFixEqpmnt_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTable(0, "")  'ź������ Table tmp_eqptrn where user_id..
   frmFixEqpmnt.lblCmd.Text = "0"  '������ʶҹ�
   Me.Dispose()
End Sub

Private Sub frmAeFixEqpmnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

   Dim DateCom As Date = Now
   Dim strCurrentDate As String
   Dim strSize As String = ""
   Dim strRecvSize As String = ""

       StdDateTimeThai()
       strCurrentDate = DateCom.Date.ToString("dd/MM/yyyy")
        strSize = Mid(frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(4).Value.ToString.Trim, 2)

        ClearDataGpbHead()
       PrePartSeek()

       Select Case frmFixEqpmnt.lblCmd.Text.Trim

              Case Is = "0"           '�ó�����������

                   With Me
                        .Text = "����������"
                   End With

                   With txtBegin
                        .Text = strCurrentDate
                        strDateDefault = strCurrentDate
                   End With

                   GenFixID()           '���ҧ�����觫��� Fix_ID 
                   gpbReceive.Visible = False

              Case Is = "1"           '��䢢�����

                   With Me
                        .Text = "���䢢�����"
                   End With

                   LockEditData()             '��Ŵ��������������´�ػ�ó�
                   staAction = "3"
                   LoadEditData(strSize)      '��Ŵ�����Ţ��������
                   ShowGroupEdit()            '��ʴ� groupbox ����
                   gpbSCsub.Visible = False

                   gpbFix.Enabled = True
                   gpbReceive.Enabled = True
                   ShowReceived()              '��Ŵ�������Ѻ��� - �觫���

                   'If CheckHaveData() Then          '����Ѻ����ػ�ó���������ʴ��������Ѻ���..
                   '   gpbFix.Enabled = False
                   '   gpbReceive.Visible = True
                   '   gpbReceive.Enabled = True
                   '   ShowReceived()                '��Ŵ�������Ѻ��� - �觫���
                   'Else
                   '   gpbFix.Enabled = True
                   '   gpbReceive.Visible = False
                   'End If

                  txtEqp_id.ReadOnly = True         '�����ҹ���ҧ����
                  txtEqpnm.ReadOnly = True
                  btnSaveData.Enabled = True

              Case Is = "2"   '����ͧ������
                   With Me
                        .Text = "����ͧ������"
                   End With

                   LockEditData()             '��Ŵ��������������´�ػ�ó�
                   'loadFixdata()             '��Ŵ�����š���觫���
                   LoadEditData(strSize)      '��Ŵ�����Ţ��������
                   ShowGroupEdit()            '��ʴ� groupbox ����
                   gpbSCsub.Visible = False
                   '----------------

                   If CheckHaveData() Then          '����Ѻ����ػ�ó���������ʴ��������Ѻ���..
                      ShowReceived()      '��Ŵ�������Ѻ��� - �觫���
                   Else
                      gpbReceive.Visible = False
                   End If

                  txtEqp_id.ReadOnly = True  '�����ҹ���ҧ����
                  btnSaveData.Enabled = False

              Case Is = "3"   '�Ѻ�׹�觫���

                   With Me
                        .Text = "�Ѻ�׹�觫���"
                   End With

                   strRecvSize = frmFixEqpmnt.dgvShow.Rows(frmFixEqpmnt.dgvShow.CurrentRow.Index).Cells(4).Value.ToString.Trim
                   LoadDataDetail()       '��Ŵ��������������´�ػ�ó�
                   'LoadDataReceive()      '��Ŵ�����š���觫���  

                   '-----------------
                   LoadEditData(strRecvSize)     '��Ŵ�����š���觫���  
                   ShowGroupEdit()             '��ʴ� groupbox ����
                   gpbSCsub.Visible = False
                   '----------------

                   If CheckRemainEqp() Then          '��������ػ�ó��ҧ�Ѻ�׹��������
                      gpbReceive.Visible = True      '�ʴ� groupbox �Ѻ�׹�觫���
                      ShowReceived()
                   End If

                   gpbFix.Enabled = False
                   gpbReceive.Visible = True  '�ʴ� groupbox �Ѻ�׹�觫���
                   txtRecv_date.Focus()
                   txtEqp_id.ReadOnly = True  '�����ҹ���ҧ����
                   btnSaveData.Enabled = True

        End Select

End Sub

Private Function CheckRemainEqp() As Boolean
  Dim conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlSelc As String

     With conn


          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 150
             .Open()

     End With

       strSqlSelc = "SELECT * FROM fixeqptrn(NOLOCK)" _
                                & " WHERE fix_sta = '3'"

       Rsd = New ADODB.Recordset

       With Rsd

            .CursorType = CursorTypeEnum.adOpenKeyset
            .LockType = LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, conn, , )

            If .RecordCount > 0 Then
                Return True

            Else
                Return False
            End If

        .ActiveConnection = Nothing
        .Close()
       End With

  conn.Close()
  conn = Nothing

End Function

Private Sub ShowReceived()

 Dim conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim strSize As String

    ' If Me.Text = "���䢢�����" Then
         strSize = Mid(frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(3).Value.ToString, 2)
    ' Else
    '       strSize = frmFixEqpmnt.dgvShow.Rows(frmFixEqpmnt.dgvShow.CurrentRow.Index).Cells(4).Value.ToString
    ' End If

     With conn
          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 150
             .Open()
     End With

       strSqlSelc = "SELECT * FROM v_fixeqptrn(NOLOCK)" _
                             & " WHERE fix_id = '" & lblFix_id.Text.Trim & "'" _
                             & " AND size_id = '" & strSize & "'"

       With Rsd

            .CursorType = CursorTypeEnum.adOpenKeyset
            .LockType = LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, conn, , )

            If .RecordCount > 0 Then

                If .Fields("recv_date").Value.ToString <> "" Then
                   txtRecv_date.Text = Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10)

                Else
                   txtRecv_date.Text = "__/__/____"
                End If

                txtRecvNm.Text = .Fields("recv_by").Value.ToString.Trim
                'lblSumFx.Text = Format(.Fields("amt_out").Value, "##0.0")
                txtRecvTotal.Text = Format(.Fields("amt_in").Value, "##0.0")
                txtIssue.Text = .Fields("issue").Value.ToString.Trim
                txtFxIssue.Text = .Fields("fix_issue").Value.ToString.Trim

            End If

            strSize = ""   '�����������

        .ActiveConnection = Nothing
        .Close()
       End With

  conn.Close()
  conn = Nothing
End Sub

Private Function CheckHaveData() As Boolean            '���觫����ػ�ó�

 Dim conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

     With conn

          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 150
             .Open()

     End With

       strSqlSelc = "SELECT * FROM v_fixeqptrn(NOLOCK)" _
                                & " WHERE fix_id = '" & lblFix_id.Text.Trim & "'" _
                                & " AND fix_sta = '2'" _
                                & " OR fix_sta = '3'"

       With Rsd

            .CursorType = CursorTypeEnum.adOpenKeyset
            .LockType = LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, conn, , )

            If .RecordCount > 0 Then
               Return True

            Else
               Return False

            End If

        .ActiveConnection = Nothing
        .Close()
       End With

  conn.Close()
  conn = Nothing

End Function

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

Private Sub GenFixID()                     ' Gen �����觫���
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim LastRec As Integer
 Dim LastRec1 As Integer
 Dim LastRec2 As Integer
 Dim DateCom As Date = Now
 Dim strCurrentDate As String
 Dim THyear As String

     strCurrentDate = DateCom.Date.ToString("dd/MM/yyyy")

     If Year(strCurrentDate) > 2500 Then       '����繻վط�
        THyear = Mid(strCurrentDate, 9, 2) '�鴻��� 5X

     Else
         strCurrentDate = ShowChangeThaiYear(strCurrentDate)
         THyear = Mid(strCurrentDate, 9, 2) '�鴻��� 5X

     End If

      With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

      End With

      strSqlSelc = "SELECT * FROM fixeqpmst(NOLOCK)" _
                                   & " ORDER BY fix_id "


      With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

             .MoveLast()              '����͹��ѧ Record �ش����

             LastRec1 = CInt(Mid((.Fields("fix_id").Value.ToString.Trim), 5))  '�Ѵʵ�ԧ ��� 4 ��Ƿ���  000x
             LastRec2 = CInt(Mid((.Fields("fix_id").Value.ToString.Trim), 3))  '�Ѵʵ�ԧ FX �͡ 5x000x
             LastRec = Mid(CStr(LastRec2), 1, 2)  '�Ѵ��һ�  5x ੾�� 2����á

               If String.Compare(LastRec, THyear) = 0 Then       '���º��º ʵ�ԧ�� 5x
                  LastRec1 += 1  ' ������� LestRec �ա 1.

               Else
                  LastRec1 = 1
               End If

          Else
             LastRec1 = 1
          End If

          lblFix_id.Text = "FX" & THyear & LastRec1.ToString("0000")

      .ActiveConnection = Nothing
      .Close()
      End With

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub ClearDataGpbHead()
     txtEqp_id.Text = ""
     txtEqpnm.Text = ""
End Sub

Private Sub PrePartSeek()
 Dim strEqptype(7) As String
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

Private Sub LockEditData()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strCmd As String                                   '��ʵ�ԧ Command
 Dim blnHavedata As Boolean                             '�纤�ҵ����� ����Ѻ������բ������������
 Dim strSqlSelc As String = ""                          '��ʵ�ԧ sql select
 Dim strPart As String = ""

 Dim strFXid As String
        strFXid = frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(2).Value.ToString

        With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

        strSqlSelc = "SELECT * FROM fixeqpmst(NOLOCK) " _
                               & " WHERE fix_id = '" & strFXid & "'"

        With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then
             txtEqp_id.Text = .Fields("eqp_id").Value.ToString.Trim
             lblFix_id.Text = .Fields("fix_id").Value.ToString.Trim
             txtEqpnm.Text = .Fields("eqp_name").Value.ToString.Trim
             lblAmount.Text = Format(.Fields("amount").Value, "##0.0")
             lblAmt.Text = Format(.Fields("price").Value, "#,##0.00")

                Select Case .Fields("group").Value.ToString.Trim

                       Case Is = "A"
                            cmbType.Text = "���촩մ EVA INJECTION"

                       Case Is = "B"
                            cmbType.Text = "���촩մ PVC INJECTION"

                       Case Is = "C"
                            cmbType.Text = "������ʹ PU"

                       Case Is = "D"
                            cmbType.Text = "����ἧ�Ѵ���˹ѧ˹��,���"

                       Case Is = "E"
                            cmbType.Text = "�մ�Ѵ"

                       Case Is = "F"
                            cmbType.Text = "���͡ʡ�չ"

                       Case Else
                            cmbType.Text = "���͡����"

                End Select

                strCmd = frmFixEqpmnt.lblCmd.Text.ToString.Trim    '��� strCmd ��ҡѺ���� lblcmd 㹿���� frmEqpSheet

                Select Case strCmd

                       Case Is = "1"   '�����ͤ�͹���
                       Case Is = "2"   '�����ͤ�͹����ͧ
                            btnSaveData.Enabled = False  '�Դ���� "�ѹ�֡������"

                End Select

              '------------------------------- �ѹ�֡�����ŧ㹵��ҧ tmp_eqptrn ----------------------------

             strSqlSelc = "INSERT INTO tmp_fixeqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "', *" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE fix_id = '" & strFXid & "' "

              Conn.Execute(strSqlSelc)
              blnHavedata = True                     '�觺͡����բ�����
              StateLockFindDept(False)               'Disable groupBox Head

         Else
              blnHavedata = False
         End If

         .ActiveConnection = Nothing                  '��� ReccordSet
         .Close()

     End With

     Rsd = Nothing
  Conn.Close()
  Conn = Nothing

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
   Me.Close()
End Sub

 Private Sub btnSearchDT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchDT.Click
    SearchEqpid()
 End Sub

Private Sub CleargpbHead()
  txtEqp_id.Text = ""
  txtEqpnm.Text = ""
  cmbType.SelectedIndex = -1
  lblAmount.Text = "0"
  lblAmt.Text = "0.00"
End Sub

Private Sub SearchEqpid()
 IsShowSearch = Not IsShowSearch

 If IsShowSearch Then

    With gpbSeekEqp

         .BringToFront()
         .Visible = True
         .Left = 285
         .Top = 120
         .Height = 411
         .Width = 504
         dgvShow.Rows.Clear()
         CleargpbHead()               '��ҧ���������� gpbHead
         LoadData()

         txtSeek.Text = ""
         txtSeek.Focus()

    End With

    StateLockFindDept(False)
    dgvShow.Focus()

 Else
      StateLockFindDept(True)
      dgvShow.Focus()

 End If

End Sub

Private Sub LoadData()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlSelc As String = ""                          '��ʵ�ԧ sql select
 Dim strPart As String = ""

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

     strSqlSelc = "SELECT * " _
                           & "FROM eqpmst (NOLOCK)" _
                           & "ORDER BY eqp_id"

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

              dgvShow.Rows.Clear()
              dgvShow.ScrollBars = ScrollBars.None                  '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ

              Do While Not .EOF

                 dgvShow.Rows.Add( _
                                   .Fields("eqp_id").Value.ToString.Trim, _
                                   .Fields("eqp_name").Value.ToString.Trim, _
                                   "���͡", _
                                   .Fields("group").Value.ToString.Trim, _
                                   .Fields("remark").Value.ToString.Trim _
                                  )
                 .MoveNext()

              Loop

           Else

             MsgBox("��辺��������к�", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Not Found Data!!")
             txtSeek.Focus()

          End If

          dgvShow.ScrollBars = ScrollBars.Both                       '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ

     .ActiveConnection = Nothing          '��� ReccordSet
     .Close()
     End With
     Rsd = Nothing

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub StateLockFindDept(ByVal sta As Boolean)
 Dim strMode As String = frmFixEqpmnt.lblCmd.Text.ToString

     gpbHead.Enabled = sta
     btnSaveData.Enabled = sta  '�����ѹ�֡������

        Select Case strMode

               Case Is = "1" '��䢢�����                        
               Case Is = "2" '����ͧ������
                    btnSaveData.Enabled = False

        End Select
End Sub

Private Sub FindPsData(ByVal txtSearch As String)
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlSelc As String

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

     strSqlSelc = "SELECT * FROM eqpmst (NOLOCK)" _
                          & " WHERE eqp_id LIKE '%" & txtSearch & "%'"


     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )


          If .RecordCount <> 0 Then

              dgvShow.Rows.Clear()
              Do While Not .EOF

                 dgvShow.Rows.Add( _
                                   .Fields("eqp_id").Value.ToString.Trim, _
                                   .Fields("eqp_name").Value.ToString.Trim, _
                                   "���͡", _
                                   .Fields("group").Value.ToString.Trim, _
                                   .Fields("remark").Value.ToString.Trim _
                                   )
                 .MoveNext()

              Loop

           Else

             MsgBox("��辺������ :" & txtSearch & " ��к�" & vbNewLine _
                                          & "�ô�кء�ä��Ң���������!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Not Found Data!!")
             txtSeek.Focus()
          End If

     .ActiveConnection = Nothing
     .Close()
     End With

Conn.Close()
Conn = Nothing

End Sub

Private Sub btnSearchExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchExit.Click
  StateLockFindDept(True)
  gpbSeekEqp.Visible = False
  IsShowSeek = False
End Sub

Private Sub txtEqp_id_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqp_id.KeyDown
 Dim intChkPoint As Integer

     With txtEqp_id

          Select Case e.KeyCode

                 Case Is = 35 '���� End 
                 Case Is = 36 '���� Home
                 Case Is = 37 '�١�ë���
                 Case Is = 38 '�����١�â��
                 Case Is = 39 '�����١�â��
                      If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                         txtEqpnm.Focus()
                      Else

                       intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtEqpnm.Focus()
                        End If
                      End If

                Case Is = 40 '����ŧ
                        cmbType.DroppedDown = True
                        cmbType.Focus()
                Case Is = 113 '���� F2
                        .SelectionStart = .Text.Trim.Length
         End Select

     End With

End Sub

Private Function FindData(ByVal strSeek As String) As Boolean     '���������ػ�ó�
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlSelc As String

      With Conn

           If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()

      End With

      strSqlSelc = "SELECT * FROM eqpmst (NOLOCK)" _
                                     & "WHERE eqp_id = '" & strSeek & "'"


      With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelc, Conn, , , )

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

Private Sub txtEqpnm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqpnm.KeyDown
 Dim intChkPoint As Integer

    With txtEqpnm

     Select Case e.KeyCode

            Case Is = 35 '���� End 
            Case Is = 36 '���� Home
            Case Is = 37 '�١�ë���
                 If .SelectionStart = 0 Then
                    txtEqp_id.Focus()
                 End If
            Case Is = 38 '�����١�â��

            Case Is = 39   '�����١�â��
                 If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                    cmbType.DroppedDown = True
                    cmbType.Focus()
                 Else

                     intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           cmbType.DroppedDown = True
                           cmbType.Focus()
                        End If
                 End If

            Case Is = 40 '����ŧ
                      cmbType.DroppedDown = True
                      cmbType.Focus()
            Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
     End Select

  End With

End Sub

Private Sub txtEqpnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqpnm.KeyPress
   If e.KeyChar = Chr(13) Then
      cmbType.DroppedDown = True
      cmbType.Focus()
   End If
End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  If staAction = "3" Then     '�ó���䢢�����
     SaveEditBySize()         '�ѹ�֡����੾�� size ������͡��
  Else
      CheckDataBfSave()
  End If
End Sub

Private Sub SaveEditBySize()    '�ѹ�֡����੾�� size ������͡��

    Dim Conn As New ADODB.Connection
    Dim strSqlCmd As String
    Dim strSqlSelc As String
    Dim Rsd As New ADODB.Recordset

    Dim strFixdate As String
    Dim strDuedate As String

        With Conn

             If .State Then Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 150
                .Open()

                '-------------------------- Ǵ�.����觫��� -----------------------------

                    If txtFixdate.Text <> "__/__/____" Then

                       strFixdate = Mid(txtFixdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 1, 2)
                       strFixdate = "'" & SaveChangeEngYear(strFixdate) & "'"

                    Else
                       strFixdate = "NULL"
                    End If

                   '---------------------------- ��˹��Ѻ���  ------------------------------

                    If txtDueDate.Text <> "__/__/____" Then

                       strDuedate = Mid(txtDueDate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 1, 2)
                       strDuedate = "'" & SaveChangeEngYear(strDuedate) & "'"

                   Else
                       strDuedate = "NULL"
                   End If

                   '------------------ Update �����ŵ��ҧ tmp_fixeqptrn -------------------

                        strSqlCmd = "UPDATE tmp_fixeqptrn SET amt_out = " & ChangFormat(txtSetQty.Text) _
                                               & "," & "price = " & ChangFormat(txtPrice.Text.Trim) _
                                               & "," & "pr_doc = '" & ReplaceQuote(txtPr.Text.Trim) & "'" _
                                               & "," & "sup_name = '" & ReplaceQuote(txtSupp.Text.Trim) & "'" _
                                               & "," & "fix_date = " & strFixdate _
                                               & "," & "fix_by = '" & ReplaceQuote(txtFixnm.Text.Trim) & "'" _
                                               & "," & "due_date = " & strDuedate _
                                               & "," & "issue = '" & ReplaceQuote(txtIssue.Text.Trim) & "'" _
                                               & "," & "fix_issue = '" & ReplaceQuote(txtFxIssue.Text.Trim) & "'" _
                                               & "," & "fix_rmk = '" & ReplaceQuote(txtRmk.Text.Trim) & "'" _
                                               & " WHERE size_id = '" & txtSize.Text.Trim & "'"
                      .Execute(strSqlCmd)

                       '------------------------ ź����������͡��͹ ------------------------

                       strSqlCmd = "DELETE FROM fixeqptrn WHERE fix_id = '" & lblFix_id.Text.Trim & "'"
                       .Execute(strSqlCmd)

                       '----------------------- ������������������� fixeqptrn --------------

                       strSqlCmd = "INSERT INTO fixeqptrn " _
                                     & " SELECT fix_sta " _
                                     & ",fix_id = '" & lblFix_id.Text.ToUpper.Trim & "'" _
                                     & ",[group],eqp_id,size_id,amt_out" _
                                     & ",amt_in,price,fix_date,fix_by,pr_doc" _
                                     & ",issue,fix_issue,sup_name,due_date,recv_date" _
                                     & ",recv_by,fix_rmk" _
                                     & " FROM tmp_fixeqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"

                        .Execute(strSqlCmd)

                   '----------------- �� sum price ��� sum amt -------------------------

                   strSqlSelc = "SELECT SUM(price) as sumPrice , SUM(amt_out) as sumAmt" _
                                                         & " FROM tmp_fixeqptrn (NOLOCK)"

                   With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenDynamic
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlSelc, Conn, , , )

                        If .RecordCount <> 0 Then

                            strSqlCmd = "UPDATE fixeqpmst SET amount = " & Format(.Fields("sumAmt").Value, "#,##0.0") _
                                                  & "," & "price = " & ChangFormat(.Fields("sumPrice").Value) _
                                                  & " WHERE fix_id = '" & lblFix_id.Text.Trim & "'"

                            Conn.Execute(strSqlCmd)
                        End If

                       .ActiveConnection = Nothing
                       .Close()

                   End With

            Conn.Close()
            Conn = Nothing

        End With

        frmFixEqpmnt.lblCmd.Text = lblFix_id.Text.ToString.Trim          '���������§��ѧ�������ѡ
        frmFixEqpmnt.Activating()
        Me.Close()

End Sub

Public Sub UpdateFixsta(ByVal sta As String)            '�Ѿഷ fix_sta ���ҧ eqpmst

 Dim Conn As New ADODB.Connection
 Dim strSqlcmd As String

 Dim DateSave As Date = Now()
 Dim strDate As String = ""
 Dim strEqpid As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 150
             .Open()

     End With

               strEqpid = txtEqp_id.Text.ToUpper.Trim

               strDate = DateSave.Date.ToString("yyyy-MM-dd")
               strDate = SaveChangeEngYear(strDate)

               strSqlcmd = " UPDATE eqpmst SET fix_sta =  '" & sta & "'" _
                                            & "," & "last_date = '" & strDate & "'" _
                                            & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & " WHERE eqp_id ='" & strEqpid & "'"

              Conn.Execute(strSqlcmd)

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub CheckDataBfSave()

 Dim bytConSave As Byte

     If txtEqp_id.Text <> "" Then

              If txtEqpnm.Text <> "" Then

                      If cmbType.Text <> "" Then

                                     bytConSave = MsgBox("�س��ͧ��úѹ�֡���������������!" _
                                                        , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                            If bytConSave = 6 Then   '�� Yes


                                                   Select Case Me.Text

                                                          Case Is = "����������"

                                                                If dgvFixDet.RowCount <> 0 Then
                                                                   SaveNewRecord()

                                                                Else
                                                                     MsgBox("�ô�кآ�������������´��ë��� ��͹�ѹ�֡������", MsgBoxStyle.Critical, "�����͹")
                                                                     dgvFixDet.Focus()

                                                                End If

                                                          Case Is = "��䢢�����"

                                                                  SaveEditRecord()

                                                          Case Else

                                                                  SaveReceiveEqp()

                                                   End Select

                                            End If

                      Else
                           MsgBox("�ô�кآ����Ż������ػ�ó� " _
                                      & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                           cmbType.DroppedDown = True
                           cmbType.Focus()

                      End If

              Else
                   MsgBox("�ô�кآ����������ػ�ó�  " _
                            & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                   txtEqpnm.Focus()
              End If


      Else
           MsgBox("�ô�кآ����������ػ�ó�  " _
                          & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
           txtEqp_id.Focus()
      End If

End Sub

Private Sub SaveReceiveEqp()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlcmd As String
 Dim strSqlUpdate As String
 Dim staFix As String = ""
 Dim staFix_ans As String = ""
 Dim strRecvdate As String

     With Conn

          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .CommandTimeout = 90
             .Open()


                        Conn.BeginTrans()      '�ش������� Transection

                       '---------------------------------------- �ѹ����Ѻ����ػ�ó� ��ѧ�觫���  ---------------------------------

                       If txtRecv_date.Text <> "__/__/____" Then

                          strRecvdate = Mid(txtRecv_date.Text.ToString, 7, 4) & "-" _
                                         & Mid(txtRecv_date.Text.ToString, 4, 2) & "-" _
                                         & Mid(txtRecv_date.Text.ToString, 1, 2)
                          strRecvdate = "'" & SaveChangeEngYear(strRecvdate) & "'"

                       Else
                          strRecvdate = "NULL"
                       End If

                        '----------- ������Ѻ�׹�ú������� ----------------------------------

                        If CSng(txtSetQty.Text) > CSng(txtRecvTotal.Text) Then    '��Ҩӹǹ�觫��� > �ӹǹ�Ѻ�׹  
                           staFix_ans = "3"   '�Ѻ�׹�ҧ��ǹ

                        ElseIf CSng(txtSetQty.Text) = CSng(txtRecvTotal.Text) Then
                               staFix_ans = "2"  '�Ѻ�׹�ú

                        End If

                        '---------------------------------------- ����������㹵��ҧ tmp_fixEqptrn -------------------------

                        strSqlUpdate = "UPDATE tmp_fixeqptrn SET fix_sta = '" & staFix_ans & "'" _
                                          & "," & "amt_in = " & ChangFormat(txtRecvTotal.Text.ToString.Trim) _
                                          & "," & "fix_issue = '" & ReplaceQuote(txtFxIssue.Text.ToString.Trim) & "'" _
                                          & "," & "recv_date = " & strRecvdate _
                                          & "," & "recv_by = '" & ReplaceQuote(txtRecvNm.Text.ToString.Trim) & "'" _
                                          & " WHERE fix_id = '" & lblFix_id.Text.ToString.Trim & "'" _
                                          & " AND size_id = '" & txtSize.Text.ToString.Trim & "'"


                        Conn.Execute(strSqlUpdate)

                       '--------------------------------------- ź������㹵��ҧ fixeqptrn2 --------------------------------

                        strSqlcmd = "Delete FROM fixeqptrn" _
                                            & " WHERE fix_id ='" & lblFix_id.Text.ToString.Trim & "'"

                       .Execute(strSqlcmd)

                          '---------------------------------------- ����������㹵��ҧ fixEqptrn2 -------------------------------

                         strSqlcmd = "INSERT INTO fixeqptrn " _
                                     & " SELECT fix_sta " _
                                     & ",fix_id = '" & lblFix_id.Text.ToUpper.Trim & "'" _
                                     & ",[group],eqp_id,size_id,amt_out" _
                                     & ",amt_in,price,fix_date,fix_by,pr_doc" _
                                     & ",issue,fix_issue,sup_name,due_date,recv_date" _
                                     & ",recv_by,fix_rmk" _
                                     & " FROM tmp_fixeqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"

                         .Execute(strSqlcmd)

                        '-------------- ��Ǩ�ͺ��Ҩӹǹ�觫����Ѻ�Ѻ��� ��ҡѹ������� -----------------

                         If CheckStaReceived() Then
                            staFix = "3"    '�Ѻ�׹�ҧ��ǹ
                         Else
                             staFix = "2"   '�Ѻ�׹�ú�ӹǹ"
                         End If

                        '------------------ �Ѿഷ�������  fixeqpmst --------------------------------

                        strSqlUpdate = "UPDATE fixeqpmst SET fix_sta = '" & staFix & "'" _
                                               & "WHERE fix_id = '" & lblFix_id.Text.ToString.Trim & "'"

                       .Execute(strSqlUpdate)
                       .CommitTrans()

                        staFix = ""
                        staFix_ans = ""

                        '------------ �Ѿഷ������㹵��ҧ eqpmst ��ʴ�ʶҹС�ë��� ------------------

                        If FindData(txtEqp_id.Text.ToUpper.Trim) Then    '���������ػ�ó�㹵��ҧ eqpmst
                           UpdateFixsta(staFix)                          '�Ѿഷʶҹ� = 2  ���¶֧ �觫���
                        End If
                        ClearTmpTable(0, "")  'ź������ Table tmp_eqptrn where user_id..

                   frmFixEqpmnt.lblCmd.Text = lblFix_id.Text.ToString.Trim          '���������§��ѧ�������ѡ
                   frmFixEqpmnt.Activating()
                   Me.Close()

    End With

    Conn.Close()
    Conn = Nothing

End Sub

Private Function CheckStaReceived() As Boolean

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
                            & " WHERE fix_id = '" & lblFix_id.Text.Trim & "'" _
                            & " AND fix_sta = '" & "1" & "'"

       With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

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

Private Sub SaveNewRecord()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmd As String
 Dim dateSave As Date = Now()    '�纤���ѹ���Ѩ�غѹ
 Dim strDate As String
 Dim strDateDoc As String
 Dim strFixdate As String        '�ѹ����觫���
 Dim strDuedate As String        '��˹��Ѻ���

 Dim strGpType As String = ""
 Dim strDateNull As String = "NULL"

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .CommandTimeout = 90
             .Open()
     End With

              Conn.BeginTrans()

              strDate = dateSave.Date.ToString("yyyy-MM-dd")
              strDate = SaveChangeEngYear(strDate)                 '��ŧ�繻� �.�.

              strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                               & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                               & Mid(txtBegin.Text.Trim, 1, 2)
              strDateDoc = SaveChangeEngYear(strDateDoc)

              '----------------------------------- Ǵ�.����觫��� -------------------------------------------

              If txtFixdate.Text <> "__/__/____" Then

                 strFixdate = Mid(txtFixdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 1, 2)
                 strFixdate = "'" & SaveChangeEngYear(strFixdate) & "'"

              Else
                   strFixdate = "NULL"

              End If

              '----------------------------------- ��˹��Ѻ��� -------------------------------------------

              If txtDueDate.Text <> "__/__/____" Then

                 strDuedate = Mid(txtDueDate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 1, 2)
                 strDuedate = "'" & SaveChangeEngYear(strDuedate) & "'"

             Else
                   strDuedate = "NULL"
             End If

             '---------------------------------- �������ػ�ó��觫��� -------------------------------------------

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

                   '----------------------- ����������ػ�ó�������觫������������ѧ   True = �ѧ�������к�  ------------------------------

                   If CheckFixEqpid() = True Then 'And CheckFixtrn() = True    '����ѧ����������觫������

                         '------------------------------- �ѹ�֡������ŧ fixeqpmst2 ------------------------------

                         strSqlCmd = "INSERT INTO fixeqpmst" _
                                    & "(fix_sta,[group],pro_sta,fix_id,eqp_id,eqp_name" _
                                    & ",amount,price,pre_date,pre_by,last_date,last_by,remark" _
                                    & ")" _
                                    & " VALUES (" _
                                    & "'" & "1" & "'" _
                                    & ",'" & strGpType & "'" _
                                    & ",'" & "" & "'" _
                                    & ",'" & ReplaceQuote(lblFix_id.Text.ToString.Trim) & "'" _
                                    & ",'" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                    & ",'" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                    & "," & ChangFormat(lblAmount.Text.ToString.Trim) _
                                    & "," & ChangFormat(lblAmt.Text.ToString.Trim) _
                                    & ",'" & strDate & "'" _
                                    & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                    & "," & strDateNull _
                                    & ",'" & "" & "'" _
                                    & ",'" & "" & "'" _
                                    & ")"

                         Conn.Execute(strSqlCmd)


                        strSqlCmd = "INSERT INTO fixeqptrn " _
                                     & " SELECT fix_sta " _
                                     & ",fix_id = '" & lblFix_id.Text.ToUpper.Trim & "'" _
                                     & ",[group],eqp_id,size_id,amt_out" _
                                     & ",amt_in,price,fix_date,fix_by,pr_doc" _
                                     & ",issue,fix_issue,sup_name,due_date,recv_date" _
                                     & ",recv_by,fix_rmk" _
                                     & " FROM tmp_fixeqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"

                        Conn.Execute(strSqlCmd)
                        Conn.CommitTrans()

                   Else
                       MsgBox("�����ū�� �ô�Դ˹�ҵ�ҧ������Ƿ���¡������!...")

                   End If


                     'lblComplete.Text = lblFix_id.Text.ToString.Trim          '���������§��ѧ�������ѡ
                     Me.Hide()

                           If FindData(txtEqp_id.Text.ToUpper.Trim) Then      '���������ػ�ó�� Table eqpmst
                              UpdateFixsta("1")                               '�Ѿഷʶҹ� �觫���� eqpmst
                           End If

        frmFixEqpmnt.lblCmd.Text = lblFix_id.Text.ToString.Trim          '���������§��ѧ�������ѡ
        frmFixEqpmnt.Activating()
        Me.Close()

    Conn.Close()
    Conn = Nothing

End Sub

Private Function CheckFixEqpid() As Boolean              '�示ѹ�֡�����ػ�ó���

  Dim conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlSelc As String

      With conn

           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 150
              .Open()

      End With

      strSqlSelc = " SELECT * FROM fixeqpmst(NOLOCK)" _
                                    & " WHERE fix_id = '" & lblFix_id.Text.ToString.Trim & "'"

      With Rsd

         .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
         .LockType = ADODB.LockTypeEnum.adLockOptimistic
         .Open(strSqlSelc, conn, , , )

          If .RecordCount <> 0 Then
             Return False

          Else
             Return True

          End If

      .ActiveConnection = Nothing
      .Close()
      End With

 conn.Close()
 conn = Nothing

End Function

Private Sub SaveEditRecord()
 Dim Conn As New ADODB.Connection
 Dim strSqlcmd As String

 Dim DateSave As Date = Now()
 Dim strDate As String = ""
 Dim strDocDate As String
 Dim strFixDate As String
 Dim strDuedate As String
 Dim strEqpid As String

 Dim strGpType As String = ""

     strEqpid = txtEqp_id.Text.ToUpper.Trim

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

              '--------------------------- �ѹ����͡��� ----------------------------------------------------

               strDocDate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocDate = SaveChangeEngYear(strDocDate)

               '---------------------------------------- Ǵ�.����觫��� -----------------------------------

                If txtFixdate.Text <> "__/__/____" Then

                   strFixDate = Mid(txtFixdate.Text.ToString.Trim, 7, 4) & "-" _
                            & Mid(txtFixdate.Text.ToString.Trim, 4, 2) & "-" _
                            & Mid(txtFixdate.Text.ToString.Trim, 1, 2)
                   strFixDate = "'" & SaveChangeEngYear(strFixDate) & "'"

                Else
                   strFixDate = "NULL"

                End If

               '---------------------------------------- ��˹��Ѻ��� ------------------------------------

               If txtDueDate.Text <> "__/__/____" Then

                  strDuedate = Mid(txtDueDate.Text.ToString.Trim, 7, 4) & "-" _
                            & Mid(txtDueDate.Text.ToString.Trim, 4, 2) & "-" _
                            & Mid(txtDueDate.Text.ToString.Trim, 1, 2)
                  strDuedate = "'" & SaveChangeEngYear(strDuedate) & "'"

               Else
                  strDuedate = "NULL"

               End If

                     '---------------------------------------- �������ػ�ó� ----------------------------------

                     Select Case cmbType.SelectedIndex

                            Case Is = 0
                                   strGpType = "A"
                            Case Is = 1
                                   strGpType = "B"
                            Case Is = 2
                                   strGpType = "C"
                            Case Is = 3
                                   strGpType = "D"
                            Case Is = 4
                                   strGpType = "E"
                            Case Is = 5
                                   strGpType = "F"
                            Case Is = 5
                                   strGpType = "G"

                      End Select

                       strSqlcmd = " UPDATE fixeqpmst SET [group]= '" & strGpType & "'" _
                                            & "," & "eqp_id = '" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                            & "," & "eqp_name = '" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                            & "," & "amount = '" & ReplaceQuote(lblAmount.Text.ToString.Trim) & "'" _
                                            & "," & "last_date = '" & strDate & "'" _
                                            & "," & "last_by = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & "," & "pro_sta = '" & "0" & "'" _
                                            & "," & "price = " & ChangFormat(lblAmt.Text.ToString.Trim) _
                                            & " WHERE fix_id = '" & lblFix_id.Text.ToString.Trim & "'"

                       Conn.Execute(strSqlcmd)

                      '--------------------------------------- ź������㹵��ҧ fixeqptrn -------------------------------

                      strSqlcmd = "Delete FROM fixeqptrn" _
                                          & " WHERE fix_id ='" & lblFix_id.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlcmd)

                      '--------------------------------------- ����������㹵��ҧ fixEqptrn -------------------------------

                     strSqlcmd = "INSERT INTO fixeqptrn " _
                                     & " SELECT fix_sta " _
                                     & ",fix_id = '" & lblFix_id.Text.ToUpper.Trim & "'" _
                                     & ",[group],eqp_id,size_id,amt_out" _
                                     & ",amt_in,price,fix_date,fix_by,pr_doc" _
                                     & ",issue,fix_issue,sup_name,due_date,recv_date" _
                                     & ",recv_by,fix_rmk" _
                                     & " FROM tmp_fixeqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"

                     Conn.Execute(strSqlcmd)
                     Conn.CommitTrans()                               '��� Commit transection

                          If CheckHaveData() Then    '����Ѻ����ػ�ó�����
                             SaveReceiveEqp()    '�ѹ�֡��� ����Ѻ����ػ�ó�
                          End If

        frmFixEqpmnt.lblCmd.Text = lblFix_id.Text.ToString.Trim          '���������§��ѧ�������ѡ
        frmFixEqpmnt.Activating()
        Me.Close()

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub txtEqp_id_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqp_id.KeyPress
 If e.KeyChar = Chr(13) Then
    txtEqpnm.Focus()
 End If
End Sub

Private Sub txtEqp_id_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqp_id.LostFocus
  txtEqp_id.Text = txtEqp_id.Text.ToUpper.Trim
End Sub

Private Sub txtFixdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFixdate.GotFocus
  With mskFixdate
       .BringToFront()
       txtFixdate.SendToBack()
       .Focus()
  End With
End Sub

Private Sub txtFixdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFixdate.KeyPress
  If e.KeyChar = Chr(13) Then
     txtFixnm.Focus()
  End If
End Sub

Private Sub mskFixdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFixdate.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerg As String = ""

        With mskFixdate
            If txtFixdate.Text.Trim <> "__/__/____" Then
                x = Len(txtFixdate.Text)

                For i = 1 To x

                    strTmp = Mid(txtFixdate.Text.Trim, i, 1)
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

Private Sub mskFixdate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskFixdate.KeyDown

 Dim intChkPoint As Integer
        With mskFixdate
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtSupp.Focus()
                    End If
                Case Is = 38 '�١�â��
                    txtSetQty.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtFixnm.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtFixnm.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtIssue.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub mskFixdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskFixdate.KeyPress
  If e.KeyChar = Chr(13) Then
     txtFixnm.Focus()
  End If
End Sub

Private Sub mskFixdate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFixdate.LostFocus
 Dim i, x As Integer
 Dim z As Date

 Dim strTmp As String = ""
 Dim strMerge As String = ""

        With mskFixdate
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
                mskFixdate.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)

                If Year(z) < 2500 Then
                    txtFixdate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))

                Else
                    txtFixdate.Text = z.ToString("dd/MM/yyyy")
                End If
            Catch ex As Exception
                mskFixdate.Text = "__/__/____"
                txtFixdate.Text = "__/__/____"

            End Try
            mskFixdate.SendToBack()
            txtFixdate.BringToFront()

        End With
End Sub

Private Sub txtDueDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDueDate.GotFocus
 With mskDueDate
      .BringToFront()
       txtDueDate.SendToBack()
      .Focus()
 End With
End Sub

Private Sub mskDueDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskDueDate.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerg As String = ""

        With mskDueDate

            If txtDueDate.Text.Trim <> "__/__/____" Then
                x = Len(txtDueDate.Text)

                For i = 1 To x

                    strTmp = Mid(txtDueDate.Text.Trim, i, 1)
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

Private Sub mskDueDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskDueDate.KeyDown
 Dim intChkPoint As Integer

        With mskDueDate
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtFixnm.Focus()
                    End If
                Case Is = 38 '�١�â��
                    txtPr.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtIssue.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                           txtIssue.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtIssue.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub mskDueDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskDueDate.KeyPress
  If e.KeyChar = Chr(13) Then
     txtIssue.Focus()
  End If
End Sub

Private Sub mskDueDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskDueDate.LostFocus
 Dim i, x As Integer
 Dim z As Date

 Dim strTmp As String = ""
 Dim strMerge As String = ""

        With mskDueDate
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
                mskDueDate.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)

                If Year(z) < 2500 Then
                    txtDueDate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))
                Else
                    txtDueDate.Text = z.ToString("dd/MM/yyyy")
                End If
            Catch ex As Exception
                mskDueDate.Text = "__/__/____"
                txtDueDate.Text = "__/__/____"

            End Try
            mskDueDate.SendToBack()
            txtDueDate.BringToFront()

        End With
End Sub

Private Function SaveSubRecord() As Boolean

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlCmd As String
  Dim strSqlSelc As String
  Dim dateSave As Date = Now()            '��ʵ�ԧ�ѹ���Ѩ�غѹ
  Dim strDate As String
  Dim strDateDoc As String
  Dim strFixdate As String
  Dim strDuedate As String
  Dim strTypeEqp As String = ""
  Dim strSizeid As String = txtSize.Text.ToString.Trim

      With Conn
            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

      End With

      strSqlSelc = "SELECT size_id FROM tmp_fixeqptrn " _
                                  & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                                  & "AND size_id = '" & strSizeid & "'"

      With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            If .RecordCount <> 0 Then

                 MessageBox.Show("Size :" & txtSize.Text.ToString & _
                                            " ����к����� ��س��к� Size ����", "�����ū��!", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                 SaveSubRecord = False

            Else

                strDate = dateSave.Date.ToString("yyyy-MM-dd")
                strDate = SaveChangeEngYear(strDate)

                strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                      & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                      & Mid(txtBegin.Text.ToString, 1, 2)
                strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"

                '-------------------------- �ѹ���觫��� ------------------------------------------------

                If txtFixdate.Text <> "__/__/____" Then   '�Ѵ �� ��͹ �ѹ
                    strFixdate = Mid(txtFixdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtFixdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtFixdate.Text.ToString, 1, 2)
                    strFixdate = "'" & SaveChangeEngYear(strFixdate) & "'"     '��ŧ����繻� ��.(� module)
                Else
                    strFixdate = "NULL"

                End If

                '-------------------------- ��˹��Ѻ���  ----------------------------------------------

                If txtDueDate.Text <> "__/__/____" Then   '�Ѵ �� ��͹ �ѹ
                    strDuedate = Mid(txtDueDate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtDueDate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtDueDate.Text.ToString, 1, 2)
                    strDuedate = "'" & SaveChangeEngYear(strDuedate) & "'"     '��ŧ����繻� ��.(� module)

                Else
                    strDuedate = "NULL"

                End If

                    Select Case cmbType.SelectedIndex

                        Case Is = 0
                            strTypeEqp = "A"
                        Case Is = 1
                            strTypeEqp = "B"
                        Case Is = 2
                            strTypeEqp = "C"
                        Case Is = 3
                            strTypeEqp = "D"
                        Case Is = 4
                            strTypeEqp = "E"
                        Case Is = 5
                            strTypeEqp = "F"
                        Case Is = 6
                            strTypeEqp = "G"

                    End Select

                    strSqlCmd = "INSERT INTO tmp_fixeqptrn " _
                                     & "(fix_sta,[group],eqp_id,size_id,fix_amount" _
                                     & ",issue,fix_issue,sup_name,fix_price,pr_doc,fix_date" _
                                     & ",fix_by,due_date,recv_date,recv_by,fix_rmk,user_id,fix_id" _
                                     & ")" _
                                     & " VALUES (" _
                                     & "'" & "1" & "'" _
                                     & ",'" & strTypeEqp & "'" _
                                     & ",'" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                     & ",'" & ChangFormat(txtSetQty.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtIssue.Text.ToString.Trim) & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & ReplaceQuote(txtSupp.Text.ToString.Trim) & "'" _
                                     & ",'" & "0.00" & "'" _
                                     & ",'" & ReplaceQuote(txtPr.Text.ToString.Trim) & "'" _
                                     & "," & strFixdate _
                                     & ",'" & ReplaceQuote(txtFixnm.Text.ToString.Trim) & "'" _
                                     & "," & strDuedate _
                                     & "," & "Null" _
                                     & ",'" & "" & "'" _
                                     & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                     & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                     & ",'" & ReplaceQuote(lblFix_id.Text.Trim.ToString) & "'" _
                                     & ")"

               Conn.Execute(strSqlCmd)
               SaveSubRecord = True

           End If

      .ActiveConnection = Nothing
      .Close()
      End With
      Rsd = Nothing

Conn.Close()
Conn = Nothing

End Function

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown

 Dim intChkPoint As Integer

        With txtSize

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                Case Is = 38 '�����١�â��        
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSetQty.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSetQty.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtSupp.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress

    Select Case Asc(e.KeyChar)

           Case 48 To 122 ' �������ѧ��������ԧ�������� 58�֧122 ������� 48��������ҵ�ͧ��õ���Ţ
                  e.Handled = False
           Case 8, 46 ' Backspace = 8,  Delete = 46
                  e.Handled = False
           Case 13     'Enter = 13
                  e.Handled = False
                  txtSetQty.Focus()
          Case 161 To 240 ' ���������ç����繤����������������駵�����+��ó�ء����¹�Ф�Ѻ
                  e.Handled = True
                  MsgBox("��س��кآ������������ѧ��� ���͵���Ţ��ҹ��", MsgBoxStyle.Critical, "�Դ��Ҵ")
          Case Else
                  e.Handled = False

    End Select

End Sub

Private Sub txtSize_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSize.LostFocus
  txtSize.Text = txtSize.Text.ToUpper.Trim
End Sub

Private Sub txtSetQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSetQty.KeyDown
 Dim intChkPoint As Integer

        With txtSetQty

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                       If .SelectionStart = 0 Then
                          txtSize.Focus()
                       End If
                Case Is = 38 '�����١�â��        
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtPrice.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtPrice.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtFixdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtPr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPr.KeyDown
 Dim intChkPoint As Integer
        With txtPr

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                       If .SelectionStart = 0 Then
                          txtPrice.Focus()
                       End If
                Case Is = 38 '�����١�â��        
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSupp.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtSupp.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtFixnm.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtPr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPr.KeyPress

    Select Case Asc(e.KeyChar)

           Case 48 To 122 ' �������ѧ��������ԧ�������� 58�֧122 ������� 48��������ҵ�ͧ��õ���Ţ
                  e.Handled = False
           Case 8, 46 ' Backspace = 8,  Delete = 46
                  e.Handled = False
           Case 13     'Enter = 13
                  e.Handled = False
                 txtSupp.Focus()
           Case 161 To 240 ' ���������ç����繤����������������駵�����+��ó�ء����¹�Ф�Ѻ
                  e.Handled = True
                  MsgBox("��س��кآ������������ѧ��� ���͵���Ţ��ҹ��", MsgBoxStyle.Critical, "�Դ��Ҵ")
           Case Else
                  e.Handled = False

    End Select

End Sub

Private Sub txtPr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPr.LostFocus
  txtPr.Text = txtPr.Text.ToUpper.Trim
End Sub

Private Sub txtSupp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupp.KeyDown

  Dim intChkPoint As Integer
        With txtSupp

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                       If .SelectionStart = 0 Then
                          txtPr.Focus()
                       End If
                Case Is = 38 '�����١�â��        
                      txtSize.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtFixdate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtFixdate.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtDueDate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtSupp_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSupp.KeyPress
 If e.KeyChar = Chr(13) Then
    txtFixdate.Focus()
 End If
End Sub

Private Sub txtSupp_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSupp.LostFocus
 txtSupp.Text = txtSupp.Text.ToUpper.Trim
 txtSupp.SelectionStart = 0       'point ��ѧ���˹��������
End Sub

Private Sub txtFixnm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFixnm.KeyDown

 Dim intChkPoint As Integer

        With txtFixnm

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                       If .SelectionStart = 0 Then
                          txtFixdate.Focus()
                       End If
                Case Is = 38 '�����١�â��        
                       txtPr.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtDueDate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtDueDate.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtIssue.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtFixnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFixnm.KeyPress
 If e.KeyChar = Chr(13) Then
    txtDueDate.Focus()
 End If
End Sub

Private Sub txtFixnm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFixnm.LostFocus
 txtFixnm.Text = txtFixnm.Text.ToUpper.Trim
End Sub

Private Sub lstIssue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

      With txtIssue

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                          txtDueDate.Focus()

                Case Is = 38 '�����١�â��     
                        txtFixdate.Focus()
                Case Is = 39 '�����١�â��
                        txtRmk.Focus()

                Case Is = 40 '����ŧ
                        txtRmk.Focus()
                Case Is = 113 '���� F2

            End Select
        End With
End Sub

Private Sub txtRmk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRmk.KeyDown

    Dim intChkPoint As Integer

        With txtRmk

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                     If .SelectionStart = 0 Then
                        txtIssue.Focus()
                     End If
                Case Is = 38 '�����١�â��     
                        txtDueDate.Focus()
                Case Is = 39 '�����١�â��
                     If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtRecv_date.Focus()
                     Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                         If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtRecv_date.Focus()
                         End If
                     End If

                Case Is = 40 '����ŧ
                     txtRecv_date.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With

End Sub

Private Sub txtDueDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDueDate.KeyPress
  If e.KeyChar = Chr(13) Then
     txtIssue.Focus()
  End If
End Sub

Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
Dim strSearch As String

   strSearch = txtSeek.Text.Trim.ToUpper
    If Len(strSearch) <> 0 Then

        FindPsData(strSearch)
        btnSearch.Focus()

    Else
        MessageBox.Show("��سҡ�͡�����ػ�ó����ͤ���", "����բ���������Ѻ����", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        txtEqp_id.Focus()
    End If
End Sub

Private Sub txtSeek_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSeek.KeyPress
  Dim strSearch As String
  If e.KeyChar = Chr(13) Then

     strSearch = txtSeek.Text.Trim.ToUpper
     If Len(strSearch) <> 0 Then

        FindPsData(strSearch)
        btnSearch.Focus()

      End If
 End If
End Sub

Private Sub txtSeek_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSeek.LostFocus
  txtSeek.Text = txtSeek.Text.ToUpper.Trim
End Sub

Private Sub CallEditData()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String
 Dim strSize As String = frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(4).Value.ToString

     With Conn

         If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

     End With

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
                 txtSetQty.Text = .Fields("amt_out").Value.ToString.Trim
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

              End If
              .ActiveConnection = Nothing
              .Close()

      End With
      Rsd = Nothing

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub txtSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSetQty.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPrice.Focus()
   End If
End Sub

Private Sub txtSetQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSetQty.GotFocus
    With mskSetQty
         .BringToFront()
         txtSetQty.SendToBack()
         .Focus()
   End With
End Sub

Private Sub mskSetQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSetQty.GotFocus

 Dim i, x As Integer
 Dim strTmp As String = ""
 Dim strMerg As String = ""

     With mskSetQty

            If txtSetQty.Text.ToString.Trim <> "" Then
                x = Len(txtSetQty.Text.ToString)

                For i = 1 To x
                    strTmp = Mid(txtSetQty.Text.ToString, i, 1)
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

Private Sub mskSetQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSetQty.KeyDown

    Dim intChkpoint As Integer

     With mskSetQty

            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtSize.Focus()
                    End If
                Case Is = 38 '�����١�â��  

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtPrice.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkpoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtPrice.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtFixdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub mskSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSetQty.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPrice.Focus()
   End If
End Sub

Private Sub mskSetQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSetQty.LostFocus

  Dim i, x, intFull As Integer
  Dim z As Double

  Dim strTmp As String = ""
  Dim strMerg As String = ""

      With mskSetQty

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

                mskSetQty.Text = ""            '������ mskSizeQty
                z = CDbl(strMerg)              '�ŧ Type dbl
                intFull = Int(z)

                If (z - intFull) > 0 Then
                    txtSetQty.Text = z.ToString("#,##0.0")

                Else
                   txtSetQty.Text = z.ToString("0")

                End If

            Catch ex As Exception
                txtSetQty.Text = "0"
                mskSetQty.Text = ""
            End Try

            mskSetQty.SendToBack()
            txtSetQty.BringToFront()

        End With

End Sub

Private Sub txtPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrice.GotFocus
  With mskPrice
       .BringToFront()
       txtPrice.SendToBack()
       .Focus()
  End With
End Sub

Private Sub mskPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrice.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerge As String = ""

     With mskPrice

          If txtPrice.Text <> "0.00" Then

                        x = Len(txtPrice.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtPrice.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = "_"
                                       Case Else

                                            If InStr(",.0123456789", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If

                                End Select

                         Next i

                        Select Case strMerge.IndexOf(".")   '�ҵ��˹觷�辺�繤����á

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

Private Sub mskPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskPrice.KeyDown
  Dim intChkPoint As Integer

    With mskPrice

            Select Case e.KeyCode

                   Case Is = 35 '���� End 
                   Case Is = 36 '���� Home
                   Case Is = 37 '�١�ë���

                       If .SelectionStart = 0 Then
                          txtSetQty.Focus()
                       End If

                   Case Is = 38 '�����١�â��    
                   Case Is = 39 '�����١�â��

                         If .SelectionLength = .Text.Trim.Length Then
                             txtPr.Focus()
                         Else
                             intChkPoint = .Text.Trim.Length
                             If .SelectionStart = intChkPoint Then
                                txtPr.Focus()
                             End If

                         End If

                   Case Is = 40 '����ŧ    
                           txtFixdate.Focus()
                   Case Is = 113 '���� F2
                           .SelectionStart = .Text.Trim.Length

            End Select
    End With

End Sub

Private Sub mskPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrice.LostFocus

 Dim i, x As Integer
 Dim z As Double

 Dim strTmp As String = ""
 Dim strMerge As String = ""

     With mskPrice

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

                    mskPrice.Text = ""
                    z = CDbl(strMerge)
                    txtPrice.Text = z.ToString("#,##0.00")


                Catch ex As Exception
                    txtPrice.Text = "0.00"
                    mskPrice.Text = ""
               End Try

        mskPrice.SendToBack()
        txtPrice.BringToFront()

   End With

End Sub

Private Sub LoadDataDetail()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strCmd As String                                   ' ��ʵ�ԧ Command

 Dim blnHavedata As Boolean                             '�纤�ҵ����� ����Ѻ������բ������������
 Dim strSqlSelc As String = ""                          '��ʵ�ԧ sql select
 Dim strPart As String = ""

 Dim strFxID As String = frmFixEqpmnt.dgvShow.Rows(frmFixEqpmnt.dgvShow.CurrentRow.Index).Cells(2).Value.ToString

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

     strSqlSelc = "SELECT * FROM fixeqpmst (NOLOCK) " _
                                   & " WHERE fix_id = '" & strFxID & "'"

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then
             txtEqp_id.Text = .Fields("eqp_id").Value.ToString.Trim
             lblFix_id.Text = .Fields("fix_id").Value.ToString.Trim
             txtEqpnm.Text = .Fields("eqp_name").Value.ToString.Trim
             lblAmount.Text = .Fields("amount").Value.ToString.Trim
             lblAmt.Text = .Fields("price").Value.ToString.Trim

                Select Case .Fields("group").Value.ToString.Trim

                       Case Is = "A"
                            cmbType.Text = "���촩մ EVA INJECTION"

                       Case Is = "B"
                            cmbType.Text = "���촩մ PVC INJECTION"

                       Case Is = "C"
                            cmbType.Text = "������ʹ PU"

                       Case Is = "D"
                            cmbType.Text = "����ἧ�Ѵ���˹ѧ˹��,���"

                       Case Is = "E"
                            cmbType.Text = "�մ�Ѵ"

                       Case Is = "F"
                            cmbType.Text = "���͡ʡ�չ"

                       Case Else
                            cmbType.Text = "���͡����"

                End Select

                strCmd = frmFixEqpmnt.lblCmd.Text.ToString.Trim    '��� strCmd ��ҡѺ���� lblcmd 㹿���� frmEqpSheet

                Select Case strCmd
                       Case Is = "1"   '�����ͤ�͹���
                       Case Is = "2"   '�����ͤ�͹����ͧ
                            btnSaveData.Enabled = False  '�Դ���� "�ѹ�֡������"
                End Select

              '------------------------------- �ѹ�֡�����ŧ㹵��ҧ tmp_eqptrn ----------------------------

             strSqlSelc = "INSERT INTO tmp_fixeqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "', *" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE fix_id = '" & strFxID & "' "

              Conn.Execute(strSqlSelc)
              blnHavedata = True                     '�觺͡����բ�����
              StateLockFindDept(False)               'Disable groupBox Head

         Else
              blnHavedata = False
         End If

         .ActiveConnection = Nothing                  '��� ReccordSet
         .Close()

     End With

     Rsd = Nothing
     Conn.Close()
     Conn = Nothing

End Sub

Private Sub txtPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPrice.KeyDown

    Dim intChkPoint As Integer

          With txtPrice

               Select Case e.KeyCode

                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���

                               If .SelectionStart = 0 Then
                                    txtSetQty.Focus()
                               End If

                      Case Is = 38 '�����١�â��    
                      Case Is = 39 '�����١�â��

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtPr.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                            txtPr.Focus()
                                    End If

                                End If

                      Case Is = 40 '����ŧ    
                                  txtFixdate.Focus()
                      Case Is = 113 '���� F2
                                .SelectionStart = .Text.Trim.Length
              End Select

     End With

End Sub

Private Sub txtPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPr.Focus()
   End If
End Sub

Private Sub dgvShow_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShow.CellClick
  Dim strType As String

   If dgvShow.RowCount > 0 Then

      With dgvShow

           Select Case .CurrentCell.ColumnIndex          '���͡�������

                  Case Is = 2

                       txtEqp_id.Text = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
                       txtEqpnm.Text = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim
                       strType = .Rows(.CurrentRow.Index).Cells(3).Value.ToString
                       'txtRemark.Text = .Rows(.CurrentRow.Index).Cells(4).Value.ToString

                       If strType = "A" Then
                          cmbType.SelectedIndex = 0

                       ElseIf strType = "B" Then
                          cmbType.SelectedIndex = 1

                       ElseIf strType = "C" Then
                          cmbType.SelectedIndex = 2

                       ElseIf strType = "D" Then
                          cmbType.SelectedIndex = 3

                       ElseIf strType = "E" Then
                          cmbType.SelectedIndex = 4

                       ElseIf strType = "F" Then
                          cmbType.SelectedIndex = 5

                       Else
                         cmbType.SelectedIndex = 6

                       End If
           End Select
           StateLockFindDept(True)
           gpbSeekEqp.Visible = False

      End With

   End If

End Sub

Private Sub btnSearchExit_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchExit.Click
  StateLockFindDept(True)
  gpbSeekEqp.Visible = False
  IsShowSearch = False
End Sub

Private Sub LoadDataReceive()            '��Ŵ��������¡���Ѻ����觫���
 Dim conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String = ""

 Dim strFxID As String = frmFixEqpmnt.dgvShow.Rows(frmFixEqpmnt.dgvShow.CurrentRow.Index).Cells(2).Value.ToString

     With conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

             strSqlSelc = "SELECT * " _
                                   & "FROM tmp_fixeqptrn (NOLOCK) " _
                                   & " WHERE fix_id = '" & strFxID & "'"

     Rsd = New ADODB.Recordset

     With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelc, conn, , , )

             If .RecordCount <> 0 Then

                txtSize.Text = .Fields("size_id").Value.ToString.Trim
                txtSetQty.Text = .Fields("amt_out").Value.ToString.Trim
                txtPrice.Text = Format(.Fields("price").Value, "#,##0.00")

                txtPr.Text = .Fields("pr_doc").Value.ToString.Trim
                txtSupp.Text = .Fields("sup_name").Value.ToString.Trim
                txtFixnm.Text = .Fields("fix_by").Value.ToString.Trim

                'lblSumFx.Text = Format(txtSetQty.Text.Trim, "##.0")

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

                txtIssue.Text = .Fields("issue").Value.ToString.ToString.Trim
             End If

     .ActiveConnection = Nothing
     .Close()
     End With

  conn.Close()
  conn = Nothing
End Sub

Private Sub txtRecv_date_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRecv_date.GotFocus
  With mskRecv_date
       .BringToFront()
       txtRecv_date.SendToBack()
       .Focus()
  End With

End Sub

Private Sub mskRecv_date_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecv_date.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerg As String = ""

        With mskRecv_date

            If txtRecv_date.Text.Trim <> "__/__/____" Then
                x = Len(txtRecv_date.Text)

                For i = 1 To x

                    strTmp = Mid(txtRecv_date.Text.Trim, i, 1)
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

Private Sub mskRecv_date_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskRecv_date.KeyDown
 Dim intChkPoint As Integer

        With mskRecv_date
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                Case Is = 38 '�١�â��
                    txtSetQty.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtRecvNm.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtRecvNm.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtFxIssue.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub mskRecv_date_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskRecv_date.KeyPress
   If e.KeyChar = Chr(13) Then
      txtRecvNm.Focus()
   End If
End Sub

Private Sub mskRecv_date_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecv_date.LostFocus
 Dim i, x As Integer
 Dim z As Date

 Dim strTmp As String = ""
 Dim strMerge As String = ""

        With mskRecv_date
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
                mskRecv_date.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)

                If Year(z) < 2500 Then
                    txtRecv_date.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))

                Else
                    txtRecv_date.Text = z.ToString("dd/MM/yyyy")
                End If
            Catch ex As Exception
                mskRecv_date.Text = "__/__/____"
                txtRecv_date.Text = "__/__/____"

            End Try
            mskRecv_date.SendToBack()
            txtRecv_date.BringToFront()

        End With
End Sub

Private Sub txtRecvNm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRecvNm.KeyDown

    Dim intChkPoint As Integer

          With txtRecvNm

               Select Case e.KeyCode

                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���

                               If .SelectionStart = 0 Then
                                    txtRecv_date.Focus()
                               End If

                      Case Is = 38 '�����١�â��    
                      Case Is = 39 '�����١�â��

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtRecvTotal.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtRecvTotal.Focus()
                                    End If

                                End If

                      Case Is = 40 '����ŧ    
                                 txtFxIssue.Focus()
                      Case Is = 113 '���� F2
                                .SelectionStart = .Text.Trim.Length
              End Select

     End With
End Sub

Private Sub txtRecvNm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRecvNm.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRecvTotal.Focus()
  End If
End Sub

Private Sub txtRecvTotal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRecvTotal.GotFocus
  With mskRecvTotal
       .BringToFront()
       txtRecvTotal.SendToBack()
       .Focus()
  End With
End Sub

Private Sub mskRecvTotal_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvTotal.GotFocus

  Dim i, x As Byte
  Dim strTmp As String = ""
  Dim strMerg As String = ""

      With mskRecvTotal

            If txtRecvTotal.Text.ToString.Trim <> "" Then
                x = Len(txtRecvTotal.Text.ToString)

                For i = 1 To x
                    strTmp = Mid(txtRecvTotal.Text.ToString, i, 1)
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

Private Sub mskRecvTotal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskRecvTotal.KeyDown
 Dim intChkPoint As Integer

     With mskRecv_date

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                       txtRecvNm.Focus()
                Case Is = 38 '�١�â��
                Case Is = 39   '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                       txtFxIssue.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                           txtFxIssue.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtFxIssue.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub mskRecvTotal_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskRecvTotal.KeyPress
  If e.KeyChar = Chr(13) Then
     txtFxIssue.Focus()
  End If
End Sub

Private Sub mskRecvTotal_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvTotal.LostFocus
  Dim i, x, intFull As Integer
  Dim z As Single

  Dim strTmp As String = ""
  Dim strMerg As String = ""
  Dim RecvAmt As Single

      With mskRecvTotal

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

                mskRecvTotal.Text = ""            '������ mskSizeQty
                z = CDbl(strMerg)              '�ŧ Type dbl
                intFull = CInt(z)

                If (z - intFull) > 0 Then
                    txtRecvTotal.Text = z.ToString("##0.0")
                Else
                    txtRecvTotal.Text = z.ToString("0.0")
                End If

            Catch ex As Exception
                txtRecvTotal.Text = "0.0"
                mskRecvTotal.Text = ""
            End Try

            mskRecvTotal.SendToBack()
            txtRecvTotal.BringToFront()

        End With

            '----------------------- �礨ӹǹ�Ѻ����觫��� ----------------------------

            If CSng(txtRecvTotal.Text) > CSng(txtSetQty.Text) Then
               MsgBox("�س�кبӹǹ�Ѻ������١��ͧ!...", MsgBoxStyle.Critical, "�Դ��Ҵ")
               txtRecvTotal.Text = "0.0"
               txtRecvTotal.Focus()

            ElseIf CSng(txtRecvTotal.Text) < CSng(txtSetQty.Text) Then
                   RecvAmt = CSng(txtSetQty.Text) - CSng(txtRecvTotal.Text)
                  ' MsgBox("��ҧ�Ѻ�׹ = " & intRecvAmt & " (SET)", MsgBoxStyle.Critical, "�����͹")
                   lblRemain.Text = Format(RecvAmt, "##0.0")

            Else
                   RecvAmt = CSng(txtSetQty.Text) - CSng(txtRecvTotal.Text)
                   lblRemain.Text = Format(RecvAmt, "##0.0")

            End If

End Sub

Private Sub txtFxIssue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFxIssue.KeyDown

   With txtFxIssue

        Select Case e.KeyCode
               Case Is = 35 '���� End 
               Case Is = 36 '���� Home
               Case Is = 37 '�١�ë���
                    txtRecvTotal.Focus()
               Case Is = 38 '�����١�â��    
               Case Is = 39 '�����١�â��
                    btnSaveData.Focus()
               Case Is = 40 '����ŧ    
               Case Is = 113 '���� F2
       End Select

   End With

End Sub

Private Sub mskPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrice.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPr.Focus()
   End If
End Sub

Private Sub loadFixdata()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String = ""

 Dim imgStaFix As Image
 Dim staFx As String
 Dim CountAmount As Byte
 Dim CountPrice As Single

 Dim strFxID As String
 Dim strSize As String
     strFxID = frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(16).Value.ToString
     strSize = Mid(frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(3).Value.ToString, 2)

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

     strSqlSelc = "SELECT * FROM tmp_fixeqptrn(NOLOCK) " _
                                 & " WHERE fix_id = '" & strFxID & "'" _
                                 & " AND size_id = '" & strSize & "'"

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

             dgvFixDet.Rows.Clear()
             Do While Not .EOF

               '--------------------------- ʶҹ��觫��� ----------------------

               Select Case .Fields("fix_sta").Value.ToString.Trim

                      Case Is = "1"     '�觫���
                           imgStaFix = My.Resources._16x16_ledred
                           staFx = "�觫���"

                      Case Is = "2"     '�Ѻ�׹�觫���
                           imgStaFix = My.Resources._16x16_ledgreen
                           staFx = "�Ѻ�׹�觫���"

                      Case Is = "3"     '�Ѻ�׹�ҧ��ǹ
                           imgStaFix = My.Resources._16x16ledyellow
                           staFx = "�Ѻ�׹�ҧ��ǹ"

                     Case Else          '����
                           imgStaFix = My.Resources.blank
                           staFx = "����"

               End Select

                 dgvFixDet.Rows.Add( _
                                      imgStaFix, _
                                      staFx, _
                                      .Fields("eqp_id").Value.ToString.Trim, _
                                      .Fields("size_id").Value.ToString.Trim, _
                                      .Fields("issue").Value.ToString.Trim, _
                                      Format(.Fields("amt_out").Value, "#0.0"), _
                                      Format(.Fields("amt_in").Value, "#0.0"), _
                                      Format(.Fields("price").Value, "#,###,##0.00"), _
                                      Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("fix_by").Value.ToString.Trim, _
                                      Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("recv_by").Value.ToString.Trim _
                                  )

                        CountAmount = CountAmount + Format(.Fields("amt_out").Value, "#,##0")
                        CountPrice = CountPrice + Format(.Fields("price").Value, "#,###,##0.00")

                .MoveNext()
             Loop

             lblAmount.Text = CountAmount.ToString("#,##0.0")
             lblAmt.Text = CountPrice.ToString("#,###,##0.00")

         End If

        .ActiveConnection = Nothing
        .Close()

     End With

  Rsd = Nothing
  Conn.Close()
  Conn = Nothing

End Sub

Private Sub dgvFixDet_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFixDet.RowsAdded
   dgvFixDet.Rows(e.RowIndex).Height = 28
End Sub

Private Sub lstFxIssue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
  If e.KeyChar = Chr(13) Then
     btnSaveData.Focus()
  End If
End Sub

Private Sub lstIssue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
   If e.KeyChar = Chr(13) Then
      txtRmk.Focus()
   End If
End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

 If lblFix_id.Text <> "" Then

     If txtEqp_id.Text <> "" Then

        If txtEqpnm.Text <> "" Then
           gpbHead.Enabled = False
           ShowGroupAdd()
           'ClearGpbFxDetail()   '��ҧ������� GROUPBOX ��������´����觫���
           staAction = "0"      '�觺͡����繡������������

        Else
            MsgBox("�ô�кت����ػ�ó�", MsgBoxStyle.Critical, "�����͹")
            txtEqpnm.Focus()

        End If

     Else
          MsgBox("�ô���͡�ػ�ó����觫���", MsgBoxStyle.Critical, "�����͹")
          txtEqp_id.Focus()

     End If

  Else
        GenFixID()

  End If

End Sub

Private Sub dgvShow_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvShow.RowsAdded
    dgvShow.Rows(e.RowIndex).Height = 28
End Sub

Private Sub ClearGpbFxDetail()

   txtSize.Text = ""
   txtSetQty.Text = "0"
   txtPrice.Text = "0.00"
   txtPr.Text = ""
   txtSupp.Text = ""
   txtFixdate.Text = "__/__/____"
   txtDueDate.Text = "__/__/____"
   txtFixnm.Text = ""
   txtIssue.Text = ""
   txtRmk.Text = ""

End Sub

Sub ShowGroupAdd()

    With gpbFxDetail

         .Visible = True
         .Top = 209
         .Left = 12
         .Width = 999
         .Height = 511
         .Text = "������������´�觫���"

         gpbHead.Enabled = False
         btnSaveData.Enabled = False
         txtSize.Enabled = True
         txtSize.Focus()
    End With

End Sub

Private Sub btnCancle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancle.Click
    gpbHead.Enabled = True
    gpbFxDetail.Visible = False
    ClearGpbFxDetail()
    btnSaveData.Enabled = True
End Sub

Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

   If txtSize.Text <> "" Then

      If txtSetQty.Text <> "0" Then

             Select Case staAction

                    Case Is = "0"          '������������´����觫���

                         If CheckSizeExist() Then   '��Ǩ�ͺ size ��ӡѹ
                            SaveSubFixEqp()       '�ѹ�֡��¡���觫���

                         Else
                             MsgBox("SIZE : " & txtSize.Text.Trim & " ����к������ �ô��˹� SIZE ���", MsgBoxStyle.Critical, "�����ū��")
                             txtSize.Focus()
                         End If

                    Case Is = "1"          '�����¡���觫���

                         SaveEditsubFxEqp()    '�����¡���觫���

             End Select
             btnSaveData.Enabled = True

      Else
         MsgBox("�ô�кبӹǹ�觫���!", MsgBoxStyle.Critical, "�����͹")
         txtSetQty.Focus()

      End If

   Else
      MsgBox("�ô�к� SIZE !", MsgBoxStyle.Critical, "�����͹")
      txtSize.Focus()

   End If

End Sub

Private Function CheckSizeExist() As Boolean

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

       strSqlSelc = "SELECT * FROM tmp_fixeqptrn (NOLOCK)" _
                              & " WHERE size_id = '" & txtSize.Text.ToUpper.Trim & "'"

       Rsd = New ADODB.Recordset

       With Rsd

            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .CursorType = CursorTypeEnum.adOpenKeyset
            .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then
                 Return False

             Else
                  Return True
             End If

           .ActiveConnection = Nothing
           .Close()

       End With

  Conn.Close()
  Conn = Nothing

End Function

Private Sub SaveSubFixEqp()      '�ѹ�֡��¡���觫���

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim dateSave As Date = Now()    '�纤���ѹ���Ѩ�غѹ
 Dim strFixdate As String        '�ѹ����觫���
 Dim strDuedate As String        '��˹��Ѻ���

 Dim strGpType As String = ""

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .CommandTimeout = 90
             .Open()

                    Conn.BeginTrans()

                    '---------------------------------------- Ǵ�.����觫��� -------------------------------------------

                    If txtFixdate.Text <> "__/__/____" Then

                       strFixdate = Mid(txtFixdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 1, 2)
                       strFixdate = "'" & SaveChangeEngYear(strFixdate) & "'"

                    Else
                       strFixdate = "NULL"
                    End If

                   '---------------------------------------- ��˹��Ѻ��� ----------------------------------------------

                    If txtDueDate.Text <> "__/__/____" Then

                       strDuedate = Mid(txtDueDate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 1, 2)
                       strDuedate = "'" & SaveChangeEngYear(strDuedate) & "'"

                   Else
                       strDuedate = "NULL"
                   End If

                   '-------------------------- �������ػ�ó��觫��� ------------------------------------

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

                        '------------------- �ѹ�֢�����ŧ tmp_fixeqptrn ---------------------

                        strSqlCmd = "INSERT INTO tmp_fixeqptrn" _
                                     & "(user_id,fix_sta,fix_id,[group],eqp_id" _
                                     & ",size_id,amt_out,amt_in,price,fix_date" _
                                     & ",fix_by,pr_doc,issue,fix_issue,sup_name" _
                                     & ",due_date,recv_date,recv_by,fix_rmk" _
                                     & ")" _
                                     & " VALUES (" _
                                     & "'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                     & ",'" & "1" & "'" _
                                     & ",'" & ReplaceQuote(lblFix_id.Text.ToString.Trim) & "'" _
                                     & ",'" & strGpType & "'" _
                                     & ",'" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                     & "," & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                     & "," & 0 _
                                     & "," & ChangFormat(txtPrice.Text.ToString.Trim) _
                                     & "," & strFixdate _
                                     & ",'" & ReplaceQuote(txtFixnm.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtPr.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtIssue.Text.ToString.Trim) & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & ReplaceQuote(txtSupp.Text.ToString.Trim) & "'" _
                                     & "," & strDuedate _
                                     & "," & "NULL" _
                                     & ",'" & "" & "'" _
                                     & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                     & ")"

                        Conn.Execute(strSqlCmd)
                        Conn.CommitTrans()

                        gpbFxDetail.Visible = False
                        ShowDetailFx()    '��ʴ���¡���觫���� DataGridview
                        staAction = ""

            Conn.Close()
            Conn = Nothing

      End With

End Sub

Private Sub ShowDataAFUpDel()  '��ʴ���������ѧ update /Delete

   Dim Conn As New ADODB.Connection
   Dim Rsd As New ADODB.Recordset
   Dim strSqlSelc As String

   Dim strFxID As String
   Dim strSize As String
       strFxID = frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(16).Value.ToString
       strSize = Mid(frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(3).Value.ToString, 2)

   Dim imgStaFix As Image
   Dim staFx As String

   Dim CountFix As Byte = 0
   Dim CountPrice As Single

       With Conn
            If .State Then .Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()
       End With

       strSqlSelc = "SELECT * FROM tmp_fixeqptrn (NOLOCK)" _
                            & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                            & " AND fix_id = '" & strFxID & "'" _
                            & " AND size_id = '" & strSize & "'" _
                            & " ORDER BY fix_id"

       Rsd = New ADODB.Recordset

      With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(strSqlSelc, Conn, , , )

           If .RecordCount <> 0 Then

             dgvFixDet.Rows.Clear()
             Do While Not .EOF

               '--------------------------- ʶҹ��觫��� ----------------------

               Select Case .Fields("fix_sta").Value.ToString.Trim

                      Case Is = "1"     '�觫���
                           imgStaFix = My.Resources._16x16_ledred
                           staFx = "�觫���"

                      Case Is = "2"     '�Ѻ�׹�觫���
                           imgStaFix = My.Resources._16x16_ledgreen
                           staFx = "�Ѻ�׹�觫���"

                      Case Is = "3"     '�Ѻ�׹�ҧ��ǹ
                           imgStaFix = My.Resources._16x16ledyellow
                           staFx = "�Ѻ�׹�ҧ��ǹ"

                     Case Else         '����
                           imgStaFix = My.Resources.blank
                           staFx = "����"

               End Select

                 dgvFixDet.Rows.Add( _
                                      imgStaFix, _
                                      staFx, _
                                      .Fields("eqp_id").Value.ToString.Trim, _
                                      .Fields("size_id").Value.ToString.Trim, _
                                      .Fields("issue").Value.ToString.Trim, _
                                      Format(.Fields("amt_out").Value, "#0.0"), _
                                      Format(.Fields("amt_in").Value, "#0.0"), _
                                      Format(.Fields("price").Value, "#,###,##0.00"), _
                                      Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("fix_by").Value.ToString.Trim, _
                                      Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("recv_by").Value.ToString.Trim _
                                  )

                               CountFix = CountFix + Format(.Fields("amt_out").Value, "#,##0")
                               CountPrice = CountPrice + Format(.Fields("price").Value, "#,###,##0.00")
                .MoveNext()
             Loop

             lblAmount.Text = CStr(CountFix)
             lblAmt.Text = CStr(CountPrice)

           End If

         .ActiveConnection = Nothing
         .Close()

      End With

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub ShowDetailFx()

   Dim Conn As New ADODB.Connection
   Dim Rsd As New ADODB.Recordset
   Dim strSqlSelc As String

   'Dim strFxID As String
   'Dim strSize As String
   '    strFxID = frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(16).Value.ToString
   '    strSize = Mid(frmFixEqpmnt.dgvFix.Rows(frmFixEqpmnt.dgvFix.CurrentRow.Index).Cells(3).Value.ToString, 2)

   Dim imgStaFix As Image
   Dim staFx As String

   Dim CountFix As Byte = 0
   Dim CountPrice As Single

       With Conn
            If .State Then .Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()
       End With

       strSqlSelc = "SELECT * FROM tmp_fixeqptrn (NOLOCK)" _
                            & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                            & " ORDER BY fix_id"
                            '& " AND fix_id = '" & strFxID & "'" _
                            '& " AND size_id = '" & strSize & "'" _


      Rsd = New ADODB.Recordset

      With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(strSqlSelc, Conn, , , )

           If .RecordCount <> 0 Then

             dgvFixDet.Rows.Clear()

             Do While Not .EOF

               '--------------------------- ʶҹ��觫��� ----------------------

               Select Case .Fields("fix_sta").Value.ToString.Trim

                      Case Is = "1"     '�觫���
                           imgStaFix = My.Resources._16x16_ledred
                           staFx = "�觫���"

                      Case Is = "2"     '�Ѻ�׹�觫���
                           imgStaFix = My.Resources._16x16_ledgreen
                           staFx = "�Ѻ�׹�觫���"

                      Case Is = "3"     '�Ѻ�׹�ҧ��ǹ
                           imgStaFix = My.Resources._16x16ledyellow
                           staFx = "�Ѻ�׹�ҧ��ǹ"

                      Case Else         '����
                           imgStaFix = My.Resources.blank
                           staFx = "����"

               End Select

                dgvFixDet.Rows.Add( _
                                      imgStaFix, _
                                      staFx, _
                                      .Fields("eqp_id").Value.ToString.Trim, _
                                      .Fields("size_id").Value.ToString.Trim, _
                                      .Fields("issue").Value.ToString.Trim, _
                                      Format(.Fields("amt_out").Value, "#0.0"), _
                                      Format(.Fields("amt_in").Value, "#0.0"), _
                                      Format(.Fields("price").Value, "#,###,##0.00"), _
                                      Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("fix_by").Value.ToString.Trim, _
                                      Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("recv_by").Value.ToString.Trim _
                                  )


                          CountFix = CountFix + Format(.Fields("amt_out").Value, "#,##0")
                          CountPrice = CountPrice + Format(.Fields("price").Value, "#,###,##0.00")

                   .MoveNext()
             Loop

             lblAmount.Text = CountFix
             lblAmt.Text = CountPrice.ToString("#,###,##0.00")

           End If

         .ActiveConnection = Nothing
         .Close()

      End With

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

  Dim staSize As String

      If dgvFixDet.RowCount <> 0 Then
         staSize = dgvFixDet.Rows(dgvFixDet.CurrentRow.Index).Cells(3).Value.ToString.Trim
         staAction = "1"
         LoadEditData(staSize)      '��Ŵ�����Ţ��������
         ShowGroupEdit()   '��ʴ� groupbox ����
      End If

End Sub

Private Sub ShowGroupEdit()

    With gpbFxDetail
         .Visible = True
         .Top = 209
         .Left = 12
         .Width = 999
         .Height = 511
         .Text = "�����������´�觫���"

         gpbHead.Enabled = False
         btnSaveData.Enabled = False
         txtSize.Focus()
   End With

End Sub

Private Sub LoadEditData(ByVal strSize As String)

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

       strSqlSelc = "SELECT * FROM tmp_fixeqptrn (NOLOCK)" _
                                & " WHERE size_id = '" & strSize & "'"

       Rsd = New ADODB.Recordset

       With Rsd

            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .CursorType = CursorTypeEnum.adOpenKeyset
            .Open(strSqlSelc, Conn, , , )

            If .RecordCount <> 0 Then
                txtSize.Text = .Fields("size_id").Value.ToString.Trim
                txtSetQty.Text = Format(.Fields("amt_out").Value, "##0.0")
                txtPrice.Text = Format(.Fields("price").Value, "#,###,##0.00")
                txtPr.Text = .Fields("pr_doc").Value.ToString.Trim
                txtSupp.Text = .Fields("sup_name").Value.ToString.Trim

                If .Fields("fix_date").Value.ToString.Trim <> "" Then
                   txtFixdate.Text = Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10)
                Else
                   txtFixdate.Text = "__/__/____"
                End If

                If .Fields("due_date").Value.ToString.Trim <> "" Then
                   txtDueDate.Text = Mid(.Fields("due_date").Value.ToString.Trim, 1, 10)
                Else
                   txtDueDate.Text = "__/__/____"
                End If

                txtFixnm.Text = .Fields("fix_by").Value.ToString.Trim
                txtIssue.Text = .Fields("issue").Value.ToString.Trim
                txtRmk.Text = .Fields("fix_rmk").Value.ToString.Trim
                txtSize.Enabled = False
             End If

           .ActiveConnection = Nothing
           .Close()

       End With

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click

   Dim strSize As String

       If dgvFixDet.RowCount <> 0 Then
          strSize = dgvFixDet.Rows(dgvFixDet.CurrentRow.Index).Cells(3).Value.ToString.Trim

          DeleteSubData(strSize)   'ź������
          ShowDataAFUpDel()
         'ShowDetailFx()    '��ʴ���¡���觫���� DataGridview
       End If

End Sub

Private Sub SaveEditsubFxEqp()   '�ѹ�֡�����������´�觫���

   Dim Conn As New ADODB.Connection
   Dim Rsd As New ADODB.Recordset
   Dim strSqlCmd As String

   Dim strFixdate As String
   Dim strDuedate As String

       With Conn

            If .State Then .Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()


                   Conn.BeginTrans()

                    '-------------------------- Ǵ�.����觫��� -----------------------------

                    If txtFixdate.Text <> "__/__/____" Then

                       strFixdate = Mid(txtFixdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtFixdate.Text.ToString, 1, 2)
                       strFixdate = "'" & SaveChangeEngYear(strFixdate) & "'"

                    Else
                       strFixdate = "NULL"
                    End If

                   '---------------------------- ��˹��Ѻ���  ------------------------------

                    If txtDueDate.Text <> "__/__/____" Then

                       strDuedate = Mid(txtDueDate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtDueDate.Text.ToString, 1, 2)
                       strDuedate = "'" & SaveChangeEngYear(strDuedate) & "'"

                   Else
                       strDuedate = "NULL"
                   End If

                   '------------------ Update �����ŵ��ҧ tmp_fixeqptrn -------------------

                   strSqlCmd = "UPDATE tmp_fixeqptrn SET amt_out = " & ChangFormat(txtSetQty.Text) _
                                           & "," & "price = " & ChangFormat(txtPrice.Text) _
                                           & "," & "pr_doc = '" & ReplaceQuote(txtPr.Text.ToUpper.Trim) & "'" _
                                           & "," & "sup_name = '" & ReplaceQuote(txtSupp.Text.ToUpper.Trim) & "'" _
                                           & "," & "fix_date =" & strFixdate _
                                           & "," & "fix_by = '" & ReplaceQuote(txtFixnm.Text.Trim) & "'" _
                                           & "," & "due_date = " & strDuedate _
                                           & "," & "issue = '" & ReplaceQuote(txtIssue.Text.Trim) & "'" _
                                           & "," & "fix_rmk =  '" & ReplaceQuote(txtRmk.Text.Trim) & "'"

                   .Execute(strSqlCmd)
                   Conn.CommitTrans()

                  gpbFxDetail.Visible = False
                  ShowDetailFx()                       '��ʴ���¡���觫���� DataGridview
                  staAction = ""

          Close()
          Conn = Nothing

       End With

End Sub

Private Sub DeleteSubData(ByVal Size As String)

   Dim Conn As New ADODB.Connection
   Dim Rsd As New ADODB.Recordset
   Dim strSqlCmd As String

   Dim btyConsider As Byte

       With Conn

            If .State Then .Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

               btyConsider = MsgBox("SIZE : " & Size.ToString.Trim & vbNewLine _
                                                   & "�س��ͧ���ź���������!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                   + MsgBoxStyle.Exclamation, "Confirm Delete Data")

               If btyConsider = 6 Then

                  strSqlCmd = "DELETE tmp_fixeqptrn" _
                                   & " WHERE size_id = '" & Size & "'"

                  .Execute(strSqlCmd)

               Else
                    dgvFixDet.Focus()
               End If

             .Close()
             Conn = Nothing

       End With

End Sub

Private Sub txtIssue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIssue.KeyPress

   If e.KeyChar = Chr(13) Then
      txtRmk.Focus()
   End If

End Sub

End Class