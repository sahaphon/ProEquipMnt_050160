Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeScreenBlk
 Dim IsShowSeek As Boolean        '�������ʴ�ʶҹ� gpbSeek
 Dim strDateDefault As String     '���������Ѻ�ѹ�������

 Public Const DrvName As String = "\\10.32.0.15\data1\EquipPicture\"
 Public Const PthName As String = "\\10.32.0.15\data1\EquipPicture"         '���������Ѻ�� part �ٻ�Ҿ

Protected Overrides ReadOnly Property CreateParams() As CreateParams       '��ͧ�ѹ��ûԴ������� Close Button(�����ҡ�ҷ)
   Get
       Dim cp As CreateParams = MyBase.CreateParams
           Const CS_DBLCLKS As Int32 = &H8
           Const CS_NOCLOSE As Int32 = &H200
           cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
           Return cp
   End Get
End Property

Private Sub frmAeScreenBlk_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTable(0, "")  'ź������ Table tmp_eqptrn where user_id..
   frmScreenBlk.lblCmd.Text = "0"  '������ʶҹ�
   Me.Dispose()     '����¿���� �׹˹��¤�����
End Sub

Private Sub frmAeScreenBlk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Dim dtComputer As Date = Now()
  Dim strCurrentDate As String

      StdDateTimeThai()        '�Ѻ�ٷչ��ŧ�ѹ�����Ǵ�.�� ����� Module
      strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

      ClearDataGpbHead()
      PrePartSeek()            '��Ŵ��������´���� cbo ������ǹ����Ե
      'txtEqp_id.Focus()

        Select Case frmScreenBlk.lblCmd.Text.ToString

               Case Is = "0" '����������

                    With txtBegin                 '��Ŵ�ѹ���Ѩ�غѹ���� txtBegin
                         .Text = strCurrentDate
                         strDateDefault = strCurrentDate
                    End With

                    With Me
                         .Text = "����������"
                    End With

                '---------------------������������������ͧ�ʴ�ʶҹ�(��͹�������� Gridview)----------------------------

                dgvSize.Columns(0).Visible = False  '��͹��������� 1
                dgvSize.Columns(1).Visible = False  '��͹��������� 2
                dgvSize.Columns(2).Visible = False  '��͹��������� 3
                dgvSize.Columns(3).Visible = False  '��͹��������� 4

              Case Is = "1" '��䢢�����

                With Me
                     .Text = "���䢢�����"
                End With

                LockEditData()
                txtEqp_id.ReadOnly = True   '�����ҹ���ҧ����
                txtEqpnm.ReadOnly = True
                txtShoe.ReadOnly = True
                txtOrder.ReadOnly = True
                txtRemark.ReadOnly = True

              Case Is = "2"   '����ͧ������

                With Me
                     .Text = "����ͧ������"
                End With

                LockEditData()
                txtEqp_id.ReadOnly = True  '�����ҹ���ҧ����
                btnSaveData.Enabled = False

        End Select

    txtEqp_id.Focus()
End Sub

Private Sub LockEditData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim Rsdwc As New ADODB.Recordset

  Dim strCmd As String  ' ��ʵ�ԧ Command

  Dim blnHavedata As Boolean   '�纤�ҵ����� ����Ѻ������բ������������
  Dim strSqlSelc As String = ""   '��ʵ�ԧ sql select
  Dim strPart As String = ""
        '�麤�� Row �Ѩ�غѹ㹿���� frmScreenBlk
  Dim strCod As String = frmScreenBlk.dgvScreenBlk.Rows(frmScreenBlk.dgvScreenBlk.CurrentRow.Index).Cells(0).Value.ToString.Trim

        With Conn

            If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

        End With

        strSqlSelc = "SELECT * " _
                                    & "FROM v_moldinj_hd (NOLOCK)" _
                                    & " WHERE eqp_id = '" & strCod & "'"

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , )

         If .RecordCount <> 0 Then

                txtBegin.Text = .Fields("creat_date").Value.ToString.Trim
                strDateDefault = .Fields("creat_date").Value.ToString.Trim

                txtEqp_id.Text = .Fields("eqp_id").Value.ToString.Trim
                txtEqpnm.Text = .Fields("eqp_name").Value.ToString.Trim
                txtShoe.Text = .Fields("shoe").Value.ToString.Trim
                txtAmount.Text = .Fields("pi_qty").Value.ToString.Trim
                txtSet.Text = Format(.Fields("set_qty").Value, "##0.0")
                txtRemark.Text = .Fields("remark").Value.ToString.Trim


                strCmd = frmScreenBlk.lblCmd.Text.ToString.Trim    '��� strCmd ��ҡѺ���� lblcmd 㹿���� frmEqpSheet

                Select Case strCmd
                    Case Is = "1"   '�����ͤ�͹���
                    Case Is = "2"   '�����ͤ�͹����ͧ
                        btnSaveData.Enabled = False  '�Դ���� "�ѹ�֡������"
                End Select

                '----------------- Insert ������ŧ���ҧ tmp_eqptrn ����Ң����Ũҡ���ҧ  tmp_eqptrn �Ҵ���----------------------

                strSqlSelc = "INSERT INTO tmp_eqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM eqptrn " _
                                  & " WHERE eqp_id = '" & strCod & "' "

                Conn.Execute(strSqlSelc)
                blnHavedata = True     '�բ�����

        Else
                blnHavedata = False    '����բ�����
        End If

            .ActiveConnection = Nothing   '��觻Դ�����������
            .Close()
        End With

        Rsd = Nothing   '�������� RecordSet
        Conn.Close()    '��觵Ѵ�����������
        Conn = Nothing  '������ Connection

             If blnHavedata Then          '��� blnHavedata = true
                ShowScrapItem()
             End If

End Sub

Private Sub ClearDataGpbHead()
  txtEqp_id.Text = ""
  txtEqpnm.Text = ""
  txtShoe.Text = ""
  txtOrder.Text = ""
  txtAmount.Text = ""
  txtSet.Text = ""
End Sub

Private Sub PrePartSeek()
 Dim strGpTopic(8) As String
 Dim i As Integer

     strGpTopic(0) = "��鹺�"
     strGpTopic(1) = "����鹺�"
     strGpTopic(2) = "������"
     strGpTopic(3) = "�����ҧ"
     strGpTopic(4) = "EVA �Դ��"
     strGpTopic(5) = "������"
     strGpTopic(6) = "˹ѧ˹��"
     strGpTopic(7) = "EVA ��˹ѧ˹��"
     strGpTopic(8) = "ONUPPER"

     With cboPart

           For i = 0 To 8

               .Items.Add(strGpTopic(i))

           Next

     End With

End Sub

Private Sub ClearTmpTable(ByVal byOption As Byte, ByVal strPsID As String)
 Dim Conn As New ADODB.Connection
 Dim strSqlcmd As String

     With Conn

            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .CommandTimeout = 90
            .Open()

            Select Case byOption

                   Case Is = "0"  'ź��������ѧ�Դ�����
                          strSqlcmd = "DELETE tmp_eqptrn " _
                                           & "WHERE user_id = '" & frmMainPro.lblLogin.Text & "'"
                          .Execute(strSqlcmd)

                   Case Is = "1"
                          strSqlcmd = "DELETE tmp_eqptrn " _
                                        & "WHERE user_id = '" & frmMainPro.lblLogin.Text & "'" _
                                        & "AND docno ='" & strPsID.ToString.Trim & "'"
                          .Execute(strSqlcmd)
            End Select

     End With
 Conn.Close()
 Conn = Nothing

End Sub

Private Sub ShowScrapItem()                     '�ʴ�������� DataGridview 
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String             '�纤�� string command

 Dim dubQty As Double
 Dim dubAmt As Double

 Dim sngSetQty As Single                 '�纨ӹǹ SET

        With Conn
            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

        End With

        strSqlCmdSelc = "SELECT * FROM v_tmp_eqptrn (NOLOCK)" _
                                 & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "' " _
                                 & "ORDER BY size_desc, size_id "

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlCmdSelc, Conn, , , )

            dgvSize.Rows.Clear()
            dgvSize.ScrollBars = ScrollBars.None                 '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ

            If .RecordCount <> 0 Then

                Do While Not .EOF()

                    dgvSize.Rows.Add( _
                                        IIf(.Fields("delvr_sta").Value.ToString.Trim = "1", My.Resources.accept, My.Resources._16x16_ledred), _
                                        "", _
                                        My.Resources.blank, _
                                        "", _
                                        .Fields("size_id").Value.ToString.Trim, _
                                        .Fields("size_act").Value.ToString.Trim, _
                                        .Fields("size_desc").Value.ToString.Trim, _
                                        .Fields("size_group").Value.ToString.Trim, _
                                        .Fields("backgup").Value.ToString.Trim, _
                                        Format(.Fields("set_qty").Value, "##0.0"), _
                                        Format(.Fields("size_qty").Value, "##0.0"), _
                                        .Fields("dimns").Value.ToString.Trim, _
                                        .Fields("price").Value, _
                                        .Fields("ord_rep").Value, _
                                        .Fields("ord_qty").Value, _
                                        .Fields("men_rmk").Value.ToString.Trim _
                                    )

                    sngSetQty = sngSetQty + .Fields("set_qty").Value
                    dubQty = dubQty + .Fields("ord_qty").Value
                    dubAmt = dubAmt + .Fields("price").Value

                    .MoveNext()

                Loop

                txtSet.Text = sngSetQty.ToString.Trim        '�ӹǹ SET
                txtAmount.Text = Format(dubQty, "#,##0")     '������ŧ��Ե
                lblAmt.Text = Format(dubAmt, "#,##0.00")     '����Ҥ��ػ�ó�

            Else
                txtSet.Text = "0.0"
                txtAmount.Text = "0"
                lblAmt.Text = "0.00"

            End If

            .ActiveConnection = Nothing
            .Close()
            Rsd = Nothing

            dgvSize.ScrollBars = ScrollBars.Both '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ

        End With

 Conn.Close()
 Conn = Nothing

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBfSave()
End Sub

Private Sub CheckDataBfSave()
 Dim IntListwc As Integer = dgvSize.Rows.Count
 Dim strProd As String = ""
 Dim strProdnm As String = ""

 Dim bytConSave As Byte  '�纤�� megbox 

     If txtEqp_id.Text <> "" Then

           If txtEqpnm.Text <> "" Then

                          If IntListwc > 0 Then

                             bytConSave = MsgBox("�س��ͧ��úѹ�֡���������������!" _
                                  , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                    If bytConSave = 6 Then

                                             Select Case Me.Text

                                                    Case Is = "����������"

                                                         If CheckCodeDuplicate(txtEqp_id.Text) Then   '�����ʫ��
                                                            SaveNewRecord()

                                                         Else
                                                             MessageBox.Show("�����ػ�ó��� ��سҡ�͡�����ػ�ó�����!....", _
                                                                                  "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                             txtEqp_id.Text = ""
                                                             txtEqp_id.Focus()

                                                         End If

                                                     Case Else
                                                            SaveEditRecord()



                                             End Select

                                    Else
                                          dgvSize.Focus()
                                    End If

                          Else

                                If CheckCodeDuplicate(txtEqp_id.Text) Then           '��Ǩ�ͺ�����ػ�ó���
                                   ShowResvrd()       '�ʴ���������� gpbSeek 
                                   gpbSeek.Text = "����������"
                                   txtSize.ReadOnly = False

                                Else

                                   MessageBox.Show("��سҡ�͡�����ػ�ó�����!....", "�����ػ�ó���!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)

                                   txtEqp_id.Text = ""
                                   txtEqp_id.Focus()
                                End If

                          End If

           Else
                MsgBox("�ô�кآ�������������´�ػ�ó�  " _
                        & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                txtEqpnm.Focus()

           End If

     Else
          MsgBox("�ô�кآ����������ػ�ó�  " _
                        & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
          txtEqp_id.Focus()

     End If

End Sub

Private Sub SaveNewRecord()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String
  Dim dateSave As Date = Now()    '�纤���ѹ���Ѩ�غѹ
  Dim strDate As String

  Dim strCredate As String
  Dim strDateDoc As String
  Dim strType As String = ""
  Dim Rsd As New ADODB.Recordset

      With Conn

           If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      Conn.BeginTrans()

             strDate = Date.Now.ToString("yyyy-MM-dd")
             strDate = SaveChangeEngYear(strDate)            '��ŧ�繻� �.�.

             strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                   & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                   & Mid(txtBegin.Text.Trim, 1, 2)
                    strDateDoc = SaveChangeEngYear(strDateDoc)

                '---------------------------------------- Ǵ�.����Դ ----------------------------------------------

                   If txtCdate.Text <> "__/__/____" Then

                       strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtCdate.Text.ToString, 1, 2)
                       strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

                   Else
                       strCredate = "NULL"
                   End If

                    strSqlCmd = "INSERT INTO eqpmst" _
                      & "(prod_sta,fix_sta,[group],eqp_id,eqp_name" _
                      & ",pi,shoe,ap_code,ap_desc,doc_ref,set_qty" _
                      & ",part,eqp_type" _
                      & ",pic_ctain,pic_io,pic_part,remark" _
                      & ",tech_desc,tech_thk,tech_lg,tech_sht,tech_eva,tech_warm" _
                      & ",tech_time1,tech_time2,creat_date,pre_date,pre_by,pi_qty" _
                      & ",eqp_amt,exp_id,tech_trait" _
                      & ")" _
                      & " VALUES (" _
                      & "'" & "0" & "'" _
                      & ",'" & "0" & "'" _
                      & ",'" & "F" & "'" _
                      & ",'" & ReplaceQuote(txtEqp_id.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtEqpnm.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(txtRef.Text.ToString.Trim) & "'" _
                      & ",'" & ChangFormat(txtSet.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & " '" _
                      & ",'" & "" & " '" _
                      & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & " '" _
                      & ",'" & "" & " '" _
                      & "," & strCredate _
                      & ",'" & strDate & "'" _
                      & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                      & ",'" & ChangFormat(txtAmount.Text.ToString.Trim) & "'" _
                      & ",'" & RetrnAmount() & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ")"

                Conn.Execute(strSqlCmd)

               '------------------------------------------------�ѹ�֡������㹵��ҧ eqptrn----------------------------------------------------------

                strSqlCmd = "INSERT INTO eqptrn " _
                                     & " SELECT [group] ='F'" _
                                     & ",eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'" _
                                     & ",size_id,size_desc,size_qty,weight,dimns,backgup,price,men_rmk" _
                                     & ",delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                     & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & " FROM tmp_eqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                Conn.Execute(strSqlCmd)
                Conn.CommitTrans()

                frmScreenBlk.lblCmd.Text = txtEqp_id.Text.ToString.Trim   '�觺͡��Һѹ�֡�����������
                frmScreenBlk.Activating()
                Me.Close()

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub SaveEditRecord()

 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim dateSave As Date = Now()
 Dim strDate As String = ""

 Dim strCredate As String
 Dim strDocdate As String           '��ʵ�ԧ�ѹ����͡���
 Dim strGpType As String = ""       '�纻������ػ�ó�
 Dim strPartType As String = ""     '�纪����ǹ����Ե

        With Conn

            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()

        End With

               Conn.BeginTrans()      '�ش������� Transection

               strDate = dateSave.Date.ToString("yyyy-MM-dd")
               strDate = SaveChangeEngYear(strDate)

              '------------------------- �ѹ����͡��� ----------------------------------------------------

               strDocdate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocdate = SaveChangeEngYear(strDocdate)


              '---------------------------------------- Ǵ�.����Դ --------------------------------------------

               If txtCdate.Text <> "__/__/____" And txtCdate.Text <> "" Then

                  strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtCdate.Text.ToString, 1, 2)
                  strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

               Else
                  strCredate = "NULL"

               End If

                      '---------------------------------- UPDATE ������㹵��ҧ eqpmst ------------------------

                      strSqlCmd = "UPDATE  eqpmst SET eqp_name ='" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                                & "," & "pi ='" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                                                & "," & "shoe ='" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                                                & "," & "part ='" & "" & "'" _
                                                & "," & "eqp_type ='" & " " & "'" _
                                                & "," & "set_qty =" & ChangFormat(txtSet.Text.ToString.Trim) _
                                                & "," & "pic_ctain ='" & "" & "'" _
                                                & "," & "pic_io ='" & "" & "'" _
                                                & "," & "pic_part ='" & "" & " '" _
                                                & "," & "remark ='" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                                & "," & "tech_desc = '" & "" & "'" _
                                                & "," & "tech_thk = '" & "" & "'" _
                                                & "," & "tech_lg = '" & "" & " '" _
                                                & "," & "tech_sht = '" & "" & "'" _
                                                & "," & "tech_eva = '" & "" & "'" _
                                                & "," & "tech_warm = '" & "" & "'" _
                                                & "," & "tech_time1 = '" & "" & "'" _
                                                & "," & "tech_time2 = '" & "" & " '" _
                                                & "," & "creat_date = " & strCredate _
                                                & "," & "eqp_amt = " & RetrnAmount() _
                                                & "," & "last_date = '" & strDate & "'" _
                                                & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                                & "," & "exp_id ='" & "" & "'" _
                                                & "," & "tech_trait ='" & "" & "'" _
                                                & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)

                      '------------------------------- ź������㹵��ҧ eqptrn ---------------------------------------

                     strSqlCmd = "Delete FROM eqptrn" _
                                          & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlCmd)

                      '-------------------------------- �ѹ�֡������㹵��ҧ eqptrn �� Select �ҡ tmp_eqptrn ---------------

                     strSqlCmd = "INSERT INTO eqptrn " _
                                      & "SELECT [group] = 'F'" _
                                      & ",eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'" _
                                      & ",size_id,size_desc,size_qty,weight,dimns,backgup" _
                                      & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date" _
                                      & ",pr_doc,recv_date,ord_rep,ord_qty,fc_date,impt_id" _
                                      & ",sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                      & " FROM tmp_eqptrn " _
                                      & " WHERE user_id= '" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                    Conn.Execute(strSqlCmd)
                    Conn.CommitTrans()  '��� Commit transection

                    frmScreenBlk.lblCmd.Text = txtEqp_id.Text.ToString.Trim   '�觺͡��Һѹ�֡�����������
                    frmScreenBlk.Activating()
                    Me.Close()

     Conn.Close()
     Conn = Nothing

End Sub

 Private Function RetrnAmount() As String   '�ѧ�������ʹ��� �Ҥ��ػ�ó��� UserLogin

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

        '-------------- ����� SELCT SUM()AS ����������� ---------------------------------------

        strSqlSelc = "SELECT SUM(price)  AS sm_amt " _
                            & " FROM tmp_eqptrn " _
                            & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                            & " GROUP BY user_id"

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            If .RecordCount <> 0 Then
                RetrnAmount = .Fields("sm_amt").Value.ToString.Trim
            Else
                RetrnAmount = "0"
            End If

            .ActiveConnection = Nothing
            .Close()

        End With

        Rsd = Nothing

        Conn.Close()
        Conn = Nothing

End Function

Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String) As Boolean
 Dim fname As String = String.Empty
 Dim dFile As String = String.Empty
 Dim dFilePath As String = String.Empty

 Dim fServer As String = String.Empty
 Dim intResult As Integer

     On Error GoTo Err70

        fname = strPicPath & "\" & strPicName  '���� \\10.32.0.15\data1\EquipPicture\"�����ٻ�Ҿ" 
        fServer = PthName & "\" & strPicName    'partServer \\10.32.0.15\data1\EquipPicture\"�����ٻ�Ҿ"

         If File.Exists(fServer) Then    '�������������ԧ
            CallCopyPicture = True      '���׹��� true

         Else
            If File.Exists(fname) Then
               dFile = Path.GetFileName(fname)
               dFilePath = DrvName + dFile


               intResult = String.Compare(fname.ToString.Trim, dFilePath.ToString.Trim)

                '--------------------------- ��Ҥ���� 0 �ʴ������Ŵ��������� �������ö Copy ����� ------------------------------

                    If intResult = 1 Then '��ҷ���� = 1 �֧ copy �ٻ�����������ͧ 10.32.0.14
                       File.Copy(fname, dFilePath, True)
                    End If
                    CallCopyPicture = True

            Else
                CallCopyPicture = True

            End If

        End If

Err70:

      If Err.Number <> 0 Then

         MsgBox("UserName �ͧ�س������Է������ٻ�Ҿ��!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Permission Can't Edit Picture")
         CallCopyPicture = True

      End If

End Function

Private Function CheckCodeDuplicate(ByVal strCod As String) As Boolean
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

        strSqlSelc = "SELECT eqp_id FROM eqpmst" _
                              & " WHERE eqp_id = '" & strCod & "'"


        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            If .RecordCount <> 0 Then
                CheckCodeDuplicate = False

            Else
                CheckCodeDuplicate = True

            End If
            .ActiveConnection = Nothing   '������ Connection
            .Close()

        End With
        Rsd = Nothing   '������ RecordSet

  Conn.Close()    '�Դ�����������
  Conn = Nothing   '������ RecordSet

End Function

Private Sub ShowResvrd()  '����ʴ� GroupBox gpbSeek �����

 tabMain.SelectedTab = tabSize
 IsShowSeek = Not IsShowSeek  '���ʶҹ�ᶺ seek ����ʴ� ����ʴ�

   If IsShowSeek Then

       With gpbSeek
            .Visible = True
            .Left = 8    '᡹ X
            .Top = 230   '᡹ Y 252
            .Height = 500
            .Width = 1014
       End With

            StateLockFindDept(False) ' ��ͤ FindDept ���觤���� False �
   Else
            StateLockFindDept(True)

        End If
   End Sub

Private Sub StateLockFindDept(ByVal sta As Boolean)

 Dim strMode As String = frmScreenBlk.lblCmd.Text.ToString   '����� lblCmd ������ѹ�֡�����ź��ͤ���� �觤������� strMode 
     btnAdd.Enabled = Sta    '��������������
     gpbHead.Enabled = Sta
     tabMain.Enabled = Sta
     btnSaveData.Enabled = Sta  '�����ѹ�֡������

     Select Case strMode

            Case Is = "1" '��䢢�����                        
            Case Is = "2" '����ͧ������
                btnSaveData.Enabled = False

     End Select

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  On Error Resume Next    '�ҡ�Դ error ������ѧ�зӧҹ���������ʹ� error ����Դ���
  Dim strCode As String

        If MessageBox.Show("��ͧ����͡�ҡ����� �������", "��س��׹�ѹ�͡�ҡ�����", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            With frmScreenBlk.dgvScreenBlk
                If .Rows.Count > 0 Then   '����բ������ Grid
                    strCode = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim          '���strCode = ��������ǻѨ�غѹ Cell �á
                    lblComplete.Text = strCode  '��� label �ʴ�������� Cell �Ѩ�غѹ   

                End If
            End With
            Me.Close()

            frmMainPro.Show()
            frmScreenBlk.Show()

        End If
End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
  ShowResvrd()
  ClearSubData2()
  'CallEditData()    '�Ѻ�ٷչ�ʴ� Size ������䢢�����
  gpbSeek.Text = "����������"
  cboPart.Enabled = True
  txtSize.ReadOnly = False
  txtSizeDesc.ReadOnly = False
End Sub

Private Sub ClearSubData2()
   txtCdate.Text = "__/__/____"
   txtSize.Text = ""
   txtSizeDesc.Text = ""
   txtSizeQty.Text = "0"
   txtSetQty.Text = "0"
   txtPrice.Text = "0.00"
   txtRmk.Text = ""
End Sub

Private Sub CallEditData()
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

   strSqlSelc = " SELECT creat_date FROM eqpmst (NOLOCK)" _
                            & " WHERE eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'"

   With Rsd

         .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
         .LockType = ADODB.LockTypeEnum.adLockOptimistic
         .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then

            If .Fields("creat_date").Value.ToString <> "" Then
                 txtCdate.Text = Mid(.Fields("creat_date").Value.ToString.Trim, 1, 10)
            Else
                 txtCdate.Text = "__/__/____"
            End If

         Else
            txtCdate.Focus()

         End If
         .ActiveConnection = Nothing
         .Close()

   End With
   Rsd = Nothing

Conn.Close()
Conn = Nothing
End Sub

Private Sub CallEditData2()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

     If dgvSize.Rows.Count <> 0 Then

        Dim strSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(4).Value.ToString.Trim         '�� Size
        Dim strGroupSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(7).Value.ToString.Trim      '�����ʺ��ͤ

        With Conn

               If .State Then Close()
                  .ConnectionString = strConnAdodb
                  .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                  .ConnectionTimeout = 90
                  .Open()

        End With

          strSqlSelc = " SELECT * " _
                          & " FROM v_tmp_eqptrn (NOLOCK)" _
                          & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                          & " AND size_id= '" & strSize & "'" _
                          & " AND size_group = '" & strGroupSize & "'"

          With Rsd

                 .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                 .LockType = ADODB.LockTypeEnum.adLockOptimistic
                 .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                Select Case .Fields("backgup").Value.ToString.Trim

                    Case Is = "�����ҧ"
                        cboPart.Text = "�����ҧ"
                    Case Is = "��鹺�"
                        cboPart.Text = "��鹺�"
                    Case Is = "����鹺�"
                       cboPart.Text = "����鹺�"
                    Case Is = "������"
                        cboPart.Text = "������"
                    Case Is = "������"
                        cboPart.Text = "������"
                    Case Is = "EVA �Դ��"
                        cboPart.Text = "EVA �Դ��"
                    Case Is = "EVA ��˹ѧ˹��"
                        cboPart.Text = "EVA ��˹ѧ˹��"
                    Case Is = "˹ѧ˹��"
                       cboPart.Text = "˹ѧ˹��"
                    Case Is = "ONUPPER"
                       cboPart.Text = "ONUPPER"

                End Select

                     txtSize.Text = .Fields("size_id").Value.ToString.Trim
                     txtSizeDesc.Text = .Fields("size_group").Value.ToString.Trim
                     txtSetQty.Text = Format(.Fields("set_qty").Value, "#,##0.0")
                     txtSizeQty.Text = Format(.Fields("size_qty").Value, "#,##0.0")
                     txtPrice.Text = Format(.Fields("price").Value, "#,##0.00")
                     txtRmk.Text = .Fields("men_rmk").Value.ToString.Trim

         End If
               .ActiveConnection = Nothing    '����������������
               .Close()

         End With
            Rsd = Nothing

    Conn.Close()
    Conn = Nothing
End If

End Sub

Private Function RetrnDiams(ByVal strDia As String, ByRef strW As String, ByRef strL As String) As Boolean
Dim i, x As Integer
Dim strTmp As String = ""
Dim strMerg As String = ""

Dim strDiamns(1) As String  '��ʵ�ԧ��������
Dim y As Integer = 0

                 x = Len(strDia)
                 For i = 1 To x
                         strTmp = Mid(strDia, i, 1)
                         Select Case strTmp

                                Case Is = "x"       '���������ͧ���¤ٳ
                                          strDiamns(y) = strMerg
                                          y = y + 1
                                          strMerg = ""

                                Case Else

                                     If InStr(",.0123456789", strTmp) Then
                                        strMerg = strMerg & strTmp
                                     End If
                         End Select
                 Next i

strW = strDiamns(0)
strL = strMerg
'strH = strMerg

End Function

Private Sub btnSeekExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekExit.Click
  StateLockFindDept(True)
  gpbSeek.Text = ""
  gpbSeek.Visible = False  '����� gpbSeek ��͹
  IsShowSeek = False
End Sub

Private Sub btnSeekSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekSave.Click
   CheckSubDataBfSave()
End Sub

Private Sub CheckSubDataBfSave()
 Dim i As Integer

     If txtSize.Text.Trim <> "" Then

        If gpbSeek.Text = "����������" Then
           SaveSubRecord()
        Else
            EditSubRecord()
        End If

            ShowScrapItem()  '�ʴ�����ŷ���ѹ��� dgvSize �� Select �ҡ v_tmp_eqptrn

            '------------------------------�������ʷ��������������------------------------------------------

            For i = 1 To dgvSize.Rows.Count - 1

                If dgvSize.Rows(i).Cells(4).Value.ToString = txtSize.Text.ToString.Trim Then    '��Ҥ������ Size � dgvSize �դ����ҡѺ txtSize
                   dgvSize.CurrentCell = dgvSize.Item(5, i)
                   dgvSize.Focus()
                   End If
               Next i

              StateLockFindDept(True)
              gpbSeek.Visible = False
              IsShowSeek = False


         Else
               MsgBox("�ô�к� Size" _
                                   & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                             txtSize.Focus()
         End If

End Sub

Private Function SaveSubRecord() As Boolean

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlSelec As String
  Dim strSqlCmd As String

  Dim dateSave As Date = Now()            '��ʵ�ԧ�ѹ���Ѩ�غѹ
  Dim strDate As String = ""              '��ʵ�ԧ�ѹ���
  Dim strDateDoc As String = ""
  Dim strCreDate As String = ""           '�ѹ����Ե
  Dim strPrdate As String = ""            '�ѹ����Դ� PR
  Dim strIndate As String = ""
  Dim strFcDate As String = ""
  Dim strPartType As String = ""
  Dim strEqpType As String = ""
  Dim strDateNull As String = "NULL"       '�ѹ�������ҧ(Null)  

     Try

      With Conn
           If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

     End With

        '------------------------------------ �礢�������͹����������������� -------------------------------------------------

        strSqlSelec = "SELECT size_id FROM tmp_eqptrn" _
                           & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                           & " AND size_group = '" & txtSizeDesc.Text.ToString.Trim & "'" _
                           & " AND size_id = '" & txtSize.Text.ToString.Trim & "'"

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelec, Conn, , , )

            If .RecordCount <> 0 Then        '��� RecordSet �բ�����
                MessageBox.Show("Size :" & txtSize.Text.ToString & " ����к����� ��س��к� Size ����", "�����ū��!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                SaveSubRecord = False

            Else

                strDate = dateSave.Date.ToString("yyyy-MM-dd")
                strDate = SaveChangeEngYear(strDate)

                strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                      & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                      & Mid(txtBegin.Text.ToString, 1, 2)
                strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"



                '---------------------------------------- Ǵ�.����Դ --------------------------------

                If txtCdate.Text <> "__/__/____" Then

                    strCreDate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                         & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                         & Mid(txtCdate.Text.ToString, 1, 2)
                    strCreDate = "'" & SaveChangeEngYear(strCreDate) & "'"

                Else
                    strCreDate = "NULL"
                End If

                strSqlCmd = "INSERT INTO tmp_eqptrn " _
                                     & "(user_id,[group],eqp_id,size_id,size_desc,size_qty,weight" _
                                     & ",dimns,backgup,price,men_rmk,delvr_sta,sent_sta,set_qty" _
                                     & ",pr_date,pr_doc,recv_date,ord_rep,ord_qty,fc_date,impt_id" _
                                     & ",sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & ")" _
                                     & " VALUES (" _
                                     & "'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                     & ",'" & "F" & "'" _
                                     & ",'" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                     & ",'" & "" & "'" _
                                     & "," & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                     & "," & 0.0 _
                                     & ",'" & "" & "'" _
                                     & ",'" & cboPart.Text.ToString.Trim & "'" _
                                     & "," & ChangFormat(txtPrice.Text.ToString.Trim) _
                                     & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                     & ",'" & "0" & "'" _
                                     & ",'" & "0" & "'" _
                                     & "," & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                     & "," & strDateNull _
                                     & ",'" & "" & "'" _
                                     & "," & strDateNull _
                                     & ",'" & "0" & "'" _
                                     & ",'" & "0" & "'" _
                                     & "," & strCreDate _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & "," & 0.0 _
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

     Catch ex As Exception
           MsgBox("����ͼԴ��Ҵ��зӡ�úѹ�֡ �ô���Թ��������ա����", MsgBoxStyle.Critical, "�Դ��Ҵ")
           MsgBox(ex.Message)
     End Try

End Function

Private Sub EditSubRecord()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

  Dim dateSave As Date = Now()
  Dim strDate As String = ""

  Dim strCredate As String
  Dim strDocdate As String            '��ʵ�ԧ�ѹ����͡���
  Dim strGpType As String = ""        '�纻������ػ�ó�
  Dim strPartType As String = ""      '�纪����ǹ����Ե
  Dim strDateNull As String = "NULL"

      Try

        With Conn
            If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()

        End With


               strDate = dateSave.Date.ToString("yyyy-MM-dd")
               strDate = SaveChangeEngYear(strDate)

              '------------------------- �ѹ����͡��� ----------------------------------------------------

               strDocdate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocdate = SaveChangeEngYear(strDocdate)


             '---------------------------------------- Ǵ�.����Դ --------------------------------------------

              If txtCdate.Text <> "__/__/____" Then

                 strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtCdate.Text.ToString, 1, 2)
                 strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

              Else
                 strCredate = "NULL"

              End If

                       strSqlCmd = "UPDATE  tmp_eqptrn SET size_qty =" & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                                        & "," & "dimns ='" & "" & "'" _
                                                        & "," & "price = " & ChangFormat(txtPrice.Text.ToString.Trim) _
                                                        & "," & "backgup = '" & cboPart.Text.ToString.Trim & "'" _
                                                        & "," & "pr_doc ='" & "" & "'" _
                                                        & "," & "men_rmk ='" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                                        & "," & "set_qty =" & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                                        & "," & "pr_date = " & strDateNull _
                                                        & "," & "recv_date = " & strDateNull _
                                                        & "," & "fc_date = " & strDateNull _
                                                        & "," & "sup_name = '" & "" & "'" _
                                                        & "," & "lp_type = '" & "" & "'" _
                                                        & "," & "size_group = '" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                                        & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                                        & " AND size_id ='" & txtSize.Text.ToString.Trim & "'" _
                                                        & " AND size_group = '" & txtSizeDesc.Text.ToString.Trim & "'"

                       Conn.Execute(strSqlCmd)

   Conn.Close()
   Conn = Nothing

   Catch ex As Exception
         MsgBox("����ͼԴ��Ҵ��зӡ�úѹ�֡ �ô���Թ��������ա����", MsgBoxStyle.Critical, "�Դ��Ҵ")
         MsgBox(ex.Message)
   End Try

End Sub

Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

 If dgvSize.Rows.Count > 0 Then
     ShowResvrd()       '�ʴ� gpbSeek 
     CallEditData()    '�Ѻ�ٷչ�ʴ� Size ������䢢�����
     CallEditData2()   '�Ѻ�ٷչ�ʴ������ŷҧ෤�Ԥ

     gpbSeek.Text = "��䢢�����"
     cboPart.Enabled = True
     txtSize.ReadOnly = True
     txtSizeDesc.ReadOnly = True
  Else
      MsgBox("�������¡�� SIZE ����ͧ������!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
      dgvSize.Focus()
  End If

End Sub

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
  DeleteSubData()
End Sub

Private Sub DeleteSubData()
 Dim btyConsider As Byte
 Dim strSize As String = ""
 Dim strSizeAct As String = ""
 Dim strSizeBlock As String = ""
 Dim strGpsize As String

     With dgvSize

        If .Rows.Count > 0 Then

            strSize = .Rows(.CurrentRow.Index).Cells(4).Value.ToString
            strSizeAct = .Rows(.CurrentRow.Index).Cells(5).Value.ToString
            strSizeBlock = .Rows(.CurrentRow.Index).Cells(6).Value.ToString
            strGpsize = .Rows(.CurrentRow.Index).Cells(7).Value.ToString

             If strSizeAct <> "" Then

                    btyConsider = MsgBox("SIZE : " & strSize.ToString.Trim & vbNewLine _
                                                   & "���ʺ��ͤ : " & strSizeBlock.ToString.Trim & vbNewLine _
                                                   & "����䫵� : " & strGpsize.ToString.Trim & vbNewLine _
                                                   & "�س��ͧ���ź���������!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                   + MsgBoxStyle.Exclamation, "Confirm Delete Data")

                   If btyConsider = 6 Then
                      Dim Conn As New ADODB.Connection
                      Dim strCmd As String

                         If Conn.State Then Close()

                              Conn.ConnectionString = strConnAdodb
                              Conn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                              Conn.ConnectionTimeout = 90
                              Conn.Open()

                              strCmd = " DELETE FROM tmp_eqptrn" _
                                                  & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                  & " AND size_id = '" & strSize.ToString.Trim & "'" _
                                                  & " AND size_desc = '" & strSizeBlock.ToString.Trim & "'" _
                                                  & " AND size_group = '" & strGpsize.ToString.Trim & "'"

                              Conn.Execute(strCmd)
                              Conn.Close()
                              Conn = Nothing

                              .Rows.RemoveAt(.CurrentRow.Index)  'ź������ Cell �Ѩ�غѹ
                              ShowScrapItem()

                           End If
                   Else
                      .Focus()
                   End If

        Else
               MsgBox("�������¡�� SIZE ����ͧ���ź������!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
               dgvSize.Focus()

        End If

    End With
End Sub

Private Sub txtEqp_id_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqp_id.KeyDown

 Dim intChkPoint As Integer
        With txtEqp_id
            Select Case e.KeyCode
                Case Is = 35 '���� End 

                Case Is = 36 '���� Home

                Case Is = 37 '�١�ë���

                Case Is = 38 '�����١�â��

                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtEqpnm.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtEqpnm.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtOrder.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With

End Sub

Private Sub txtEqp_id_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqp_id.KeyPress

  Select Case Asc(e.KeyChar)

         Case 45 To 122 ' �������ѧ��������ԧ�������� 58�֧122 ������� 48��������ҵ�ͧ��õ���Ţ
              e.Handled = False

         Case 8, 46 ' Backspace = 8,  Delete = 46
              e.Handled = False

         Case 13   'Enter = 13
              e.Handled = False
              txtEqpnm.Focus()

         Case Else
              e.Handled = True
              MsgBox("��س��кآ������������ѧ���", MsgBoxStyle.Critical, "�Դ��Ҵ")
  End Select

End Sub

Private Sub txtEqp_id_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqp_id.LostFocus
 txtEqp_id.Text = txtEqp_id.Text.ToUpper.Trim
End Sub

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
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtShoe.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtShoe.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtOrder.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtEqpnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqpnm.KeyPress
 If e.KeyChar = Chr(13) Then
    txtShoe.Focus()
 End If
End Sub

Private Sub txtEqpnm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpnm.LostFocus
  txtEqpnm.Text = txtEqpnm.Text.ToUpper.Trim
End Sub

Private Sub txtShoe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShoe.KeyDown
 Dim intChkPoint As Integer
        With txtShoe
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtEqpnm.Focus()
                    End If
                Case Is = 38 '�����١�â��        
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtOrder.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtOrder.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     cboPart.DroppedDown = True
                     cboPart.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtShoe_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShoe.KeyPress
 If e.KeyChar = Chr(13) Then
    txtOrder.Focus()
 End If
End Sub

Private Sub txtShoe_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShoe.LostFocus
 txtShoe.Text = txtShoe.Text.ToUpper.Trim
End Sub

Private Sub txtOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrder.KeyDown
 Dim intChkPoint As Integer
        With txtOrder
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtShoe.Focus()
                    End If
                Case Is = 38 '�����١�â��        
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                       txtRemark.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtRemark.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                      txtRemark.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrder.KeyPress
 If e.KeyChar = Chr(13) Then
   txtRemark.Focus()
 End If
End Sub

Private Sub txtOrder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrder.LostFocus
  txtOrder.Text = txtOrder.Text.ToUpper.Trim
End Sub

Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
  Dim intChkPoint As Integer
        With txtOrder
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                      txtOrder.Focus()
                    End If
                Case Is = 38 '�����١�â��      
                       txtOrder.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                       txtCdate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                         txtCdate.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtCdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

End Sub

Private Sub txtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.LostFocus
  txtRemark.Text = txtRemark.Text.ToUpper.Trim
End Sub

Private Sub cboPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboPart.KeyPress
  If e.KeyChar = Chr(13) Then
     txtSize.Focus()
  End If
End Sub

Private Sub txtCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCdate.GotFocus
 With mskCdate
      .BringToFront()
      txtCdate.SendToBack()
      .Focus()
 End With
End Sub

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown
  Dim intChkPoint As Integer

        With txtSize
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

                    If .SelectionStart = 0 Then
                        cboPart.DroppedDown = True
                    End If

                Case Is = 38 '�����١�â��    
                     txtCdate.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        txtSizeDesc.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtSizeDesc.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtSetQty.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress
  If e.KeyChar = Chr(13) Then
     txtSizeDesc.Focus()
  End If
End Sub

Private Sub txtSize_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSize.LostFocus
 txtSize.Text = txtSize.Text.ToUpper.Trim
End Sub

Private Sub txtSizeDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSizeDesc.KeyDown
 Dim intChkPoint As Integer

        With txtSizeDesc

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

                    If .SelectionStart = 0 Then
                        txtSize.Focus()
                    End If

                Case Is = 38 '�����١�â��    
                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        txtSetQty.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtSetQty.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtPrice.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSizeDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSizeDesc.KeyPress
 If e.KeyChar = Chr(13) Then
    txtSetQty.Focus()
 End If
End Sub

Private Sub txtSizeDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeDesc.LostFocus
 txtSizeDesc.Text = txtSizeDesc.Text.ToUpper.Trim
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

                            If InStr("0123456789.", strTmp) > 0 Then
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

Private Sub mskSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSetQty.KeyPress

  Select Case Asc(e.KeyChar)

         Case 48 To 57            ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
         Case 13
                e.Handled = False
                txtSizeQty.Focus()
         Case 8                   ' ���� Backspace
                e.Handled = False
         Case 32                   '��� spacebar
                e.Handled = False
         Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSetQty.Focus()

  End Select

End Sub

Private Sub mskSetQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSetQty.LostFocus
 Dim i, x As Integer
 Dim z As Double
 Dim strTmp As String = ""
 Dim strMerge As String = ""

        With mskSetQty

            x = .Text.Length

            For i = 1 To x

                strTmp = Mid(.Text.ToString, i, 1)
                Select Case strTmp
                    Case Is = ","
                    Case Is = "+"
                    Case Is = "_"
                    Case Else

                        If InStr("0123456789.", strTmp) > 0 Then
                            strMerge = strMerge & strTmp
                        End If

                End Select
                strTmp = ""
            Next i
            Try

                mskSetQty.Text = ""
                z = CDbl(strMerge)
                txtSetQty.Text = z.ToString("#,##0.0")


            Catch ex As Exception
                txtSetQty.Text = "0.0"
                mskSetQty.Text = ""
            End Try

            mskSetQty.SendToBack()
            txtSetQty.BringToFront()

        End With
End Sub

Private Sub mskSetQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSetQty.KeyDown
  Dim intChkPoint As Integer

     With mskSetQty

            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtSizeDesc.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtSize.Focus()

                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSizeQty.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     '��� InChkPoint = ������Ǣͧ  mskLdate

                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSizeQty.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ
                    txtRmk.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
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

Private Sub mskPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskPrice.KeyDown
 Dim intChkPoint As Integer

     With mskPrice

            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtSizeQty.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtSize.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtRmk.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtRmk.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ
                    txtRmk.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

      End With
End Sub

Private Sub mskPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrice.KeyPress

 Select Case Asc(e.KeyChar)

            Case 48 To 57            ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13                  '���� Enter
                e.Handled = False
                txtRmk.Focus()
            Case 8                   ' ���� Backspace
                e.Handled = False
            Case 32                   ' ���� Tab
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPrice.Focus()
  End Select

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

Private Sub txtRmk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRmk.KeyDown
  Dim intChkPoint As Integer

      With txtRmk
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtPrice.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtSetQty.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        btnSeekSave.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           btnSeekSave.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub txtRmk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRmk.KeyPress
  If e.KeyChar = Chr(13) Then
     btnSeekSave.Focus()
  End If
End Sub

Private Sub txtPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
 If e.KeyChar = Chr(13) Then
    txtSize.Focus()
 End If
End Sub

Private Sub btnDelEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
End Sub

Private Sub mskSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerge As String = ""

     With mskSizeQty

        If txtSizeQty.Text.ToString.Trim <> "0" Then
           x = Len(txtSizeQty.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtSizeQty.Text.ToString, i, 1)

                                Select Case strTmp

                                          Case Is = "_"
                                          Case Else

                                                    If InStr("0123456789.", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select

                         Next i


                 Select Case strMerge.IndexOf(".")

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

                        .SelectedText = strMerge

                End If

        .SelectAll()

End With

End Sub

Private Sub mskSizeQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSizeQty.KeyDown
 Dim intChkPoint As Integer

        With mskSizeQty

            Select Case e.KeyCode

                   Case Is = 35 '���� End 
                   Case Is = 36 '���� Home
                   Case Is = 37 '�١�ë���

                         If .SelectionStart = 0 Then
                             mskSetQty.Focus()
                         End If

                   Case Is = 38 '�����١�â��    
                          cboPart.DroppedDown = True

                   Case Is = 39 '�����١�â��

                        If .SelectionLength = .Text.Trim.Length Then
                           txtPrice.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length
                            If .SelectionStart = intChkPoint Then
                               txtPrice.Focus()
                            End If

                        End If

                   Case Is = 40 '����ŧ    
                        txtRmk.Focus()
                   Case Is = 113 '���� F2
                        .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress

      Select Case Asc(e.KeyChar)

            Case 48 To 57            ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13                  '���� Enter
                e.Handled = False
                txtPrice.Focus()
            Case 8                   ' ���� Backspace
                e.Handled = False
            Case 32                   ' ���� Tab
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                mskSizeQty.Focus()
      End Select

End Sub

Private Sub txtSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeQty.GotFocus
  With mskSizeQty
       .BringToFront()
       txtSizeQty.SendToBack()
       .Focus()
  End With
End Sub

Private Sub mskSizeQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.LostFocus

 Dim i, x As Integer
 Dim z As Double

 Dim strTmp As String = ""
 Dim strMerge As String = ""
 Dim intFull As Integer

      With mskSizeQty

           x = .Text.Length

                For i = 1 To x
                    strTmp = Mid(.Text.ToString, i, 1)

                    Select Case strTmp
                           Case Is = ","
                           Case Is = "+"
                           Case Is = "_"
                           Case Else

                                If InStr("0123456789.", strTmp) > 0 Then
                                   strMerge = strMerge & strTmp
                                End If

                    End Select
                    strTmp = ""

               Next i

         Try

              mskSizeQty.Text = ""
              z = CDbl(strMerge)

              intFull = Int(z)

              If (z - intFull) > 0 Then
                  txtSizeQty.Text = z.ToString("#,##0.0")
              Else
                  txtSizeQty.Text = z.ToString("0.0")
              End If

         Catch ex As Exception
               txtSizeQty.Text = "0.0"
               mskSizeQty.Text = ""
         End Try

          mskSizeQty.SendToBack()
          txtSizeQty.BringToFront()

       End With

End Sub

Private Sub mskCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskCdate.GotFocus

 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerg As String = ""

     With mskCdate

           If txtCdate.Text.Trim <> "__/__/____" Then
                x = Len(txtCdate.Text)

                For i = 1 To x

                    strTmp = Mid(txtCdate.Text.Trim, i, 1)
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

Private Sub mskCdate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskCdate.KeyDown

    Dim intChkPoint As Integer
        With mskCdate

            Select Case e.KeyCode
                   Case Is = 35 '���� End 
                   Case Is = 36 '���� Home
                   Case Is = 37 '�١�ë���
                   Case Is = 38 '�١�â��
                   Case Is = 39   '�����١�â��
                        If .SelectionLength = .Text.Trim.Length Then
                           cboPart.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length
                            If .SelectionStart = intChkPoint Then
                               cboPart.Focus()
                            End If
                        End If

                  Case Is = 40 '����ŧ
                       cboPart.Focus()
                  Case Is = 113 '���� F2
                       .SelectionStart = .Text.Trim.Length
            End Select

        End With

End Sub

Private Sub mskCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskCdate.KeyPress
    If e.KeyChar = Chr(13) Then
       cboPart.Focus()
    End If
End Sub

Private Sub mskCdate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskCdate.LostFocus

  Dim i, x As Integer
  Dim z As Date

  Dim strTmp As String = ""
  Dim strMerge As String = ""

        With mskCdate
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

                mskCdate.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)
                txtCdate.Text = z.ToString("dd/MM/yyyy")

                'If Year(z) < 2500 Then  '�óա�͡�� �.�. ����ŧ�� �.�. �ѹ��
                '    txtRecvDate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))
                'Else
                '    txtRecvDate.Text = z.ToString("dd/MM/yyyy")
                'End If

            Catch ex As Exception
                  mskCdate.Text = "__/__/____"
                  txtCdate.Text = "__/__/____"
            End Try

          mskCdate.SendToBack()
          txtCdate.BringToFront()

        End With

End Sub

End Class