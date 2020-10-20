Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmPastMold
    Dim IsShowSeek As Boolean        '�������ʴ�ʶҹ� gpbSeek
    Dim strDateDefault As String     '���������Ѻ�ѹ�������

    Public Const DrvName As String = "\\10.32.0.15\data1\EquipPicture\"
    Public Const PthName As String = "\\10.32.0.15\data1\EquipPicture"         '���������Ѻ�� part �ٻ�Ҿ

   Private tt As ToolTip = New ToolTip '�ʴ����ŷԻ ��ٻ�Ҿ��������͹��������

Protected Overrides ReadOnly Property CreateParams() As CreateParams       '��ͧ�ѹ��ûԴ������� Close Button(�����ҡ�ҷ)
  Get
       Dim cp As CreateParams = MyBase.CreateParams
           Const CS_DBLCLKS As Int32 = &H8
           Const CS_NOCLOSE As Int32 = &H200
           cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
           Return cp
       End Get
End Property

Private Sub frmPastMold_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTable(0, "")  'ź������ Table tmp_eqptrn where user_id..
   ClearTmpTable(2, "") '��ҧ�����ŵ��ҧ tmp_fixeqptrn
   ClearTmpTable(3, "") '��ҧ�����ŵ��ҧ tmp_eqptrn_newsize
   frmEqpSheet.lblCmd.Text = "0"  '������ʶҹ�
   Me.Dispose()     '����¿���� �׹˹��¤�����
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

               Case Is = "2"
                    strSqlcmd = "DELETE tmp_fixeqptrn " _
                               & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'"
                    .Execute(strSqlcmd)

               Case Is = "3"
                    strSqlcmd = "DELETE tmp_eqptrn_newsize " _
                               & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'"
                    .Execute(strSqlcmd)

         End Select

     End With

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub frmPastMold_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

 Dim dtComputer As Date = Now       '������纤���ѹ���Ѩ�غѹ
 Dim strCurrentDate As String       '�纤��ʵ�ԧ�ѹ���Ѩ�غѹ

     StdDateTimeThai()        '�Ѻ�ٷչ��ŧ�ѹ�����Ǵ�.�� ����� Module
     strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

     ClearDataGpbHead()
     PrePartSeek()            '��Ŵ��������´���� cbo ������ǹ����Ե
     PreMoldStatus()            'ʶҹ�����

     Select Case frmEqpSheet.lblCmd.Text.ToString

            Case Is = "0" '����������

                With txtBegin
                     .Text = strCurrentDate
                     strDateDefault = strCurrentDate
                End With

                With Me
                     .Text = "����������"
                     txtEqp_id.Focus()
                End With

                '--------------------- ������������������ͧ�ʴ�ʶҹ�(��͹�������� Gridview) ----------------------------

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
                cboPart.Enabled = False

            Case Is = "2"   '����ͧ������

                With Me
                     .Text = "����ͧ������"
                End With

                LockEditData()
                txtEqp_id.ReadOnly = True  '�����ҹ���ҧ����
                cboPart.Enabled = False
                btnSaveData.Enabled = False

        End Select

    txtEqp_id.Focus()
End Sub

Private Sub ClearDataGpbHead()
     txtEqp_id.Text = ""
     txtEqpnm.Text = ""
     txtShoe.Text = ""
     txtOrder.Text = ""
     cboPart.Text = ""
     txtAmount.Text = ""
     txtSet.Text = ""
End Sub

Private Sub LockEditData()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim Rsdwc As New ADODB.Recordset

 Dim strCmd As String              '��ʵ�ԧ Command

 Dim strLoadFilePicture As String   '�纤��ʵ�ԧ��Ŵ Picture
 Dim strPathPicture As String = "\\10.32.0.15\data1\EquipPicture\"   '�� part

 Dim blnHavedata As Boolean       '�纤�ҵ����� ����Ѻ������բ������������
 Dim strSqlSelc As String = ""    '��ʵ�ԧ sql select
 Dim strPart As String = ""

 Dim strCode As String = frmEqpSheet.dgvShoe.Rows(frmEqpSheet.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim

     With Conn
         If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

    End With
        ' NOLOCK ���¶֧����ö�ӧҹ�Ѻ Table �����١��ҡ user ���
        strSqlSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                    & " WHERE eqp_id = '" & strCode & "'"

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , )

         If .RecordCount <> 0 Then

                cboPart.Text = .Fields("part").Value.ToString.Trim
                txtBegin.Text = .Fields("creat_date").Value.ToString.Trim
                strDateDefault = .Fields("creat_date").Value.ToString.Trim

                txtEqp_id.Text = .Fields("eqp_id").Value.ToString.Trim
                txtEqpnm.Text = .Fields("eqp_name").Value.ToString.Trim
                txtShoe.Text = .Fields("shoe").Value.ToString.Trim
                txtAmount.Text = .Fields("pi_qty").Value.ToString.Trim
                txtSet.Text = Format(.Fields("set_qty").Value, "#.##0.0")
                txtRemark.Text = .Fields("remark").Value.ToString.Trim

                lblPicName1.Text = .Fields("pic_ctain").Value.ToString.Trim
                lblPicName2.Text = .Fields("pic_io").Value.ToString.Trim
                lblPicName3.Text = .Fields("pic_part").Value.ToString.Trim
                lblPicPath1.Text = PthName
                lblPicPath2.Text = PthName
                lblPicPath3.Text = PthName

                '-----------------------------------����鹧ҹ------------------------------

                Select Case .Fields("part").Value.ToString.Trim
                    Case Is = "SOLE2"
                        cboPart.Text = "�����ҧ"
                    Case Is = "SOLE1"
                        cboPart.Text = "��鹺�"
                    Case Is = "FILL"
                       cboPart.Text = "����鹺�"
                    Case Is = "LOGOHEEL"
                        cboPart.Text = "������"
                    Case Is = "LOGO"
                        cboPart.Text = "������"
                    Case Is = "EVAH"
                        cboPart.Text = "EVA �Դ��"
                    Case Is = "EVAF"
                        cboPart.Text = "EVA ��˹ѧ˹��"
                    Case Is = "UPPER"
                       cboPart.Text = "˹ѧ˹��"
                    Case Is = "ONUPPER"
                       cboPart.Text = "ONUPPER"
                    Case Is = "ONUPPER"
                       cboPart.Text = "ONUPPER"
                End Select

                '-------------------------------Load �ٻ�Ҿ(�ٻẺἧ)-----------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_ctain").Value.ToString.Trim
                ' �紾�������������ԧ
                If File.Exists(strLoadFilePicture) Then

                    Dim img1 As Image      '��С�ȵ���� img1 �������Ҿ
                    img1 = Image.FromFile(strLoadFilePicture)  'img1 ��ҡѺpicture �����Ŵ�Ҩҡ db
                    picEqp1.Image = ScaleImage(img1, picEqp1.Height, picEqp1.Width)

                Else
                    picEqp1.Image = Nothing  '���������ٻ�Ҿ��� picEqp1 ��ҧ����
                End If

                '-------------------------------Load �ٻ�Ҿ�������鹧ҹ ----------------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_io").Value.ToString.Trim
                If File.Exists(strLoadFilePicture) Then
                    Dim img2 As Image
                    img2 = Image.FromFile(strLoadFilePicture)
                     picEqp2.Image = ScaleImage(img2, picEqp2.Height, picEqp2.Width)

                Else
                    picEqp2.Image = Nothing
                End If


                 '-------------------------------Load �ٻ�Ҿ�ͧ���������ٻ  ----------------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_part").Value.ToString.Trim
                If File.Exists(strLoadFilePicture) Then
                    Dim img3 As Image
                    img3 = Image.FromFile(strLoadFilePicture)
                    picEqp3.Image = ScaleImage(img3, picEqp3.Height, picEqp3.Width)
                Else
                    picEqp3.Image = Nothing
                End If


                strCmd = frmEqpSheet.lblCmd.Text.ToString.Trim    '��� strCmd ��ҡѺ���� lblcmd 㹿���� frmEqpSheet

                Select Case strCmd
                       Case Is = "1"   '�����ͤ�͹���
                       Case Is = "2"   '�����ͤ�͹����ͧ
                            btnSaveData.Enabled = False  '�Դ���� "�ѹ�֡������"
                End Select

                '------------------------------- ��ҹ�����š���觫������㹵��ҧ tmp_eqptrn --------------------

                 strSqlSelc = "INSERT INTO tmp_fixeqptrn " _
                                  & "SELECT user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE eqp_id = '" & strCode & "'" _
                                  & " AND fix_sta= '" & "1" & "'"

                Conn.Execute(strSqlSelc)

                '------------------------------- �ѹ�֡�����ŧ㹵��ҧ tmp_eqptrn----------------------------

                strSqlSelc = "INSERT INTO tmp_eqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM eqptrn " _
                                  & " WHERE eqp_id = '" & strCode & "' "

                Conn.Execute(strSqlSelc)

                '------------------------------ ���§ SIZE ���� -------------------------------------------

                ReSizeSort(strCode)   '�Ѵ���§ size ��������

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

Private Sub ShowScrapItem()  'ǹ�ٻ��������ʴ���� �ӹǹ SET, ������ŧ��Ե, ����Ҥ��ػ�ó�

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmdSelc As String    '�纤�� string command

 Dim dubQty As Double
 Dim dubAmt As Double
 Dim sngSetQty As Single     '�纨ӹǹ SET
 Dim user As String = frmMainPro.lblLogin.Text.ToString.Trim
 Dim mold_id As String
 Dim mold_size As String
 Dim strArr() As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
    End With

         strSqlCmdSelc = "SELECT * FROM v_tmpeqptrn_newsize (NOLOCK)" _
                                 & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "' " _
                                 & "ORDER BY tmp_newsize"

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlCmdSelc, Conn, , , )

            dgvSize.Rows.Clear()
            dgvSize.ScrollBars = ScrollBars.None '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ

            If .RecordCount <> 0 Then

                Do While Not .EOF()

                   mold_id = .Fields("eqp_id").Value.ToString.Trim
                   strArr = Split(.Fields("size_desc").Value.ToString.Trim, "-")  '�Ѵ array �͡��
                   mold_size = .Fields("size_id").Value.ToString.Trim + strArr(0)

                   dgvSize.Rows.Add( _
                                        IIf(.Fields("delvr_sta").Value.ToString.Trim = "1", My.Resources.accept, My.Resources._16x16_ledred), _
                                        IIf(Find_fixmold(user, mold_id, mold_size) = "1", My.Resources.accept, My.Resources.blank), _
                                        .Fields("size_id").Value.ToString.Trim, _
                                        .Fields("size_act").Value.ToString.Trim, _
                                        .Fields("size_desc").Value.ToString.Trim, _
                                        .Fields("size_group").Value.ToString.Trim, _
                                        Format(.Fields("set_qty").Value, "##0.0"), _
                                        Format(.Fields("size_qty").Value, "##0.0"), _
                                        .Fields("dimns").Value.ToString.Trim, _
                                        .Fields("price").Value, _
                                         Mid(.Fields("pr_date").Value.ToString.Trim, 1, 10), _
                                        .Fields("pr_doc").Value.ToString.Trim, _
                                        .Fields("sup_name").Value.ToString.Trim, _
                                         Mid(.Fields("fc_date").Value.ToString.Trim, 1, 10), _
                                         Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
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

Private Sub PrePartSeek()    '��Ŵ��������´���� cbo ������ǹ����Ե
   Dim strGpTopic(8) As String
   Dim i As Byte

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
            Next i

        End With

End Sub

Private Sub CheckDataBfSave()

  Dim intListwc As Integer = dgvSize.Rows.Count  '����� inListwc �纨ӹǹ�ä���촢ͧ dgvSize
  Dim strProd As String = ""
  Dim strProdnm As String = ""

  Dim bytConSave As Byte  '�纤�� megbox 

        If txtEqp_id.Text.ToString.Trim <> "" Then

            If txtEqpnm.Text.ToString.Trim <> "" Then

                If cboPart.Text.ToString.Trim <> "" Then

                    If intListwc > 0 Then  '��� dgvSize �բ�����

                        '��� MsgBox = ���� Yes 
                           bytConSave = MsgBox("�س��ͧ��úѹ�֡���������������!" _
                                  , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")


                               If bytConSave = 6 Then

                                        Select Case Me.Text
                                               Case Is = "����������"

                                                    If CheckCodeDuplicate() Then
                                                       SaveNewRecord()

                                                    Else
                                                       MessageBox.Show("��سҡ�͡�����ػ�ó�����!....", "***�����ػ�ó���***", _
                                                                                     MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                       txtEqp_id.Text = ""
                                                       txtEqp_id.Enabled = True
                                                       txtEqp_id.ReadOnly = False
                                                       txtEqp_id.Focus()

                                                    End If

                                                Case Else
                                                    SaveEditRecord()

                                        End Select

                               Else
                                     dgvSize.Focus()
                               End If

                    Else

                        If CheckCodeDuplicate() Then
                             ShowResvrd()       '�ʴ���������� gpbSeek 
                             gpbSeek.Text = "����������"
                             txtSize.ReadOnly = False

                        Else

                             MessageBox.Show("��سҡ�͡�����ػ�ó�����!....", "�����ػ�ó���", MessageBoxButtons.OK, MessageBoxIcon.Error)
                             txtEqp_id.Text = ""
                             txtEqp_id.ReadOnly = False
                             txtEqp_id.Enabled = True
                             txtEqp_id.Focus()

                        End If


                    End If

                Else
                    MsgBox("�ô�кآ�������������´�ػ�ó�  " _
                                & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                    cboPart.DroppedDown = True
                    cboPart.Focus()

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

Private Sub SaveEditRecord()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

  Dim dateSave As Date = Now()
  Dim strDate As String = ""

  Dim strCredate As String
  Dim strDocdate As String           '��ʵ�ԧ�ѹ����͡���
  Dim strGpType As String = ""       '�纻������ػ�ó�
  Dim strPartType As String = ""     '�纪����ǹ����Ե

  Dim blnReturnCopyPic As Boolean

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

        If txtCdate.Text <> "__/__/____" Then

            strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtCdate.Text.ToString, 1, 2)
            strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

        Else
            strCredate = "NULL"

        End If

        '------------------------------------ ��˹�������ͧ��鹧ҹ -----------------------------------------

             Select Case cboPart.Text.ToString
                    Case Is = "�����ҧ"
                        strPartType = "SOLE2"
                    Case Is = "��鹺�"
                        strPartType = "SOLE1"
                    Case Is = "����鹺�"
                        strPartType = "FILL"
                    Case Is = "������"
                        strPartType = "LOGOHEEL"
                    Case Is = "������"
                        strPartType = "LOGO"
                    Case Is = "EVA �Դ��"
                        strPartType = "EVAH"
                    Case Is = "EVA ��˹ѧ˹��"
                        strPartType = "EVAF"
                    Case Is = "˹ѧ˹��"
                        strPartType = "UPPER"
                    Case Is = "ONUPPER"
                        strPartType = "ONUPPER"

            End Select


                      '------------------------------------ �ѹ�֡�ٻἧ ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath1.Text.ToString.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath1.Text = PthName

                       Else
                          lblPicName1.Text = ""
                          lblPicPath1.Text = ""
                          picEqp1.Image = Nothing

                       End If


                      '------------------------------------ �ѹ�֡�ٻ���� ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath2.Text.ToString.Trim, ReturnImageName(lblPicName2.Text.ToString.Trim), lblPicName2.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath2.Text = PthName

                       Else
                          lblPicName2.Text = ""
                          lblPicPath2.Text = ""
                          picEqp2.Image = Nothing

                       End If

                      '------------------------------------ �ѹ�֡�ٻ PIC_PART------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath3.Text.ToString.Trim, ReturnImageName(lblPicName3.Text.ToString.Trim), lblPicName3.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath3.Text = PthName

                       Else
                          lblPicName3.Text = ""
                          lblPicPath3.Text = ""
                          picEqp3.Image = Nothing

                       End If

                      '---------------------------- UPDATE ������㹵��ҧ eqpmst ------------------------------

                      strSqlCmd = "UPDATE  eqpmst SET eqp_name ='" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                 & "," & "pi ='" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                                 & "," & "shoe ='" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                                 & "," & "part ='" & strPartType & "'" _
                                 & "," & "eqp_type ='" & "LCA" & "'" _
                                 & "," & "set_qty =" & ChangFormat(txtSet.Text.ToString.Trim) _
                                 & "," & "pic_ctain ='" & ReplaceQuote(lblPicName1.Text.ToString.Trim) & "'" _
                                 & "," & "pic_io ='" & ReplaceQuote(lblPicName2.Text.ToString.Trim) & "'" _
                                 & "," & "pic_part ='" & ReplaceQuote(lblPicName3.Text.ToString.Trim) & "'" _
                                 & "," & "remark ='" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                 & "," & "tech_desc = '" & ReplaceQuote(txtTdesc.Text.ToString.Trim) & "'" _
                                 & "," & "tech_thk = '" & ReplaceQuote(txtThk.Text.ToString.Trim) & "'" _
                                 & "," & "tech_lg = '" & ReplaceQuote(txtTtrait.Text.ToString.Trim) & "'" _
                                 & "," & "tech_sht = '" & ReplaceQuote(txtTsht.Text.ToString.Trim) & "'" _
                                 & "," & "tech_eva = '" & ReplaceQuote(txtTeva.Text.ToString.Trim) & "'" _
                                 & "," & "tech_warm = '" & ReplaceQuote(txtTwarm.Text.ToString.Trim) & "'" _
                                 & "," & "tech_time1 = '" & ReplaceQuote(txtTtime1.Text.ToString.Trim) & "'" _
                                 & "," & "tech_time2 = '" & ReplaceQuote(txtTtime2.Text.ToString.Trim) & "'" _
                                 & "," & "creat_date = " & strCredate _
                                 & "," & "eqp_amt = " & RetrnAmount() _
                                 & "," & "last_date = '" & strDate & "'" _
                                 & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                 & "," & "exp_id ='" & "" & "'" _
                                 & "," & "tech_trait ='" & ReplaceQuote(txtTtrait.Text.ToString.Trim) & "'" _
                                 & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)


        '------------------------------------------------ź������㹵��ҧ eqptrn--------------------------------------------------------------------

                     strSqlCmd = "Delete FROM eqptrn" _
                                          & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlCmd)

        '-----------------------------------------�ѹ�֡������㹵��ҧ eqptrn �� Select �ҡ tmp_eqptrn-------------------------------------------------

        strSqlCmd = "INSERT INTO eqptrn " _
                        & "SELECT [group] = 'D'" _
                        & ",eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'" _
                        & ",size_id,size_desc,size_qty,weight,dimns,backgup" _
                        & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date" _
                        & ",pr_doc,recv_date,ord_rep,ord_qty,fc_date,impt_id" _
                        & ",sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                        & " FROM tmp_eqptrn " _
                        & " WHERE user_id= '" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

        Conn.Execute(strSqlCmd)
        Conn.CommitTrans()  '��� Commit transection

        frmEqpSheet.lblCmd.Text = txtEqp_id.Text.ToString.Trim
        frmEqpSheet.Activating()
        Me.Close()

        Conn.Close()
        Conn = Nothing

    End Sub

    Private Sub SaveNewRecord()

        Dim Conn As New ADODB.Connection
        Dim strSqlCmd As String
        Dim dateSave As Date = Now()    '�纤���ѹ���Ѩ�غѹ
        Dim strDate As String

        Dim blnRetuneCopyPic As Boolean
        Dim strCredate As String
        Dim strPRdate As String
        Dim strDateDoc, strINdate As String
        Dim strFCdate As String
        Dim strPartType As String = ""
        Dim strType As String = ""
        Dim Rsd As New ADODB.Recordset

        With Conn
            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()

        End With

            Conn.BeginTrans()


                    strDate = dateSave.Date.ToString("yyyy-MM-dd")
                    strDate = SaveChangeEngYear(strDate)               '��ŧ�繻� �.�.

                    strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                   & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                   & Mid(txtBegin.Text.Trim, 1, 2)
                    strDateDoc = SaveChangeEngYear(strDateDoc)

                    '---------------------------------------- �ѹ�֡�ٻẺἧ ----------------------------------------------------

                    blnRetuneCopyPic = CallCopyPicture(lblPicPath1.Text.Trim, ReturnImageName(lblPicName1.Text.Trim), lblPicName1.Text.Trim)

                    If blnRetuneCopyPic Then       '��� CallCopyPicture = true
                       lblPicPath1.Text = PthName
                    Else
                       lblPicPath1.Text = ""
                       lblPicName1.Text = ""
                       picEqp1 = Nothing

                    End If

                    '---------------------------------------- �ѹ�֡�ٻ�������鹧ҹ  ----------------------------------------------------
                     blnRetuneCopyPic = CallCopyPicture(lblPicPath2.Text.Trim, ReturnImageName(lblPicName2.Text.Trim), lblPicName2.Text.Trim)

                    If blnRetuneCopyPic Then
                       lblPicPath2.Text = PthName

                    Else
                       lblPicPath2.Text = ""
                       lblPicName2.Text = ""
                       picEqp2 = Nothing
                    End If


                    '---------------------------------------- �ѹ�֡�ٻ�ͧ���������ٻ ----------------------------------------------------
                     blnRetuneCopyPic = CallCopyPicture(lblPicPath3.Text.Trim, ReturnImageName(lblPicName3.Text.Trim), lblPicName3.Text.Trim)

                    If blnRetuneCopyPic Then
                       lblPicPath3.Text = PthName

                    Else
                       lblPicPath3.Text = ""
                       lblPicName3.Text = ""
                       picEqp3 = Nothing
                    End If


                    '---------------------------------------- Ǵ�.����Դ ----------------------------------------------------
                    If txtCdate.Text <> "__/__/____" Then

                       strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtCdate.Text.ToString, 1, 2)
                       strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

                   Else
                       strCredate = "NULL"
                   End If


                    '---------------------------------------- Ǵ�.�Դ���觫��� -------------------------------------------------
                   If txtPrdate.Text <> "__/__/____" Then

                      strPRdate = Mid(txtPrdate.Text.ToString, 7, 4) & "-" _
                                    & Mid(txtPrdate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtPrdate.Text.ToString, 1, 2)
                      strPRdate = "'" & SaveChangeEngYear(strPRdate) & "'"

                   Else
                      strPRdate = "NULL"
                   End If


                    '---------------------------------------- �ѹ�չѴ��� -------------------------------------------------
                   If txtFCdate.Text <> "__/__/____" Then

                       strFCdate = Mid(txtFCdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtFCdate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtFCdate.Text.ToString, 1, 2)
                       strFCdate = "'" & SaveChangeEngYear(strFCdate) & "'"

                   Else
                       strFCdate = "NULL"
                   End If


                    '---------------------------------------- Ǵ�.����Ѻ��� -------------------------------------------------
                  If txtIndate.Text <> "__/__/____" Then

                     strINdate = Mid(txtIndate.Text.ToString, 7, 4) & "-" _
                                    & Mid(txtIndate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtIndate.Text.ToString, 1, 2)
                     strINdate = "'" & SaveChangeEngYear(strINdate) & "'"

                 Else
                     strINdate = "NULL"
                 End If

                    '------------------------------------��˹�������ͧ��鹧ҹ--------------------------------------------------

                    Select Case cboPart.Text.ToString.Trim

                        Case Is = "�����ҧ"
                            strPartType = "SOLE2"
                        Case Is = "��鹺�"
                            strPartType = "SOLE1"
                        Case Is = "����鹺�"
                            strPartType = "FILL"
                        Case Is = "������"
                            strPartType = "LOGOHEEL"
                        Case Is = "������"
                            strPartType = "LOGO"
                        Case Is = "EVA �Դ��"
                            strPartType = "EVAH"
                        Case Is = "EVA ��˹ѧ˹��"
                            strPartType = "EVAF"
                        Case Is = "˹ѧ˹��"
                            strPartType = "UPPER"
                        Case Is = "ONUPPER"
                            strPartType = "ONUPPER"

                    End Select


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
                      & ",'" & "D" & "'" _
                      & ",'" & ReplaceQuote(txtEqp_id.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtEqpnm.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(txtSupplier.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtRef.Text.ToString.Trim) & "'" _
                      & ",'" & ChangFormat(txtSet.Text.ToString.Trim) & "'" _
                      & ",'" & strPartType.ToString.Trim & "'" _
                      & ",'" & "LCA" & "'" _
                      & ",'" & ReplaceQuote(lblPicName1.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(lblPicName2.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(lblPicName3.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTdesc.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtThk.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTtrait.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTsht.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTeva.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTwarm.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTtime1.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTtime2.Text.ToString.Trim) & "'" _
                      & "," & strCredate _
                      & ",'" & strDate & "'" _
                      & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                      & ",'" & ChangFormat(txtAmount.Text.ToString.Trim) & "'" _
                      & ",'" & RetrnAmount() & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(txtTtrait.Text.ToString.Trim) & "'" _
                      & ")"

                Conn.Execute(strSqlCmd)


                   '------------------------------------------------�ѹ�֡������㹵��ҧ eqptrn----------------------------------------------------------

                strSqlCmd = "INSERT INTO eqptrn " _
                                     & " SELECT [group] ='D'" _
                                     & ",eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'" _
                                     & ",size_id,size_desc,size_qty,weight,dimns,backgup,price,men_rmk" _
                                     & ",delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                     & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & " FROM tmp_eqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                Conn.Execute(strSqlCmd)
                Conn.CommitTrans()

                frmEqpSheet.lblCmd.Text = txtEqp_id.Text.ToString.Trim
                frmEqpSheet.Activating()
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

    '�ѧ���� CopyPicture
Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String, ByVal strNewPicName As String) As Boolean

  Dim fname As String = String.Empty  '��ҡѺ ""
  Dim dFile As String = String.Empty
  Dim dFilePath As String = String.Empty

  Dim fServer As String = String.Empty
  Dim intResult As Integer   '�׹����繨ӹǹ���

  On Error GoTo Err70

        fname = strPicPath & "\" & strPicName  '���� \\10.32.0.15\data1\EquipPicture\"�����ٻ�Ҿ"
        fServer = PthName & "\" & strNewPicName    'partServer \\10.32.0.15\data1\EquipPicture\"�����ٻ�Ҿ"

        If File.Exists(fServer) Then    '�������������ԧ
           CallCopyPicture = True      '���׹��� true

        Else

            If File.Exists(fname) Then
               dFile = Path.GetFileName(fname)
               dFilePath = DrvName + dFile

               intResult = String.Compare(fname.ToString.Trim, dFilePath.ToString.Trim)

                '------------------------------------��Ҥ���� 0 �ʴ������Ŵ��������� �������ö Copy �����------------------------------

                    If intResult = 1 Then '��ҷ���� = 1 �֧ copy �ٻ�����������ͧ 10.32.0.14
                       File.Copy(fname, dFilePath, True)
                    End If

                    My.Computer.FileSystem.RenameFile(dFilePath, strNewPicName)  '����¹��������ٻ����
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

    Private Function CheckCodeDuplicate() As Boolean     '�ѧ���������ʫ��

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

        strSqlSelc = "SELECT eqp_id FROM eqpmst(NOLOCK)" _
                                  & " WHERE eqp_id = '" & txtEqp_id.Text.Trim & "'"

        Rsd = New ADODB.Recordset

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

    Private Sub ClearSubData1()
        txtTdesc.Text = ""
        txtThk.Text = "0.0"
        txtTtrait.Text = ""
        txtTsht.Text = ""
        txtTeva.Text = ""
        txtTwarm.Text = "0"
        txtTtime1.Text = "0"
        txtTtime2.Text = "0"
        txtCdate.Text = "__/__/____"

    End Sub

    Private Sub ClearSubData2()
        txtPId.Text = ""
        txtSize.Text = ""
        txtSizeDesc.Text = ""
        txtSetQty.Text = "0"
        txtSizeQty.Text = "0"
        txtPr.Text = ""
        txtw.Text = "0.00"
        txtLg.Text = "0.00"
        txtPrice.Text = "0.00"
        'txtPrdate.Text = "__/__/____"
        'txtFCdate.Text = "__/__/____"
        'txtIndate.Text = "__/__/____"
        txtSupplier.Text = ""
        'txtRmk.Text = ""
    End Sub

    Private Sub SaveCancle()
        btnSaveData.Enabled = True
        btnExit.Enabled = False
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

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBfSave()
End Sub

Private Sub ShowResvrd()  '����ʴ� GroupBox gpbSeek �����

  tabMain.SelectedTab = tabSize
  IsShowSeek = Not IsShowSeek  '���ʶҹ�ᶺ seek ����ʴ� ����ʴ�

  If IsShowSeek Then

     With gpbSeek
          .Visible = True
          .Left = 10    '᡹ X
          .Top = 215   '᡹ Y 252
          .Height = 504
          .Width = 1006
    End With

    StateLockFindDept(False) ' ��ͤ FindDept ���觤���� False �
    txtTdesc.Focus()

 Else
      StateLockFindDept(True)
 End If

End Sub

    Private Sub StateLockFindDept(ByVal Sta As Boolean)  '��ͤ gpbHead ������ enabled = false
        Dim strMode As String = frmEqpSheet.lblCmd.Text.ToString   '����� lblCmd � frmEqpSheet �觤������� strMode 

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

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        ShowResvrd()
        ClearSubData2()
        CallEditData()    '�Ѻ�ٷչ�ʴ� Size ������䢢�����
        'CallEditData2()
        gpbSeek.Text = "����������"

        txtPId.ReadOnly = False
        txtSize.ReadOnly = False
        txtSizeDesc.ReadOnly = False
        txtPId.Focus()
    End Sub

    Private Sub btnSeekExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekExit.Click
        StateLockFindDept(True)
        gpbSeek.Text = ""
        gpbSeek.Visible = False  '����� gpbSeek ��͹
        IsShowSeek = False
    End Sub

Private Sub CheckSubDataBfSave()  '�礢����� gpbSeek ��͹ save ������
Dim i As Integer

   If txtPId.Text.Trim <> "" Then

        If txtSize.Text.Trim <> "" Then

                 If gpbSeek.Text = "����������" Then
                    SaveSubRecord()
                    ReSizeSort(txtEqp_id.Text.Trim)

                 Else
                    EditSubRecord()
                    ReSizeSort(txtEqp_id.Text.Trim)
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

   Else
        MsgBox("�ô�к�����ἧ" _
                       & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
        txtPId.Focus()

   End If

End Sub

Private Function EditSubRecord() As Boolean

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String

Dim strPrDate As String = ""
Dim strRecvDate As String = ""
Dim strFcDate As String = ""
Dim strEqpType As String = ""

    Try

    With Conn
         If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

    End With

             '----------------------------------------�ѹ����Դ����-----------------------------------------------

             If txtPrdate.Text <> "__/__/____" Then

                strPrDate = Mid(txtPrdate.Text.ToString, 7, 4) & "-" _
                                & Mid(txtPrdate.Text.ToString, 4, 2) & "-" _
                                & Mid(txtPrdate.Text.ToString, 1, 2)
                                strPrDate = "'" & SaveChangeEngYear(strPrDate) & "'"

            Else
                strPrDate = "NULL"
            End If

           '----------------------------------------�ѹ����Ѻ���--------------------------------------------------

            If txtIndate.Text <> "__/__/____" Then

               strRecvDate = Mid(txtIndate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtIndate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtIndate.Text.ToString, 1, 2)
                                 strRecvDate = "'" & SaveChangeEngYear(strRecvDate) & "'"

           Else
               strRecvDate = "NULL"
           End If


          '----------------------------------------�ѹ���Ѵ���---------------------------------------------------

           If txtFCdate.Text <> "__/__/____" Then

              strFcDate = Mid(txtFCdate.Text.ToString, 7, 4) & "-" _
                              & Mid(txtFCdate.Text.ToString, 4, 2) & "-" _
                              & Mid(txtFCdate.Text.ToString, 1, 2)
                              strFcDate = "'" & SaveChangeEngYear(strFcDate) & "'"
           Else
               strFcDate = "NULL"
           End If


            '------------------------------------��˹������������------------------------------------------------


             strSqlCmd = "UPDATE  tmp_eqptrn SET size_desc ='" & ReplaceQuote(txtPId.Text.ToString.Trim) & "'" _
                            & "," & "size_qty =" & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                            & "," & "dimns ='" & txtw.Text.ToString.Trim & " x " & _
                                                  txtLg.Text.ToString.Trim & "'" _
                            & "," & "price = " & ChangFormat(txtPrice.Text.ToString.Trim) _
                            & "," & "pr_doc ='" & ReplaceQuote(txtPr.Text.ToString.Trim) & "'" _
                            & "," & "men_rmk ='" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                            & "," & "set_qty =" & ChangFormat(txtSetQty.Text.ToString.Trim) _
                            & "," & "pr_date = " & strPrDate _
                            & "," & "recv_date = " & strRecvDate _
                            & "," & "fc_date = " & strFcDate _
                            & "," & "sup_name = '" & ReplaceQuote(txtSupplier.Text.ToString.Trim) & "'" _
                            & "," & "lp_type = '" & "LCA" & "'" _
                            & "," & "size_group = '" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                            & "," & "mouth_long = '" & txtMouth_mold.Text.Trim & "'" _
                            & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                            & " AND size_id ='" & txtSize.Text.ToString.Trim & "'" _
                            & " AND size_group = '" & txtSizeDesc.Text.ToString.Trim & "'" _
                            & " AND size_desc = '" & txtPId.Text.ToString.Trim & "'"

       Conn.Execute(strSqlCmd)

       Conn.Close()
       Conn = Nothing

    Catch ex As Exception
          MsgBox("����ͼԴ��Ҵ��зӡ�úѹ�֡ �ô���Թ��������ա����", MsgBoxStyle.Critical, "�Դ��Ҵ")
          MsgBox(ex.Message)
    End Try

End Function


Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
   Me.Close()
End Sub

Private Sub btnSeekSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekSave.Click
   CheckSubDataBfSave()
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

     Try

     With Conn
          If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

    End With

        '------------------------------------�礢�������͹�����������������-------------------------------------------------

        strSqlSelec = "SELECT * FROM tmp_eqptrn " _
                          & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                          & "AND size_id = '" & txtSize.Text.ToString.Trim & "'" _
                          & "AND size_desc = '" & txtPId.Text.ToString.Trim & "'" _
                          & "AND size_group = '" & txtSizeDesc.Text & "'"

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


                '---------------------------------------- Ǵ�.����Դ ----------------------------------------------------
                If txtCdate.Text <> "__/__/____" Then

                    strCreDate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                         & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                         & Mid(txtCdate.Text.ToString, 1, 2)
                    strCreDate = "'" & SaveChangeEngYear(strCreDate) & "'"

                Else
                    strCreDate = "NULL"
                End If

                '--------------------------�ѹ����Դ���� ------------------------------------------------

                If txtCdate.Text <> "__/__/____" Then   '�Ѵ �� ��͹ �ѹ
                    strCreDate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtCdate.Text.ToString, 1, 2)
                    strCreDate = "'" & SaveChangeEngYear(strCreDate) & "'"     '��ŧ����繻� ��.(� module)
                Else
                    strCreDate = "NULL"
                End If

                '-------------------------- �ѹ����Դ��觫���� -----------------------------------------------

                If txtPrdate.Text <> "__/__/____" Then
                    strPrdate = Mid(txtPrdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtPrdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtPrdate.Text.ToString, 1, 2)
                    strPrdate = "'" & SaveChangeEngYear(strPrdate) & "'"

                Else
                    strPrdate = "NULL"
                End If

                '-------------------------- �ѹ���Ѵ��� -----------------------------------------------

                If txtFCdate.Text <> "__/__/____" Then
                    strFcDate = Mid(txtFCdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtFCdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtFCdate.Text.ToString, 1, 2)
                    strFcDate = "'" & SaveChangeEngYear(strFcDate) & "'"

                Else
                    strFcDate = "NULL"
                End If


                '-------------------------- �ѹ����Ѻ��� -----------------------------------------------

                If txtIndate.Text <> "__/__/____" Then
                    strIndate = Mid(txtIndate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtIndate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtIndate.Text.ToString, 1, 2)
                    strIndate = "'" & SaveChangeEngYear(strIndate) & "'"

                Else
                    strIndate = "NULL"
                End If

                '----------------------- ��˹�������ͧ��鹧ҹ  -------------------------------

                Select Case cboPart.Text.ToString.Trim

                    Case Is = "�����ҧ"
                        strPartType = "SOLE2"
                    Case Is = "��鹺�"
                        strPartType = "SOLE1"
                    Case Is = "����鹺�"
                        strPartType = "FILL"
                    Case Is = "������"
                        strPartType = "LOGOHEEL"
                    Case Is = "������"
                        strPartType = "LOGO"
                    Case Is = "EVA �Դ��"
                        strPartType = "EVAH"
                    Case Is = "EVA ��˹ѧ˹��"
                        strPartType = "EVAF"
                    Case Is = "˹ѧ˹��"
                        strPartType = "UPPER"
                    Case Is = "ONUPPER"
                        strPartType = "ONUPPER"
                End Select


                strSqlCmd = "INSERT INTO tmp_eqptrn " _
                                     & "(user_id,size_id,size_desc,size_qty,set_qty" _
                                     & ",dimns,backgup,price,men_rmk,[group],eqp_id" _
                                     & ",delvr_sta,sent_sta,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                     & ",fc_date,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & ")" _
                                     & " VALUES (" _
                                     & "'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                     & ",'" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtPId.Text.ToString.Trim) & "'" _
                                     & "," & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                     & "," & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                     & ",'" & txtw.Text.ToString.Trim & " x " & _
                                                  txtLg.Text.ToString.Trim & "'" _
                                     & ",'" & "" & "'" _
                                     & "," & ChangFormat(txtPrice.Text.ToString.Trim) _
                                     & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                     & ",'" & "D" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "0" & "'" _
                                     & ",'" & "0" & "'" _
                                     & "," & strPrdate _
                                     & ",'" & ReplaceQuote(txtPr.Text.ToString.Trim) & "'" _
                                     & "," & strIndate _
                                     & "," & "0" _
                                     & "," & "0" _
                                     & "," & strFcDate _
                                     & ",'" & ReplaceQuote(txtSupplier.Text.ToString.Trim) & "'" _
                                     & ",'" & "LCA" & "'" _
                                     & ",'" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & txtMouth_mold.Text.Trim & "'" _
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

    Private Sub txtThk_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtThk.GotFocus
        With mskThk
            txtThk.SendToBack()
            .BringToFront()
            .Focus()
        End With

    End Sub

    Private Sub mskThk_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskThk.GotFocus
        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerge As String = ""

        With mskThk

            If txtThk.Text.Trim <> "0.0" Then
                x = Len(txtThk.Text.ToString)   '�Ѻ����ѡ��� txtThk

                For i = 1 To x
                    strTmp = Mid(txtThk.Text, i, 1)  '��ʵ�ԧ������ҡ����á��ش��ͤ���

                    Select Case strTmp     '�礤�� strTmp
                        Case Is = "_"
                        Case Is = "+"
                        Case Is = "_"
                        Case Else
                            If InStr(".0123456789", strTmp) > 0 Then
                                strMerge = strMerge & strTmp
                            End If
                    End Select
                Next i

                Select Case strMerge.IndexOf(".")
                    Case Is = 5
                        .SelectionStart = 0
                    Case Is = 3
                        .SelectionStart = 2
                    Case Is = 2
                        .SelectionStart = 3
                    Case Is = 1
                        .SelectionStart = 4
                    Case Else
                        .SelectionStart = 0

                End Select
                .SelectedText = strMerge

            End If
            .SelectAll()

        End With
    End Sub

    Private Sub mskThk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskThk.KeyDown
        Dim intChkPoint As Integer

        With mskThk
            Select Case e.KeyCode
                Case Is = 35   '���� End
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtTdesc.Focus()
                    End If

                Case Is = 38 '�����١�â��                                      
                Case Is = 39 '�����١�â��
                    If .SelectionStart = 0 Then
                        txtTtrait.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length   '���� ChkPoint = ������ǵ���ѡ��
                        If .SelectionStart = intChkPoint Then
                            txtTtrait.Focus()
                        End If

                    End If

                Case Is = 40  '����ŧ
                    txtPId.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With

    End Sub

    Private Sub mskThk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskThk.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTtrait.Focus()
        End If
    End Sub

    Private Sub mskThk_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskThk.LostFocus
        Dim i, x As Integer
        Dim z As Double

        Dim strTmp As String = ""
        Dim strMerge As String = ""

        With mskThk
            x = .Text.Length

            For i = 1 To x
                strTmp = Mid(.Text.ToString.Trim, i, 1)
                Select Case strTmp
                    Case Is = "_"
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
                mskThk.Text = ""
                z = CDbl(strMerge)
                txtThk.Text = z.ToString("#,##0.0")

            Catch ex As Exception
                txtThk.Text = "0.0"
                mskThk.Text = ""

            End Try

            mskThk.SendToBack()
            txtThk.BringToFront()
        End With

    End Sub

    Private Sub txtCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCdate.GotFocus
        With mskCdate
            txtCdate.SendToBack()
            .BringToFront()
            .Focus()
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
                    If .SelectionStart = 0 Then
                        txtTtime2.Focus()
                    End If
                Case Is = 38 '�١�â��
                    txtTeva.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtPId.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtPId.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtPId.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
    End Sub

    Private Sub mskCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskCdate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtPId.Focus()
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

            Catch ex As Exception
                mskCdate.Text = "__/__/____"
                txtCdate.Text = "__/__/____"

            End Try
            mskCdate.SendToBack()
            txtCdate.BringToFront()

        End With
    End Sub


    Private Sub mskLdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar = Chr(13) Then
            txtPId.Focus()
        End If
    End Sub


    Private Sub txtPrdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrdate.GotFocus
        With mskPrdate
             txtPrdate.SendToBack()
             .BringToFront()
             .Focus()
        End With
    End Sub

    Private Sub mskPrdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrdate.GotFocus
        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskPrdate

            If txtPrdate.Text.Trim <> "__/__/____" Then
               x = Len(txtPrdate.Text)

                For i = 1 To x

                    strTmp = Mid(txtPrdate.Text.Trim, i, 1)

                    Select Case strTmp
                        Case Is = "_"
                        Case Else
                            If InStr("0123456789/", strTmp) > 0 Then
                                strMerg = strMerg & strTmp
                            End If

                    End Select

                Next i

                Select Case strMerg.ToString.Length
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
                .SelectedText = strMerg

            End If
            .SelectAll()
        End With
    End Sub

    Private Sub mskPrdate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskPrdate.KeyDown
        Dim intChkPoint As Integer
        With mskPrdate
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtPr.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtPrice.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtFCdate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtFCdate.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ
                    txtPId.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
    End Sub

    Private Sub mskPrdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrdate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtFCdate.Focus()
        End If
    End Sub

    Private Sub mskPrdate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrdate.LostFocus
        Dim i, x As Integer
        Dim z As Date

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskPrdate
            x = .Text.Length

            For i = 1 To x
                strTmp = Mid(.Text.ToString.Trim, i, 1)
                Select Case strTmp
                    Case Is = ","
                    Case Is = "+"
                    Case Is = "_"
                    Case Else
                        If InStr("0123456789/", strTmp) > 0 Then
                            strMerg = strMerg & strTmp

                        End If
                End Select
                strTmp = ""
            Next i

            Try
                mskPrdate.Text = ""
                strMerg = "#" & strMerg & "#"
                z = CDate(strMerg)
                txtPrdate.Text = z.ToString("dd/MM/yyyy")

            Catch ex As Exception
                mskPrdate.Text = "__/__/____"
                txtPrdate.Text = "__/__/____"

            End Try

            mskPrdate.SendToBack()
            txtPrdate.BringToFront()
        End With

    End Sub

    Private Sub txtIndate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndate.GotFocus
        With mskIndate
            txtIndate.SendToBack()
            .BringToFront()
            .Focus()
        End With
    End Sub

    Private Sub mskIndate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskIndate.GotFocus
        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskIndate

            If txtIndate.Text.Trim <> "__/__/____" Then
                x = Len(txtIndate.Text)

                For i = 1 To x

                    strTmp = Mid(txtIndate.Text.Trim, i, 1)
                    Select Case strTmp
                        Case Is = "_"
                        Case Else
                            If InStr("0123456789/", strTmp) > 0 Then
                                strMerg = strMerg & strTmp
                            End If
                    End Select

                Next i

                Select Case strMerg.ToString.Length
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
                .SelectedText = strMerg

            End If
            .SelectAll()
        End With
    End Sub

Private Sub mskIndate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskIndate.KeyDown
  Dim intChkPoint As Integer
  With mskIndate
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtFCdate.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtPrdate.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSupplier.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSupplier.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ
                    txtRmk.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

  End With
End Sub

    Private Sub mskIndate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskIndate.KeyPress
        If e.KeyChar = Chr(13) Then
           txtSupplier.Focus()
        End If
    End Sub

    Private Sub mskIndate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskIndate.LostFocus
        Dim i, x As Integer
        Dim z As Date

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskIndate
            x = .Text.Length

            For i = 1 To x
                strTmp = Mid(.Text.ToString.Trim, i, 1)
                Select Case strTmp
                    Case Is = ","
                    Case Is = "+"
                    Case Is = "_"
                    Case Else
                        If InStr("0123456789/", strTmp) > 0 Then
                            strMerg = strMerg & strTmp

                        End If
                End Select
                strTmp = ""
            Next i

            Try
                mskIndate.Text = ""
                strMerg = "#" & strMerg & "#"
                z = CDate(strMerg)
                txtIndate.Text = z.ToString("dd/MM/yyyy")

                'If Year(z) < 2500 Then  '�óա�͡�� �.�. ����ŧ�� �.�. �ѹ��
                '    txtRecvDate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))
                'Else
                '    txtRecvDate.Text = z.ToString("dd/MM/yyyy")
                'End If

            Catch ex As Exception
                mskIndate.Text = "__/__/____"
                txtIndate.Text = "__/__/____"

            End Try


            mskIndate.SendToBack()
            txtIndate.BringToFront()
        End With

    End Sub

    Private Sub txtw_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtw.GotFocus
        With mskw
            txtw.SendToBack()
            .BringToFront()
            .Focus()
        End With
    End Sub

    Private Sub mskw_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskw.GotFocus
        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskw

            If txtw.Text <> "0.00" Then
                x = Len(txtw.Text.ToString)

                For i = 1 To x

                    strTmp = Mid(txtw.Text.ToString, i, 1)
                    Select Case strTmp
                        Case Is = "_"
                        Case Else
                            If InStr(",.0123456789", strTmp) > 0 Then
                                strMerg = strMerg & strTmp
                            End If

                    End Select
                Next i

                Select Case strMerg.IndexOf(".")
                    Case Is = 5
                        .SelectionStart = 0
                    Case Is = 3
                        .SelectionStart = 2
                    Case Is = 2
                        .SelectionStart = 3
                    Case Is = 1
                        .SelectionStart = 4
                    Case Else
                        .SelectionStart = 0
                End Select

                .SelectedText = strMerg


            End If
            .SelectAll()

        End With
    End Sub

    Private Sub mskw_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskw.KeyDown
        Dim intChkPoint As Integer
        With mskw
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtPr.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtPId.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtLg.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtLg.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ
                    txtRmk.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
    End Sub

    Private Sub mskw_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskw.KeyPress
        If e.KeyChar = Chr(13) Then
            txtLg.Focus()
        End If
    End Sub

    Private Sub mskw_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskw.LostFocus
        Dim i, x As Integer
        Dim z As Double

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskw

            x = .Text.ToString.Length

            For i = 1 To x
                strTmp = Mid(.Text.ToString, i, 1)
                Select Case strTmp
                    Case Is = ","
                    Case Is = "+"
                    Case Is = "_"
                    Case Else
                        If InStr(".0123456789", strTmp) Then
                            strMerg = strMerg & strTmp
                        End If

                End Select
                strTmp = ""
            Next i

            Try
                mskw.Text = ""
                z = CDbl(strMerg)
                txtw.Text = z.ToString("#,##0.00")
            Catch ex As Exception
                mskw.Text = ""
                txtw.Text = "0.00"
            End Try

            mskw.SendToBack()
            txtw.BringToFront()
        End With

    End Sub

    Private Sub txtLg_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLg.GotFocus
        With mskLg
            txtLg.SendToBack()
            .BringToFront()
            .Focus()
        End With
    End Sub

    Private Sub mskLg_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskLg.GotFocus
        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskLg

            If txtLg.Text <> "0.00" Then
                x = Len(txtLg.Text.ToString)

                For i = 1 To x

                    strTmp = Mid(txtLg.Text.ToString, i, 1)
                    Select Case strTmp
                        Case Is = "_"
                        Case Else
                            If InStr(",.0123456789", strTmp) > 0 Then
                                strMerg = strMerg & strTmp

                            End If
                    End Select
                Next i

                Select Case strMerg.IndexOf(".")
                    Case Is = 5
                        .SelectionStart = 0
                    Case Is = 3
                        .SelectionStart = 2
                    Case Is = 2
                        .SelectionStart = 3
                    Case Is = 1
                        .SelectionStart = 4
                    Case Else
                        .SelectionStart = 0

                End Select
                .SelectedText = strMerg
            End If

            .SelectAll()
        End With
    End Sub

    Private Sub mskLg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskLg.KeyDown
        Dim intChkPoint As Integer
        With mskLg
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtw.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtFCdate.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtPrice.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
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

    Private Sub mskLg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskLg.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)
                      txtPrice.Focus()
        End Select
    End Sub

    Private Sub mskLg_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskLg.LostFocus
        Dim i, x As Integer
        Dim z As Double

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskLg

            x = .Text.ToString.Length

            For i = 1 To x

                strTmp = Mid(.Text.ToString, i, 1)
                Select Case strTmp
                    Case Is = ","
                    Case Is = "+"
                    Case Is = "_"
                    Case Else
                        If InStr(".0123456789", strTmp) > 0 Then
                            strMerg = strMerg & strTmp
                        End If

                End Select
                strTmp = ""
            Next i

            Try
                mskLg.Text = ""
                z = CDbl(strMerg)
                txtLg.Text = z.ToString("#,##0.00")

            Catch ex As Exception
                mskLg.Text = ""
                txtLg.Text = "0.00"
            End Try

            mskLg.SendToBack()
            txtLg.BringToFront()
        End With

    End Sub

    Private Sub txtFCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFCdate.GotFocus
        With mskFCdate
            txtFCdate.SendToBack()
            .BringToFront()
            .Focus()
        End With
    End Sub

    Private Sub mskFCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFCdate.GotFocus
        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerg As String = ""

        With mskFCdate
            If txtFCdate.Text.Trim <> "__/__/____" Then
                x = Len(txtFCdate.Text.Trim)

                For i = 1 To x
                    strTmp = Mid(txtFCdate.Text.Trim, i, 1)
                    Select Case strTmp
                        Case Is = "_"
                        Case Else
                            If InStr("0123456789/", strTmp) > 0 Then
                                strMerg = strMerg & strTmp
                            End If
                    End Select

                Next i

                Select Case strMerg.ToString.Length
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
                .SelectedText = strMerg

            End If
            .SelectAll()

        End With
    End Sub

    Private Sub txtSetQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSetQty.GotFocus
        With mskSetQty
            txtSetQty.SendToBack()
            .BringToFront()
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
                    txtPrice.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

         End Select
        End With
    End Sub

    Private Sub mskSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSetQty.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)  '��� Enter
                txtSizeQty.Focus()
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

    Private Sub mskFCdate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskFCdate.KeyDown
        Dim intChkPoint As Integer
        With mskFCdate
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtw.Focus()
                    End If

                Case Is = 38 '�����١�â��
                    txtPr.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtIndate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtIndate.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ           
                    txtSupplier.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub mskFCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskFCdate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtIndate.Focus()
        End If
    End Sub

    Private Sub mskFCdate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFCdate.LostFocus
        Dim i, x As Integer
        Dim z As Date

        Dim strTmp As String = ""
        Dim strMerge As String = ""

        With mskFCdate

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

                mskFCdate.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)
                txtFCdate.Text = z.ToString("dd/MM/yyyy")

            Catch ex As Exception
                txtFCdate.Text = "__/__/____"
                mskFCdate.Text = ""

            End Try

            mskFCdate.SendToBack()
            txtFCdate.BringToFront()
        End With

    End Sub


Private Sub txtPId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPId.KeyDown
  Dim intChkPoint As Integer
      With txtPId

            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtCdate.Focus()
                    End If
                Case Is = 38 '�����١�â��  
                    txtTwarm.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSize.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSize.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtw.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
    End Sub

Private Sub txtPId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPId.KeyPress

  Select Case Asc(e.KeyChar)

         Case 48 To 122 ' �������ѧ��������ԧ�������� 58�֧122 ������� 48��������ҵ�ͧ��õ���Ţ
             e.Handled = False

         Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

         Case 13   'Enter = 13
              e.Handled = False
              txtSize.Focus()

         Case Else
             e.Handled = True
             MsgBox("��س��кآ������������ѧ������͵���Ţ", MsgBoxStyle.Critical, "�Դ��Ҵ")
  End Select

End Sub

Private Sub btnEditEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp1.Click

  Dim OpenFileDialog1 As New OpenFileDialog
  Dim strFileFullPath As String   '�纾������
  Dim strFileName As String       '�� filenaem
  Dim img As Image = Nothing

  Dim dateNow As Date = Now
  Dim typePic As String
  Dim strNamePic As String
  Dim lengPic, lengTypePic As Integer

      With OpenFileDialog1
           .CheckFileExists = True
           .ShowReadOnly = False
           .Filter = "All Files|*.*|����ٻ�Ҿ (*)|*.bmp;*.gif;*.jpg;*.png"
           .FilterIndex = 2

           Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    ' Load ������ picturebox
                    strFileName = New System.IO.FileInfo(.FileName).Name '�Ѻ���੾�Ъ������
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) '�Ѻ���੾�оҸ���

                    img = ScaleImage(Image.FromFile(.FileName), picEqp1.Height, picEqp1.Width)
                    picEqp1.Image = img

                    '----------- �����ѹ������ҵ�ͪ������ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4) ' �Ѵ��� .jpg .png .gif 
                    lengPic = strFileName.Length - 4   '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
                    strNamePic = Mid(strFileName, 1, lengPic)     '�Ѵ���੾�Ъ����ٻ
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ

                    lblPicPath1.Text = strFileFullPath
                    lblPicName1.Text = strFileName

                End If

            Catch ex As Exception
                  ClearBlankPicture1()
            End Try

        End With
        CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
    End Sub

Private Sub btnEditEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp2.Click
 Dim OpenFileDialog2 As New OpenFileDialog
 Dim strFileFullPath As String   '�纾������
 Dim strFileName As String       '�� filenaem
 Dim img As Image = Nothing

 Dim dateNow As Date = Now
 Dim typePic As String
 Dim strNamePic As String
 Dim lengPic, lengTypePic As Integer

        With OpenFileDialog2
            .CheckFileExists = True
            .ShowReadOnly = False
            .Filter = "All Files|*.*|����ٻ�Ҿ (*)|*.bmp;*.gif;*.jpg;*.png"
            .FilterIndex = 2

            Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    ' Load the specified file into a PictureBox control.
                    strFileName = New System.IO.FileInfo(.FileName).Name '�Ѻ���੾�Ъ������
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) '�Ѻ���੾�оҸ���

                    img = ScaleImage(Image.FromFile(.FileName), picEqp2.Height, picEqp2.Width)
                    picEqp2.Image = img

                    '----------- �����ѹ������ҵ�ͪ������ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4) ' �Ѵ��� .jpg .png .gif 
                    lengPic = strFileName.Length - 4   '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
                    strNamePic = Mid(strFileName, 1, lengPic)     '�Ѵ���੾�Ъ����ٻ
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ

                    lblPicPath2.Text = strFileFullPath
                    lblPicName2.Text = strFileName '�����������С�ȵ�������С���

                End If

            Catch ex As Exception
                ClearBlankPicture2()
            End Try

        End With
         CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
    End Sub

 Private Sub btnEditEqp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp3.Click

  Dim OpenFileDialog3 As New OpenFileDialog
  Dim strFileFullPath As String   '�纾������
  Dim strFileName As String       '�� filenaem
  Dim img As Image = Nothing

  Dim dateNow As Date = Now
  Dim typePic As String
  Dim strNamePic As String
  Dim lengPic, lengTypePic As Integer

      With OpenFileDialog3
           .CheckFileExists = True
           .ShowReadOnly = False
           .Filter = "All Files|*.*|����ٻ�Ҿ (*)|*.bmp;*.gif;*.jpg;*.png"
           .FilterIndex = 2

            Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    ' Load the specified file into a PictureBox control.
                    strFileName = New System.IO.FileInfo(.FileName).Name '�Ѻ���੾�Ъ������
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) '�Ѻ���੾�оҸ���

                    img = ScaleImage(Image.FromFile(.FileName), picEqp3.Height, picEqp3.Width)
                    picEqp3.Image = img

                    '----------- �����ѹ������ҵ�ͪ������ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4)        ' �Ѵ��� .jpg .png .gif 
                    lengPic = strFileName.Length - 4                      '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
                    strNamePic = Mid(strFileName, 1, lengPic)             '�Ѵ���੾�Ъ����ٻ
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ

                    lblPicPath3.Text = strFileFullPath
                    lblPicName3.Text = strFileName                    '�����������С�ȵ�������С���

                End If

            Catch ex As Exception
                ClearBlankPicture3()
            End Try

        End With
         CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
    End Sub

 Private Sub txtSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeQty.GotFocus

        With mskSizeQty
             txtSizeQty.SendToBack()
            .BringToFront()
            .Focus()
        End With

End Sub

Private Sub ClearBlankPicture3()
  picEqp3.Image = Nothing
  lblPicPath3.Text = ""
  lblPicName3.Text = ""
End Sub

Private Sub mskSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.GotFocus

 Dim i, x As Integer
 Dim strTmp As String = ""
 Dim strMerg As String = ""

     With mskSizeQty

            If txtSizeQty.Text.ToString.Trim <> "" Then
                x = Len(txtSizeQty.Text.ToString)

                For i = 1 To x
                    strTmp = Mid(txtSizeQty.Text.ToString, i, 1)
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

Private Sub mskSizeQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSizeQty.KeyDown

        Dim intChkpoint As Integer
        With mskSizeQty
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtSetQty.Focus()
                    End If
                Case Is = 38 '�����١�â��  

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtw.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkpoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtw.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtPr.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

    Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)
                txtw.Focus()
            Case Is = Chr(46)   '����ͧ���¨���Ҥ(.) 
                txtw.SelectionStart = 3
        End Select
    End Sub

    Private Sub mskSizeQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.LostFocus
        Dim i, x, intFull As Integer
        Dim z As Double

        Dim strTmp As String = ""
        Dim strMerg As String = ""
        With mskSizeQty
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
                mskSizeQty.Text = ""     '������ mskSizeQty
                z = CDbl(strMerg)        '�ŧ Type dbl
                intFull = CInt(z)

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

    Private Sub txtTwarm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTwarm.KeyDown
        Dim intChkPoint As Integer
        With txtTwarm
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtTeva.Focus()
                    End If
                Case Is = 38 '�����١�â��     
                    txtTsht.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtTtime1.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtTtime1.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtPId.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTwarm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTwarm.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13
                e.Handled = False
                txtTtime1.Focus()
            Case 8                  '���� Backspace
                e.Handled = False
            Case 32                 '��� spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Select
    End Sub

    Private Sub txtTtime1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTtime1.KeyDown
        Dim intChkPoint As Integer
        With txtTtime1
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtTwarm.Focus()
                    End If
                Case Is = 38 '�����١�â��  
                    txtTsht.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtTtime2.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtTtime2.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtPId.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTtime1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTtime1.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13
                e.Handled = False
                txtTtime2.Focus()
            Case 8                  '���� Backspace
                e.Handled = False
            Case 32                   '��� spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Select
    End Sub

    Private Sub txtTtime2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTtime2.KeyDown
        Dim intChkPoint As Integer
        With txtTtime2
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtTtime1.Focus()
                    End If
                Case Is = 38 '�����١�â��  
                    txtTeva.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtCdate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtCdate.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtPId.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTtime2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTtime2.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13
                e.Handled = False
                txtCdate.Focus()
            Case 8                  '���� Backspace
                e.Handled = False
            Case 32                  ' ��� spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Select
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
                     cboPart.DroppedDown = True
                     cboPart.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

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
                     txtOrder.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

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
                    txtEqp_id.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        cboPart.DroppedDown = True
                        cboPart.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            cboPart.DroppedDown = True
                            cboPart.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                     txtRemark.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
        Dim intChkPoint As Integer
        With txtRemark
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        cboPart.DroppedDown = True
                        cboPart.Focus()
                    End If
                Case Is = 38 '�����١�â��     
                    txtShoe.Focus()
                Case Is = 39 '�����١�â��
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

    Private Sub txtTdesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTdesc.KeyDown
        Dim intChkPoint As Integer
        With txtTdesc
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                    End If
                Case Is = 38 '�����١�â��                       
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtThk.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtThk.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ      
                    txtTsht.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
    End Sub

    Private Sub txtTtrait_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTtrait.KeyDown
        Dim intChkPoint As Integer
        With txtTtrait
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtThk.Focus()
                    End If
                Case Is = 38 '�����١�â��                       
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtTsht.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtTsht.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ      
                    txtTeva.Focus()
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

    Private Sub txtOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrder.KeyPress
        If e.KeyChar = Chr(13) Then
            cboPart.DroppedDown = True 
            cboPart.Focus()
        End If
    End Sub


    Private Sub txtAmount_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAmount.KeyDown
        Dim intChkPoint As Integer
        With txtAmount
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        cboPart.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtShoe.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSet.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSet.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ      
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtAmount_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAmount.KeyPress
        If e.KeyChar = Chr(13) Then
            txtSet.Focus()
        End If
    End Sub

    Private Sub txtSet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSet.KeyDown

        Dim intChkPoint As Integer
        With txtSet
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtAmount.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtOrder.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtRemark.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtRemark.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ      
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtSet_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSet.KeyPress
        If e.KeyChar = Chr(13) Then
            txtRemark.Focus()
        End If
    End Sub

    Private Sub txtTdesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTdesc.KeyPress
        If e.KeyChar = Chr(13) Then
            txtThk.Focus()
        End If
    End Sub

    Private Sub txtThk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtThk.KeyDown
        Dim intChkPoint As Integer
        With txtThk
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtTdesc.Focus()
                    End If
                Case Is = 38 '�����١�â��    

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtTtrait.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtTtrait.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtTeva.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtThk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtThk.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTtrait.Focus()
        End If
    End Sub

    Private Sub txtTtrait_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTtrait.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTsht.Focus()
        End If
    End Sub

    Private Sub txtTsht_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTsht.KeyDown
        Dim intChkPoint As Integer
        With txtTsht
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtTtrait.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtTdesc.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtTeva.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtTeva.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtTwarm.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTsht_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTsht.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTeva.Focus()
        End If
    End Sub

    Private Sub txtTeva_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTeva.KeyDown
        Dim intChkPoint As Integer
        With txtTeva
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtTsht.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtTtrait.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtTwarm.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtTwarm.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtCdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTeva_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTeva.KeyPress
        If e.KeyChar = Chr(13) Then
            txtTwarm.Focus()
        End If
    End Sub

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown
 Dim intChkPoint As Integer
 With txtSize

      Select Case e.KeyCode
             Case Is = 35 '���� End 
             Case Is = 36 '���� Home
             Case Is = 37 '�١�ë���
                  If .SelectionStart = 0 Then
                     txtPId.Focus()
                  End If
             Case Is = 38 '�����١�â��    
                    txtCdate.Focus()
             Case Is = 39 '�����١�â��
                   If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                       txtSizeDesc.Focus()   ''
                   Else
                       intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSizeDesc.Focus()
                        End If
                    End If
             Case Is = 40  '����ŧ  
                    txtPrice.Focus()
             Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress

   Select Case Asc(e.KeyChar)
          Case 48 To 57 ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ 0 57����Ţ 9����ӴѺ
                e.Handled = False

          Case 8, 46, 45  ' ���� Backspace = 8, ����Delete = 46 , ���� 45 = �մ��ҧ
                e.Handled = False

          Case 13 '���� Enter = 13
               e.Handled = False
               txtSizeDesc.Focus()

          Case Else
                e.Handled = True
                MsgBox("��س��кآ������������ѧ������͵���Ţ", MsgBoxStyle.Critical, "�Դ��Ҵ")

  End Select

End Sub

    Private Sub txtSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSizeQty.KeyPress
        If e.KeyChar = Chr(13) Then
            txtw.Focus()
        End If
    End Sub

Private Sub txtPr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPr.KeyDown
 Dim intChkPoint As Integer
     With txtPr
         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtSizeQty.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtSize.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtPrdate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtPrdate.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtFCdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

 End Sub

Private Sub txtPr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPr.KeyPress

Select Case Asc(e.KeyChar)

       Case 45 To 122 ' �������ѧ��������ԧ�������� 58�֧122 ������� 48��������ҵ�ͧ��õ���Ţ
             e.Handled = False

       Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

       Case 13   'Enter = 13
              e.Handled = False
              txtPrdate.Focus()

      Case Else
             e.Handled = True
             MessageBox.Show("��س��кآ������������ѧ���")
  End Select

End Sub

Private Sub txtPrdate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPrdate.KeyDown
 Dim intChkPoint As Integer
     With txtPrdate

         Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtPr.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtSizeQty.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtFCdate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtFCdate.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtFCdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtPrdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrdate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtFCdate.Focus()
        End If
    End Sub

    Private Sub txtFCdate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFCdate.KeyDown
        Dim intChkPoint As Integer
        With txtFCdate
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtPrdate.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtPr.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtIndate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtIndate.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtSupplier.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtFCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFCdate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtIndate.Focus()
        End If
    End Sub

    Private Sub txtIndate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtIndate.KeyDown
        Dim intChkPoint As Integer
        With txtIndate
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtFCdate.Focus()
                    End If
                Case Is = 38 '�����١�â��    
                    txtPrdate.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSupplier.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSupplier.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtSupplier.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtIndate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndate.KeyPress
        If e.KeyChar = Chr(13) Then
            txtSupplier.Focus()
        End If
    End Sub

    Private Sub txtSupplier_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSupplier.KeyDown

        Dim intChkPoint As Integer
        With txtSupplier
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtIndate.Focus()
                    End If

                Case Is = 38 '�����١�â��    
                    txtFCdate.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtMouth_mold.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtMouth_mold.Focus()
                        End If
                    End If
                Case Is = 40  '����ŧ  
                    txtRmk.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
    End Sub

    Private Sub txtSupplier_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSupplier.KeyPress
        If e.KeyChar = Chr(13) Then
            txtMouth_mold.Focus()
        End If
    End Sub

    Private Sub txtPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrice.GotFocus
        With mskPrice
             txtPrice.SendToBack()
            .BringToFront()
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

                Case Is = 38 '�����١�â��    
                    txtPId.Focus()
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
                    txtFCdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
    End Sub

    Private Sub mskPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrice.KeyPress
        Select Case e.KeyChar
               Case Is = Chr(13)  '��� Enter
                        txtPr.Focus()
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

    Private Sub txtSeries_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShoe.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)  '��� Enter
                txtOrder.Focus()
        End Select
    End Sub

    Private Sub txtLg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLg.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)
                     txtPrice.Focus()
        End Select
    End Sub

    Private Sub txtSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSetQty.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)
                txtSizeQty.Focus()
        End Select
    End Sub

    Private Sub txtWeight_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Select Case e.KeyChar
            Case Is = Chr(13)
                txtPrice.Focus()
        End Select
    End Sub

Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

  If dgvSize.Rows.Count > 0 Then
     ShowResvrd()       '�ʴ� gpbSeek 
     CallEditData()    '�Ѻ�ٷչ�ʴ� Size ������䢢�����
     CallEditData2()   '�Ѻ�ٷչ�ʴ������ŷҧ෤�Ԥ

     gpbSeek.Text = "��䢢�����"
     txtPId.ReadOnly = True
     txtSize.ReadOnly = True
     txtSizeDesc.ReadOnly = True
  Else
      MsgBox("�������¡�� SIZE ����ͧ������!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
            dgvSize.Focus()
  End If
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

   strSqlSelc = " SELECT tech_desc,tech_thk,tech_sht " _
                     & ",tech_eva,tech_warm,tech_time1 " _
                     & ",tech_time2,creat_date,tech_trait" _
                     & " FROM eqpmst (NOLOCK)" _
                     & " WHERE eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'" _
                     & " AND eqp_name = '" & txtEqpnm.Text.ToString.Trim & "'"

   With Rsd

         .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
         .LockType = ADODB.LockTypeEnum.adLockOptimistic
         .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then

             txtTdesc.Text = .Fields("tech_desc").Value.ToString.Trim
             txtThk.Text = .Fields("tech_thk").Value.ToString.Trim
             txtTtrait.Text = .Fields("tech_trait").Value.ToString.Trim
             txtTsht.Text = .Fields("tech_sht").Value.ToString.Trim
             txtTeva.Text = .Fields("tech_eva").Value.ToString.Trim
             txtTwarm.Text = .Fields("tech_warm").Value.ToString.Trim
             txtTtime1.Text = .Fields("tech_time1").Value.ToString.Trim
             txtTtime2.Text = .Fields("tech_time2").Value.ToString.Trim
            
                 If .Fields("creat_date").Value.ToString <> "" Then
                    txtCdate.Text = Mid(.Fields("creat_date").Value.ToString.Trim, 1, 10)
                 Else
                    txtCdate.Text = "__/__/____"
                 End If
        
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

Dim strWd As String = ""   '�纤�ҡ��ҧ
Dim strLg As String = ""   '�纤������
Dim strHg As String = ""   '�纤����٧

If dgvSize.Rows.Count <> 0 Then

   Dim strCode As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(2).Value.ToString.Trim     '�� Size
   Dim strLot As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(4).Value.ToString.Trim      '����ἧ
   Dim strGpsize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(5).Value.ToString.Trim    '���� size
 
    With Conn
         If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

    End With

            strSqlSelc = " SELECT * FROM v_tmp_eqptrn (NOLOCK)" _
                          & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                          & " AND size_id= '" & strCode & "'" _
                          & " AND size_desc = '" & strLot & "'" _
                          & " AND size_group = '" & strGpsize & "'"
    With Rsd

         .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
         .LockType = ADODB.LockTypeEnum.adLockOptimistic
         .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then

             txtPId.Text = .Fields("size_desc").Value.ToString.Trim
             txtSize.Text = .Fields("size_id").Value.ToString.Trim
             txtSizeDesc.Text = .Fields("size_group").Value.ToString.Trim
             txtSetQty.Text = Format(.Fields("set_qty").Value, "#,##0")
             txtSizeQty.Text = Format(.Fields("size_qty").Value, "#,##0")
             txtPrice.Text = Format(.Fields("price").Value, "#,##0.00")
             txtPr.Text = .Fields("pr_doc").Value.ToString.Trim
             txtSupplier.Text = .Fields("sup_name").Value.ToString.Trim
             txtRmk.Text = .Fields("men_rmk").Value.ToString.Trim


             If .Fields("pr_date").Value.ToString.Trim <> "" Then
                txtPrdate.Text = Mid(.Fields("pr_date").Value.ToString.Trim, 1, 10)
             Else
                txtPrdate.Text = "__/__/____"
             End If


             If .Fields("recv_date").Value.ToString.Trim <> "" Then
                txtIndate.Text = Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10)
             Else
                txtIndate.Text = "__/__/____"
             End If


             If .Fields("fc_date").Value.ToString.Trim <> "" Then
                txtFCdate.Text = Mid(.Fields("fc_date").Value.ToString.Trim, 1, 10)
             Else
                txtFCdate.Text = "__/__/____"
             End If

             If .Fields("dimns").Value.ToString.Trim <> "" Then

                 RetrnDiams(.Fields("dimns").Value.ToString.Trim, strWd, strLg)
                 txtw.Text = strWd
                 txtLg.Text = strLg

             Else
                 txtw.Text = "0.00"
                 txtLg.Text = "0.00"
             End If

             If .Fields("mouth_long").Value.ToString.Trim <> "" Then
                txtMouth_mold.Text = .Fields("mouth_long").Value.ToString.Trim

             Else
                  txtMouth_mold.Text = "0.00"
             End If

         End If
            .ActiveConnection = Nothing    '����������������
            .Close()

    End With
    Rsd = Nothing

Conn.Close()
Conn = Nothing
End If

End Sub

'----------------------- �ѧ���蹤׹��� Dimension -------------------------------------------------------------
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

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
  DeleteSubData()
End Sub

Private Sub DeleteSubData()
Dim btyConsider As Byte
Dim strSize As String = ""
Dim strSizeAct As String = ""
Dim strSizeDesc As String = ""
Dim strGpsize As String = ""

   With dgvSize

    If .Rows.Count > 0 Then
             strSize = .Rows(.CurrentRow.Index).Cells(2).Value.ToString
             strSizeAct = .Rows(.CurrentRow.Index).Cells(3).Value.ToString
             strSizeDesc = .Rows(.CurrentRow.Index).Cells(4).Value.ToString
             strGpsize = .Rows(.CurrentRow.Index).Cells(5).Value.ToString       '���� size


             If strSizeAct <> "" Then

                    btyConsider = MsgBox("SIZE : " & strSize.ToString.Trim & vbNewLine _
                                                   & "�����༧ : " & strSizeDesc.ToString.Trim & vbNewLine _
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
                                                  & " AND size_desc = '" & strSizeDesc.ToString.Trim & "'" _
                                                  & " AND size_group = '" & strGpsize.ToString.Trim & "'"

                              Conn.Execute(strCmd)

                             '------------------ ź������㹵��ҧ tmp_eqptrn_newsize -------------------

                              strCmd = "DELETE FROM tmp_eqptrn_newsize" _
                                                 & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                 & " AND size_id ='" & strSize.ToString.Trim & "'" _
                                                 & " AND size_desc = '" & strSizeDesc.ToString.Trim & "'"
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

Private Sub btnDelEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp1.Click
  ClearBlankPicture1()
  CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
End Sub

Private Sub ClearBlankPicture1()
  picEqp1.Image = Nothing
  lblPicPath1.Text = ""
  lblPicName1.Text = ""
End Sub

Private Sub ClearBlankPicture2()
  picEqp2.Image = Nothing
  lblPicPath2.Text = ""
  lblPicName2.Text = ""
End Sub

Private Sub btnDelEqp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp3.Click
  ClearBlankPicture3()
  CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
End Sub

Private Sub btnDelEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp2.Click
  ClearBlankPicture2()
  CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
End Sub

Private Sub txtPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice.KeyPress
  Select Case e.KeyChar
         Case Is = Chr(13)  '��� Enter
                   txtPr.Focus()
  End Select
End Sub

Private Sub txtRmk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRmk.KeyDown
    Dim intChkPoint As Integer
        With txtRmk
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

                    If .SelectionStart = 0 Then
                        txtMouth_mold.Focus()
                    End If

                Case Is = 38 '�����١�â��    
                    txtIndate.Focus()
                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        btnSeekSave.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            btnSeekSave.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
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

  Select Case Asc(e.KeyChar)

         Case 48 To 122 ' �������ѧ��������ԧ�������� 58�֧122 ������� 48��������ҵ�ͧ��õ���Ţ
             e.Handled = False

         Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

         Case 13   'Enter = 13
              e.Handled = False
             txtSetQty.Focus()

         Case Else
             e.Handled = True
             MessageBox.Show("��س��кآ������������ѧ���")
  End Select

End Sub

Private Sub txtEqp_id_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqp_id.LostFocus
  txtEqp_id.Text = txtEqp_id.Text.ToString.ToUpper.Trim()
End Sub

Private Sub txtPId_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPId.LostFocus
  txtPId.Text = txtPId.Text.ToString.ToUpper.Trim
End Sub

Private Sub txtSizeDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeDesc.LostFocus
  txtSizeDesc.Text = txtSizeDesc.Text.ToString.ToUpper.Trim
End Sub

Private Sub txtPr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPr.LostFocus
  txtPr.Text = txtPr.Text.ToString.ToUpper.Trim
End Sub

Private Sub txtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.LostFocus
  txtRemark.Text = txtRemark.Text.ToString.ToUpper.Trim
End Sub

Private Sub txtShoe_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShoe.LostFocus
  txtShoe.Text = txtShoe.Text.ToString.ToUpper.Trim
End Sub

'-------------------------------- �Ѻ�ٷչ����ʵ�ԧ Eqpid -----------------------------------------------
Private Sub AddtextEqpid()

  Select Case cboPart.Text.ToString.Trim
        Case Is = "��鹺�"
              txtEqp_id.Text = txtEqp_id.Text & "-A"
        Case Is = "����鹺�"
              txtEqp_id.Text = txtEqp_id.Text & "-B"
        Case Is = "������"
              txtEqp_id.Text = txtEqp_id.Text & "-C"
        Case Is = "�����ҧ"
              txtEqp_id.Text = txtEqp_id.Text & "-D"
        Case Is = "EVA �Դ��"
              txtEqp_id.Text = txtEqp_id.Text & "-E"
        Case Is = "������"
              txtEqp_id.Text = txtEqp_id.Text & "-F"
        Case Is = "˹ѧ˹��"
              txtEqp_id.Text = txtEqp_id.Text & "-G"
        Case Is = "EVA ��˹ѧ˹��"
              txtEqp_id.Text = txtEqp_id.Text & "-H"
        Case Is = "ONUPPER"
              txtEqp_id.Text = txtEqp_id.Text & "-I"

  End Select
End Sub

Private Sub lnkSave_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
    SaveNewRecord()
End Sub

Private Sub cboPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboPart.KeyPress
  Select Case e.KeyChar
         Case Is = Chr(13)  '��� Enter
                   txtRemark.Focus()
  End Select

End Sub

Private Sub cboPart_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPart.LostFocus
  Dim strEqpid As String
      strEqpid = txtEqp_id.Text.Trim

         '--------------- ����ʵ�ԧ "-" ������������ ----------------------

      If InStr(1, strEqpid, "-") > 0 Then
         txtEqp_id.Text = strEqpid
      Else
             AddtextEqpid()  '����ʵ�ԧ��ͷ��� eqp_id
      End If

End Sub

Private Sub txtEqpnm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpnm.LostFocus
  txtEqpnm.Text = txtEqpnm.Text.ToUpper.Trim
End Sub

Private Sub picEqp1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp1.Click

  Dim strFilePicture As String = ""

      If Not picEqp1.Image Is Nothing Then
         strFilePicture = lblPicPath1.Text.ToString.Trim & "\" & lblPicName1.Text.ToString.Trim

         Try

            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(strFilePicture)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

         Catch ex As Exception
               MessageBox.Show("File " & strFilePicture & " Could not be found!", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
         End Try

     End If

End Sub

Private Sub picEqp1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp1.MouseHover
  With tt
       .Show("��ԡ���ʹ��ٻ�˭�", picEqp1)
       .AutomaticDelay = 500
       .AutoPopDelay = 5000
       .InitialDelay = 100
 End With
End Sub

Private Sub picEqp1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp1.MouseLeave
  tt.Hide(picEqp1)
End Sub

Private Sub picEqp2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp2.Click

  Dim strFilePicture As String = ""

      If Not picEqp2.Image Is Nothing Then
         strFilePicture = lblPicPath2.Text.ToString.Trim & "\" & lblPicName2.Text.ToString.Trim

         Try

            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(strFilePicture)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

         Catch ex As Exception
               MessageBox.Show("File " & strFilePicture & " Could not be found!", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
         End Try

     End If

End Sub

Private Sub picEqp2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp2.MouseHover
  With tt
       .Show("��ԡ���ʹ��ٻ�˭�", picEqp2)
       .AutomaticDelay = 500
       .AutoPopDelay = 5000
       .InitialDelay = 100
 End With
End Sub

Private Sub picEqp2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp2.MouseLeave
   tt.Hide(picEqp2)
End Sub

Private Sub picEqp3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp3.Click
  Dim strFilePicture As String = ""

      If Not picEqp3.Image Is Nothing Then
         strFilePicture = lblPicPath3.Text.ToString.Trim & "\" & lblPicName3.Text.ToString.Trim

         Try

            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(strFilePicture)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

         Catch ex As Exception
               MessageBox.Show("File " & strFilePicture & " Could not be found!", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
         End Try

     End If

End Sub

Private Sub picEqp3_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp3.MouseHover
  With tt
       .Show("��ԡ���ʹ��ٻ�˭�", picEqp3)
       .AutomaticDelay = 500
       .AutoPopDelay = 5000
       .InitialDelay = 100
 End With
End Sub

Private Sub picEqp3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp3.MouseLeave
  tt.Hide(picEqp1)
End Sub

Private Function chkPicName(ByVal fnames As String) As Boolean

 Dim di As New DirectoryInfo("\\10.32.0.15\data1\EquipPicture\")
 Dim aryFi As FileInfo() = di.GetFiles(fnames)
 Dim fi As FileInfo

    For Each fi In aryFi
        If fnames = fi.Name Then
           Exit Function
           Return False
        End If
    Next

    Return True

End Function

Private Function Find_fixmold(ByVal user As String, ByVal idMold As String, ByVal mSize As String) As String

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim sqlSelc As String

     With Conn
          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

     sqlSelc = "SELECT fix_sta FROM tmp_fixeqptrn " _
                  & " WHERE user_id='" & user & "'" _
                  & " AND eqp_id='" & idMold & "'" _
                  & " AND size_id ='" & mSize & "'"

     With Rsd

         .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
         .LockType = ADODB.LockTypeEnum.adLockOptimistic
         .Open(sqlSelc, Conn, , , )

         If .RecordCount <> 0 Then
            Return .Fields("fix_sta").Value.ToString.Trim
         Else
             Return ""
         End If

       .ActiveConnection = Nothing
       .Close()
     End With

  Conn.Close()
End Function

Private Sub ReSizeSort(ByVal Eqpid As String)

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim sql As String
  Dim sqlCmd As String

  Dim strArr() As String
  Dim SearchWithinThis As String
  Dim newSize As String

  Dim prDate As String
  Dim RecvDate As String
  Dim FcDate As String
  Dim weight As Integer

      Try

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      sql = "SELECT * FROM tmp_eqptrn " _
                & " WHERE user_id= '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                & " ORDER BY size_id"

      With Rsd
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .CursorType = ADODB.CursorLocationEnum.adUseClient
           .Open(sql, Conn, , , )

           If .RecordCount <> 0 Then

               '----------------------- ��ҧ������� tmp_eqptrn_newsize ------------------------------

                 sqlCmd = "DELETE FROM tmp_eqptrn_newsize " _
                              & "WHERE user_id= '" & frmMainPro.lblLogin.Text.ToString.Trim & "'"

                 Conn.Execute(sqlCmd)

               ' ---------- ǹ�ٻ�Ѵ���§ size ���� --------------

               Do While Not .EOF

                  SearchWithinThis = .Fields("size_id").Value.ToString.Trim
                  If SearchWithinThis.IndexOf("-") <> -1 Then          '�ҡ size �鹩�Ѻ����� size ���� (x-xx)
                     strArr = SearchWithinThis.Split("-")              '��ҹ��� size �����㹵�����
                     newSize = strArr(0)
                  Else
                       newSize = .Fields("size_id").Value.ToString.Trim
                  End If

                  If .Fields("pr_date").Value.ToString.Trim <> "" Then
                      prDate = Mid(.Fields("pr_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("pr_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("pr_date").Value.ToString.Trim, 1, 2)
                      prDate = "'" & SaveChangeEngYear(prDate) & "'"

                  Else
                      prDate = "NULL"
                  End If

                  If .Fields("recv_date").Value.ToString.Trim <> "" Then
                      RecvDate = Mid(.Fields("recv_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("recv_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("recv_date").Value.ToString.Trim, 1, 2)
                      RecvDate = "'" & SaveChangeEngYear(RecvDate) & "'"
                  Else
                       RecvDate = "NULL"
                  End If

                  If .Fields("fc_date").Value.ToString.Trim <> "" Then
                     FcDate = Mid(.Fields("fc_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("fc_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("fc_date").Value.ToString.Trim, 1, 2)
                      FcDate = "'" & SaveChangeEngYear(FcDate) & "'"
                  Else
                       FcDate = "NULL"
                  End If

                  If .Fields("weight").Value.ToString.Trim = "" Then
                      weight = 0
                  Else
                       weight = .Fields("weight").Value.ToString.Trim
                  End If

                  '----------------------- Insert ������ŧ㹵��ҧ������ѧ���§ size ���� ----------------------

                   sqlCmd = "INSERT INTO tmp_eqptrn_newsize " _
                           & "(user_id,[group],eqp_id,size_id,size_desc,size_qty,weight,dimns,backgup " _
                           & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty " _
                           & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,tmp_newsize) " _
                           & "VALUES( " _
                           & "'" & frmMainPro.lblLogin.Text.Trim & "'" _
                           & ",'" & .Fields("group").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("eqp_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_desc").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("size_qty").Value.ToString.Trim) _
                           & "," & ChangFormat(weight) _
                           & ",'" & .Fields("dimns").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("backgup").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("price").Value.ToString.Trim) _
                           & ",'" & .Fields("men_rmk").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("delvr_sta").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("sent_sta").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("set_qty").Value.ToString.Trim) _
                           & "," & prDate _
                           & ",'" & .Fields("pr_doc").Value.ToString.Trim & "'" _
                           & "," & RecvDate _
                           & "," & ChangFormat(.Fields("ord_rep").Value.ToString.Trim) _
                           & "," & ChangFormat(.Fields("ord_qty").Value.ToString.Trim) _
                           & "," & FcDate _
                           & ",'" & .Fields("impt_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("sup_name").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("lp_type").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_group").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("cut_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("mate_type").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("cut_detail").Value.ToString.Trim & "'" _
                           & "," & newSize _
                           & ")"

                        Conn.Execute(sqlCmd)


                  .MoveNext()
               Loop

           End If

          .ActiveConnection = Nothing
          .Close()
      End With

   Conn.Close()

      Catch ex As Exception
            MsgBox(ex.Message)
      End Try

End Sub

Private Sub PreMoldStatus()
 Dim sta(4) As String
 Dim i As Byte

     sta(0) = "------ �ô���͡ʶҹ� ------"
     sta(1) = "���� / �ա������͹���"
     sta(2) = "����ҹ / ����ա������͹���"
     sta(3) = "¡��ԡ�����ҹ"

  With cmbStatus_mold

       For i = 0 To 3
           .Items.Add(sta(i))
       Next
      .SelectedIndex = 0
  End With

End Sub

Private Sub txtMouth_mold_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMouth_mold.GotFocus
  With mskMouth_mold
       .BringToFront()
       txtMouth_mold.SendToBack()
       .Focus()
  End With
End Sub

Private Sub mskMouth_mold_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskMouth_mold.GotFocus
  Dim i, x As Integer
  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskMouth_mold

           If txtMouth_mold.Text <> "0.00" Then

                        x = Len(txtMouth_mold.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtMouth_mold.Text.ToString, i, 1)
                                Select Case strTmp

                                       Case Is = "_"
                                       Case Else
                                            If InStr(",.0123456789", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If

                                End Select
                         Next i

                        Select Case strMerge.IndexOf(".")
                               Case Is = 5
                                    .SelectionStart = 0
                               Case Is = 3
                                    .SelectionStart = 2
                               Case Is = 2
                                    .SelectionStart = 3
                               Case Is = 1
                                    .SelectionStart = 4
                               Case Else
                                    .SelectionStart = 0
                        End Select

                        .SelectedText = strMerge

                End If
          .SelectAll()
       End With
End Sub

Private Sub mskMouth_mold_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskMouth_mold.KeyDown
   Dim intChkPoint As Integer

       With mskMouth_mold

               Select Case e.KeyCode
                      Case Is = 35 '���� End 
                      Case Is = 36 '���� Home
                      Case Is = 37 '�١�ë���

                               If .SelectionStart = 0 Then
                                  txtSupplier.Focus()
                               End If

                      Case Is = 38 '�����١�â��                                      
                      Case Is = 39 '�����١�â��

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtRmk.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
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

Private Sub mskMouth_mold_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskMouth_mold.KeyPress
   Select Case Asc(e.KeyChar)
          Case Is = 13
               txtRmk.Focus()

   End Select
End Sub

Private Sub mskMouth_mold_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskMouth_mold.LostFocus

  Dim i, x As Integer
  Dim z As Double

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskMouth_mold

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

                    mskMouth_mold.Text = ""
                    z = CDbl(strMerge)
                    txtMouth_mold.Text = z.ToString("#,##0.00")


                Catch ex As Exception
                    txtMouth_mold.Text = "0.00"
                    mskMouth_mold.Text = ""
               End Try

  mskMouth_mold.SendToBack()
  txtMouth_mold.BringToFront()

End With

End Sub

End Class