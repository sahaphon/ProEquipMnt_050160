Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeArkBlk

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

Private Sub frmAeArkBlk_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTable(0, "")  'ź������ Table tmp_eqptrn where user_id..
   frmArkBlk.lblCmd.Text = "0"  '������ʶҹ�
        Me.Dispose()     '����¿���� �׹˹��¤�����
    End Sub

Private Sub frmAeArkBlk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Dim dtComputer As Date = Now()
  Dim strCurrentDate As String

      StdDateTimeThai()        '�Ѻ�ٷչ��ŧ�ѹ�����Ǵ�.�� ����� Module
      strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

      ClearDataGpbHead()
      PrePartSeek()            '��Ŵ��������´���� cbo ������ǹ����Ե

        Select Case frmArkBlk.lblCmd.Text.ToString

               Case Is = "0" '����������

                    With Me
                         .Text = "����������"
                    End With

                    With txtBegin                 '��Ŵ�ѹ���Ѩ�غѹ���� txtBegin
                         .Text = strCurrentDate
                         strDateDefault = strCurrentDate
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

Private Sub LockEditData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim Rsdwc As New ADODB.Recordset

  Dim strCmd As String  ' ��ʵ�ԧ Command

  Dim strLoadFilePicture As String   '�纤��ʵ�ԧ��Ŵ Picture
  Dim strPathPicture As String = "\\10.32.0.15\data1\EquipPicture\"   '�� part

  Dim blnHavedata As Boolean   '�纤�ҵ����� ����Ѻ������բ������������
  Dim strSqlSelc As String = ""   '��ʵ�ԧ sql select
  Dim strPart As String = ""
        '�麤�� Row �Ѩ�غѹ㹿���� frmEqpSheet
  Dim strCod As String = frmArkBlk.dgvBlock.Rows(frmArkBlk.dgvBlock.CurrentRow.Index).Cells(0).Value.ToString.Trim

      With Conn

           If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
     End With
        ' NOLOCK ���¶֧����ö�ӧҹ�Ѻ Table �����١��ҡ user ���
        strSqlSelc = "SELECT * " _
                                    & "FROM v_moldinj_hd (NOLOCK)" _
                                    & " WHERE eqp_id = '" & strCod & "'"

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
            txtSet.Text = Format(.Fields("set_qty").Value, "##0.0")
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

                '-------------------------------Load �ٻ�Ҿ���ͤ-----------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_ctain").Value.ToString.Trim
                ' �紾�������������ԧ
               If File.Exists(strLoadFilePicture) Then

                    Dim img1 As Image      '��С�ȵ���� img1 �������Ҿ
                    img1 = Image.FromFile(strLoadFilePicture)  'img1 ��ҡѺpicture �����Ŵ�Ҩҡ db
                    'Dim s1 As String = ImageToBase64(img1, System.Drawing.Imaging.ImageFormat.Jpeg)  '��С�ȵ���� s1 �纤��ʵ�ԧ����ŧ����
                    'img1.Dispose()     '����µ���� img1

                    'picEqp1.Image = Base64ToImage(s1)
                    'picEqp1.SizeMode = PictureBoxSizeMode.CenterImage
                     picEqp1.Image = ScaleImage(img1, picEqp1.Height, picEqp1.Width)

                Else
                    picEqp1.Image = Nothing  '���������ٻ�Ҿ��� picEqp1 ��ҧ����
                End If


                '-------------------------------Load �ٻ�Ҿ��鹧ҹ ----------------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_io").Value.ToString.Trim
                '�社�������������ԧ
                If File.Exists(strLoadFilePicture) Then
                    Dim img2 As Image
                    img2 = Image.FromFile(strLoadFilePicture)
                    'Dim s2 As String = ImageToBase64(img2, System.Drawing.Imaging.ImageFormat.Jpeg)
                    'img2.Dispose()

                    'picEqp2.Image = Base64ToImage(s2)
                    'picEqp2.SizeMode = PictureBoxSizeMode.CenterImage
                     picEqp2.Image = ScaleImage(img2, picEqp2.Height, picEqp2.Width)

                Else
                    picEqp2.Image = Nothing
                End If


                 '-------------------------------Load �ٻ�Ҿ�ͧ���������ٻ  ----------------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_part").Value.ToString.Trim
                '�社�������������ԧ
                If File.Exists(strLoadFilePicture) Then
                    Dim img3 As Image
                    img3 = Image.FromFile(strLoadFilePicture)
                    'Dim s3 As String = ImageToBase64(img3, System.Drawing.Imaging.ImageFormat.Jpeg)
                    'img3.Dispose()

                    'picEqp3.Image = Base64ToImage(s3)
                    'picEqp3.SizeMode = PictureBoxSizeMode.CenterImage
                    picEqp3.Image = ScaleImage(img3, picEqp3.Height, picEqp3.Width)

                Else
                    picEqp3.Image = Nothing
                End If



                strCmd = frmArkBlk.lblCmd.Text.ToString.Trim    '��� strCmd ��ҡѺ���� lblcmd 㹿���� frmEqpSheet

                Select Case strCmd
                    Case Is = "1"   '�����ͤ�͹���
                    Case Is = "2"   '�����ͤ�͹����ͧ
                        btnSaveData.Enabled = False  '�Դ���� "�ѹ�֡������"
                End Select


                '------------------------------- �ѹ�֡�����ŧ㹵��ҧ tmp_eqptrn----------------------------
                '-- Insert ������ŧ���ҧ tmp_eqptrn ����Ң����Ũҡ���ҧ  tmp_eqptrn �Ҵ���----------------------

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
     cboPart.Text = ""
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

Private Sub ShowScrapItem()              '�ʴ�������� DataGridview 
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String             '�纤�� string command

 Dim dubQty As Double
 Dim dubAmt As Double
 Dim sngSetQty As Single                '�纨ӹǹ SET

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
                                        .Fields("dimns").Value.ToString.Trim, _
                                        .Fields("pr_doc").Value.ToString.Trim, _
                                        Format(.Fields("set_qty").Value, "##0.0"), _
                                        Format(.Fields("size_qty").Value, "##0.0"), _
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

                  If cboPart.Text.ToString.Trim <> "" Then

                          If IntListwc > 0 Then

                             bytConSave = MsgBox("�س��ͧ��úѹ�֡���������������!" _
                                  , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                    If bytConSave = 6 Then

                                             Select Case Me.Text

                                                    Case Is = "����������"

                                                        If CheckCodeDuplicate(txtEqp_id.Text) Then
                                                           SaveNewRecord()

                                                        Else
                                                            MessageBox.Show("�����ػ�ó��� ��سҡ�͡�����ػ�ó�����!....", "Error", _
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

Private Sub SaveNewRecord()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String
  Dim dateSave As Date = Now()    '�纤���ѹ���Ѩ�غѹ
  Dim strDate As String

  Dim blnRetuneCopyPic As Boolean
  Dim strCredate As String
  Dim strDateDoc As String
  Dim strPartType As String = ""
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

                '---------------------------------------- �ѹ�֡�ٻ���ͤ ------------------------------------------

                    blnRetuneCopyPic = CallCopyPicture(lblPicPath1.Text.Trim, ReturnImageName(lblPicName1.Text.Trim), lblPicName1.Text.Trim)

                    If blnRetuneCopyPic Then                     '��� CallCopyPicture = true
                       lblPicPath1.Text = PthName
                    Else
                       lblPicPath1.Text = ""
                       lblPicName1.Text = ""
                       picEqp1 = Nothing

                    End If

                '---------------------------------------- �ѹ�֡�ٻ�����ǹ ------------------------------------------

                    blnRetuneCopyPic = CallCopyPicture(lblPicPath2.Text.Trim, ReturnImageName(lblPicName2.Text.Trim), lblPicName2.Text.Trim)

                    If blnRetuneCopyPic Then                     '��� CallCopyPicture = true
                       lblPicPath2.Text = PthName
                    Else
                       lblPicPath2.Text = ""
                       lblPicName2.Text = ""
                       picEqp2 = Nothing

                    End If

                '---------------------------------------- �ѹ�֡�ٻ��Ե�ѳ�� -----------------------------------------

                    blnRetuneCopyPic = CallCopyPicture(lblPicPath3.Text.Trim, ReturnImageName(lblPicName3.Text.Trim), lblPicName3.Text.Trim)

                    If blnRetuneCopyPic Then                   '��� CallCopyPicture = true
                       lblPicPath3.Text = PthName
                    Else
                       lblPicPath3.Text = ""
                       lblPicName3.Text = ""
                       picEqp3 = Nothing

                    End If

                '---------------------------------------- Ǵ�.����Դ ----------------------------------------------

                    If txtCdate.Text <> "__/__/____" Then

                       strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtCdate.Text.ToString, 1, 2)
                       strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

                   Else
                       strCredate = "NULL"
                   End If

                    '------------------------------------��˹�������ͧ��鹧ҹ-----------------------------------------

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
                      & ",'" & "G" & "'" _
                      & ",'" & ReplaceQuote(txtEqp_id.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtEqpnm.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(txtRef.Text.ToString.Trim) & "'" _
                      & ",'" & ChangFormat(txtSet.Text.ToString.Trim) & "'" _
                      & ",'" & strPartType.ToString.Trim & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(lblPicName1.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(lblPicName2.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(lblPicName3.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtTdesc.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtThk.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(txtMaterial.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
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

               '------------------------------------------------�ѹ�֡������㹵��ҧ eqptrn------------------------------------------------

                strSqlCmd = "INSERT INTO eqptrn " _
                                     & " SELECT [group] ='G'" _
                                     & ",eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'" _
                                     & ",size_id,size_desc,size_qty,weight,dimns,backgup,price,men_rmk" _
                                     & ",delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                     & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & " FROM tmp_eqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                Conn.Execute(strSqlCmd)
                Conn.CommitTrans()

                frmArkBlk.lblCmd.Text = txtEqp_id.Text.ToString.Trim     '�觺͡��Һѹ�֡�����������
                frmArkBlk.Activating()
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


              '---------------------------------------- Ǵ�.����Դ -------------------------------------

               If txtCdate.Text <> "__/__/____" Then

                  strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtCdate.Text.ToString, 1, 2)
                  strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

               Else
                  strCredate = "NULL"

               End If

              '------------------------------------ ��˹�������ͧ��鹧ҹ --------------------------------

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


                      '------------------------------------ �ѹ�֡�ٻ���ͤ ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath1.Text.ToString.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath1.Text = PthName

                       Else
                          lblPicName1.Text = ""
                          lblPicPath1.Text = ""
                          picEqp1.Image = Nothing

                       End If


                      '------------------------------------ �ѹ�֡�ٻ�����ǹ ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath2.Text.ToString.Trim, ReturnImageName(lblPicName2.Text.ToString.Trim), lblPicName2.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath2.Text = PthName

                       Else
                          lblPicName2.Text = ""
                          lblPicPath2.Text = ""
                          picEqp2.Image = Nothing

                       End If

                      '------------------------------------ �ѹ�֡�ٻ��Ե�ѳ�� ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath3.Text.ToString.Trim, ReturnImageName(lblPicName3.Text.ToString.Trim), lblPicName3.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath3.Text = PthName

                       Else
                          lblPicName3.Text = ""
                          lblPicPath3.Text = ""
                          picEqp3.Image = Nothing

                       End If

                      '---------------------------------- UPDATE ������㹵��ҧ eqpmst -------------------------------------

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
                                                & "," & "tech_lg = '" & "" & "'" _
                                                & "," & "tech_sht = '" & "" & "'" _
                                                & "," & "tech_eva = '" & ReplaceQuote(txtMaterial.Text.ToString.Trim) & "'" _
                                                & "," & "tech_warm = '" & "" & "'" _
                                                & "," & "tech_time1 = '" & "" & "'" _
                                                & "," & "tech_time2 = '" & "" & "'" _
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
                                      & "SELECT [group] = 'G'" _
                                      & ",eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'" _
                                      & ",size_id,size_desc,size_qty,weight,dimns,backgup" _
                                      & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date" _
                                      & ",pr_doc,recv_date,ord_rep,ord_qty,fc_date,impt_id" _
                                      & ",sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long " _
                                      & " FROM tmp_eqptrn " _
                                      & " WHERE user_id= '" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                    Conn.Execute(strSqlCmd)
                    Conn.CommitTrans()  '��� Commit transection

                    frmArkBlk.lblCmd.Text = txtEqp_id.Text.ToString.Trim     '�觺͡��Һѹ�֡�����������
                    frmArkBlk.Activating()
                    Me.Close()

                    'lblComplete.Text = txtEqp_id.Text.ToString.Trim  '�觺͡��Һѹ�֡�����������

                    'Me.Hide()
                    'frmMainPro.Show()
                    'frmArkBlk.Show()

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

        Rsd = New ADODB.Recordset

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

Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String, ByVal strNewPicName As String) As Boolean
 Dim fname As String = String.Empty
 Dim dFile As String = String.Empty
 Dim dFilePath As String = String.Empty

 Dim fServer As String = String.Empty
 Dim intResult As Integer

     On Error GoTo Err70

        fname = strPicPath & "\" & strPicName            '���� \\10.32.0.15\data1\EquipPicture\"�����ٻ�Ҿ" 
        fServer = PthName & "\" & strNewPicName          'partServer \\10.32.0.15\data1\EquipPicture\"�����ٻ�Ҿ"

         If File.Exists(fServer) Then          '�������������ԧ
            CallCopyPicture = True             '���׹��� true

         Else
            If File.Exists(fname) Then
               dFile = Path.GetFileName(fname)
               dFilePath = DrvName + dFile


               intResult = String.Compare(fname.ToString.Trim, dFilePath.ToString.Trim)

                    '--------------------------- ��Ҥ���� 0 �ʴ������Ŵ��������� �������ö Copy ����� ------------------------------

                    If intResult = 1 Then                                                     '��ҷ���� = 1 �֧ copy �ٻ�����������ͧ 10.32.0.14
                       File.Copy(fname, dFilePath, True)
                    End If

                    My.Computer.FileSystem.RenameFile(dFilePath, strNewPicName)               '����¹��������ٻ����
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

Private Sub ShowResvrd()  '����ʴ� GroupBox gpbSeek �����

 tabMain.SelectedTab = tabSize
 IsShowSeek = Not IsShowSeek  '���ʶҹ�ᶺ seek ����ʴ� ����ʴ�

   If IsShowSeek Then

       With gpbSeek
            .Visible = True
            .Left = 8    '᡹ X
            .Top = 230   '᡹ Y 252
            .Height = 500
            .Width = 1006
       End With

            StateLockFindDept(False) ' ��ͤ FindDept ���觤���� False �
            txtTdesc.Focus()
            'dgvSeek.Focus()
   Else
            StateLockFindDept(True)
            'dgvSeek.Visible = False
        End If
   End Sub

Private Sub StateLockFindDept(ByVal sta As Boolean)
 Dim strMode As String = frmArkBlk.lblCmd.Text.ToString   '����� lblCmd ������ѹ�֡�����ź��ͤ���� �觤������� strMode 
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
   Me.Close()
End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
  ShowResvrd()
  ClearSubData2()
  CallEditData()    '�Ѻ�ٷչ�ʴ� Size ������䢢�����
  'CallEditData2()
  gpbSeek.Text = "����������"
  txtBlockID.ReadOnly = False
  txtSize.ReadOnly = False
  txtSizeDesc.ReadOnly = False
  txtBlockID.Focus()
End Sub

Private Sub ClearSubData2()
   txtBlockID.Text = ""
   txtPart.Text = ""
   txtSize.Text = ""
   txtSizeDesc.Text = ""
   txtSetQty.Text = "0"
   txtw.Text = "0.00"
   txtLg.Text = "0.00"
   txtPrice.Text = "0.00"
   txtRmk.Text = ""

End Sub

Private Sub ClearSubData()
  txtTdesc.Text = ""
  txtThk.Text = "0.0"
  txtMaterial.Text = ""
  txtCdate.Text = "__/__/____"

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
             txtMaterial.Text = .Fields("tech_eva").Value.ToString.Trim

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

          Dim strSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(4).Value.ToString.Trim     '�� Size
          Dim strBlockID As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(6).Value.ToString.Trim      '�����ʺ��ͤ
          Dim strGroupSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(7).Value.ToString.Trim      '�����ʺ��ͤ

          With Conn

               If .State Then Close()
                  .ConnectionString = strConnAdodb
                  .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                  .ConnectionTimeout = 90
                  .Open()

          End With

          strSqlSelc = " SELECT * FROM v_tmp_eqptrn (NOLOCK)" _
                          & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                          & " AND size_id= '" & strSize & "'" _
                          & " AND size_group = '" & strGroupSize & "'" _
                          & " AND size_desc = '" & strBlockID & "'"

          With Rsd

                 .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                 .LockType = ADODB.LockTypeEnum.adLockOptimistic
                 .Open(strSqlSelc, Conn, , , )

                 If .RecordCount <> 0 Then

                     txtBlockID.Text = .Fields("size_desc").Value.ToString.Trim
                     txtPart.Text = .Fields("backgup").Value.ToString.Trim
                     txtSize.Text = .Fields("size_id").Value.ToString.Trim
                     txtPr.Text = .Fields("pr_doc").Value.ToString.Trim
                     txtSizeDesc.Text = .Fields("size_group").Value.ToString.Trim
                     txtSetQty.Text = Format(.Fields("set_qty").Value, "##0.0")
                     txtSizeQty.Text = Format(.Fields("size_qty").Value, "##0.0")
                     txtPrice.Text = Format(.Fields("price").Value, "#,##0.00")
                     txtRmk.Text = .Fields("men_rmk").Value.ToString.Trim

                    If .Fields("dimns").Value.ToString.Trim <> "" Then
                        RetrnDiams(.Fields("dimns").Value.ToString.Trim, strWd, strLg)
                        txtw.Text = strWd
                        txtLg.Text = strLg

                    Else
                        txtw.Text = "0.00"
                        txtLg.Text = "0.00"
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

   If txtBlockID.Text.Trim <> "" Then

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


   Else
        MsgBox("�ô�к�����ἧ" _
                       & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
        txtBlockID.Focus()
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
  Dim strDateNull As String = "NULL"       '����Ѻ������� null �ѹ��� 

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
                             & " AND size_id = '" & txtSize.Text.ToString.Trim & "'" _
                             & " AND size_desc = '" & txtBlockID.Text.ToString.Trim & "'" _
                             & " AND size_group = '" & txtSizeDesc.Text.ToString.Trim & "'"

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelec, Conn, , , )

            If .RecordCount <> 0 Then        '��� RecordSet �բ�����
                MessageBox.Show("���ʺ��ͤ : " & txtBlockID.Text.ToString & " / " & "Size : " & txtSize.Text.ToString _
                                              & " ����к����� ��س��кآ���������", "�����ū��!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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

                '--------------------------�ѹ����Դ���� ------------------------------------------------

                If txtCdate.Text <> "__/__/____" Then   '�Ѵ �� ��͹ �ѹ
                    strCreDate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtCdate.Text.ToString, 1, 2)
                    strCreDate = "'" & SaveChangeEngYear(strCreDate) & "'"     '��ŧ����繻� ��.(� module)
                Else
                    strCreDate = "NULL"
                End If

                '------------------------------------ ��˹�������ͧ��鹧ҹ ------------------------------

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
                                     & ",'" & ReplaceQuote(txtBlockID.Text.ToString.Trim) & "'" _
                                     & "," & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                     & "," & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                     & ",'" & txtw.Text.ToString.Trim & " x " & _
                                              txtLg.Text.ToString.Trim & "'" _
                                     & ",'" & ReplaceQuote(txtPart.Text.ToString.Trim) & "'" _
                                     & "," & ChangFormat(txtPrice.Text.ToString.Trim) _
                                     & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                     & ",'" & "G" & "'" _
                                     & ",'" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                     & ",'" & "0" & "'" _
                                     & ",'" & "0" & "'" _
                                     & "," & strDateNull _
                                     & ",'" & ReplaceQuote(txtPr.Text.ToString.Trim) & "'" _
                                     & "," & strDateNull _
                                     & "," & "0" _
                                     & "," & "0" _
                                     & "," & strDateNull _
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
  Dim blnReturnCopyPic As Boolean

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


                      '------------------------------------ �ѹ�֡�ٻ���ͤ ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath1.Text.ToString.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath1.Text = PthName

                       Else
                          lblPicName1.Text = ""
                          lblPicPath1.Text = ""
                          picEqp1.Image = Nothing

                       End If


                      '------------------------------------ �ѹ�֡�ٻ�����ǹ ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath2.Text.ToString.Trim, ReturnImageName(lblPicName2.Text.ToString.Trim), lblPicName2.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath2.Text = PthName

                       Else
                          lblPicName2.Text = ""
                          lblPicPath2.Text = ""
                          picEqp2.Image = Nothing

                       End If

                       '------------------------------------ �ѹ�֡�ٻ��Ե�ѳ�� ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath3.Text.ToString.Trim, ReturnImageName(lblPicName3.Text.ToString.Trim), lblPicName3.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath3.Text = PthName

                       Else
                          lblPicName3.Text = ""
                          lblPicPath3.Text = ""
                          picEqp3.Image = Nothing

                       End If

                       strSqlCmd = "UPDATE  tmp_eqptrn SET size_desc ='" & ReplaceQuote(txtBlockID.Text.ToString.Trim) & "'" _
                                                        & "," & "size_qty = " & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                                        & "," & "dimns ='" & txtw.Text.ToString.Trim & " x " & _
                                                                             txtLg.Text.ToString.Trim & "'" _
                                                        & "," & "price = " & ChangFormat(txtPrice.Text.ToString.Trim) _
                                                        & "," & "backgup = '" & ReplaceQuote(txtPart.Text.ToString.Trim) & "'" _
                                                        & "," & "pr_doc ='" & ReplaceQuote(txtPr.Text.ToString.Trim) & "'" _
                                                        & "," & "men_rmk ='" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                                        & "," & "set_qty =" & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                                        & "," & "pr_date = " & strDateNull _
                                                        & "," & "recv_date = " & strDateNull _
                                                        & "," & "fc_date = " & strDateNull _
                                                        & "," & "sup_name = '" & "" & "'" _
                                                        & "," & "lp_type = '" & "" & "'" _
                                                        & "," & "size_group = '" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                                        & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                                        & " AND size_id = '" & txtSize.Text.ToString.Trim & "'" _
                                                        & " AND size_desc = '" & txtBlockID.Text.ToString.Trim & "'" _
                                                        & " AND size_group = '" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'"

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
     ClearAllData()
     ShowResvrd()       '�ʴ� gpbSeek 
     CallEditData()    '�Ѻ�ٷչ�ʴ� Size ������䢢�����
     CallEditData2()   '�Ѻ�ٷչ�ʴ������ŷҧ෤�Ԥ
     gpbSeek.Text = "��䢢�����"
     txtBlockID.ReadOnly = True
     txtSize.ReadOnly = True
     txtSizeDesc.ReadOnly = True
  Else
      MsgBox("�������¡�� SIZE ����ͧ������!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
      dgvSize.Focus()
  End If
End Sub

Private Sub ClearAllData()
  txtBlockID.Text = ""
  txtPart.Text = ""
  txtSize.Text = ""
  txtSizeDesc.Text = ""
  txtSetQty.Text = "0"
  txtSizeQty.Text = "0"
  txtPrice.Text = "0.00"
  txtw.Text = "0.00"
  txtLg.Text = "0.00"
  txtPr.Text = ""
  txtRmk.Text = ""
End Sub

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
  DeleteSubData()
End Sub

Private Sub DeleteSubData()
 Dim btyConsider As Byte
 Dim strSize As String = ""
 Dim strSizeAct As String = ""
 Dim strSizeBlock As String = ""
 Dim stgGpSize As String

   With dgvSize

    If .Rows.Count > 0 Then

          strSize = .Rows(.CurrentRow.Index).Cells(4).Value.ToString
          strSizeAct = .Rows(.CurrentRow.Index).Cells(5).Value.ToString
          strSizeBlock = .Rows(.CurrentRow.Index).Cells(6).Value.ToString
          stgGpSize = .Rows(.CurrentRow.Index).Cells(7).Value.ToString


             If strSizeAct <> "" Then

                    btyConsider = MsgBox("SIZE : " & strSize.ToString.Trim & vbNewLine _
                                                   & "���ʺ��ͤ : " & strSizeBlock.ToString.Trim & vbNewLine _
                                                   & "����䫵� : " & stgGpSize.ToString.Trim & vbNewLine _
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
                                                  & " AND size_group = '" & stgGpSize.ToString.Trim & "'" _
                                                  & " AND size_desc = '" & strSizeBlock.ToString.Trim & "'"

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
                     cboPart.DroppedDown = True
                     cboPart.Focus()

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

Private Sub txtOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrder.KeyPress
 If e.KeyChar = Chr(13) Then
    cboPart.DroppedDown = True
    cboPart.Focus()
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
                        cboPart.DroppedDown = True
                        cboPart.Focus()
                    End If
                Case Is = 38 '�����١�â��      
                        cboPart.DroppedDown = True
                        cboPart.Focus()
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

Private Sub txtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.LostFocus
  txtRemark.Text = txtRemark.Text.ToUpper.Trim
End Sub

Private Sub cboPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboPart.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRemark.Focus()
  End If
End Sub

Private Sub txtTdesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTdesc.KeyDown
  Dim intChkPoint As Integer
        With txtTdesc
            Select Case e.KeyCode

                Case Is = 35 '���� End 

                Case Is = 36 '���� Home

                Case Is = 37 '�١�ë���

                Case Is = 38 '�����١�â��      

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtThk.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtThk.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                      txtMaterial.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtTdesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTdesc.KeyPress
  If e.KeyChar = Chr(13) Then
     txtThk.Focus()
  End If
End Sub

Private Sub txtTdesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTdesc.LostFocus
  txtTdesc.Text = txtTdesc.Text.ToUpper.Trim
End Sub

Private Sub txtThk_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtThk.GotFocus
  With mskThk
       .BringToFront()
       txtThk.SendToBack()
       .Focus()
  End With
End Sub

Private Sub txtThk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtThk.KeyDown
 Dim intChkPoint As Integer
        With txtThk
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
                        txtMaterial.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtMaterial.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                         txtMaterial.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtThk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtThk.KeyPress
 If e.KeyChar = Chr(13) Then
    txtMaterial.Focus()
 End If
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
                        txtMaterial.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length   '���� ChkPoint = ������ǵ���ѡ��
                        If .SelectionStart = intChkPoint Then
                            txtMaterial.Focus()
                        End If

                    End If

                Case Is = 40  '����ŧ
                    txtMaterial.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub mskThk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskThk.KeyPress
  If e.KeyChar = Chr(13) Then
            txtMaterial.Focus()
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

Private Sub txtMaterial_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMaterial.KeyDown
 Dim intChkPoint As Integer

     With txtMaterial
            Select Case e.KeyCode

                Case Is = 35   '���� End
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

                    If .SelectionStart = 0 Then
                        txtThk.Focus()
                    End If

                Case Is = 38 '�����١�â��      
                      txtThk.Focus()
                Case Is = 39 '�����١�â��

                    If .SelectionStart = 0 Then
                        txtCdate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length   '���� ChkPoint = ������ǵ���ѡ��
                        If .SelectionStart = intChkPoint Then
                            txtCdate.Focus()
                        End If

                    End If

                Case Is = 40  '����ŧ
                    txtCdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub txtMaterial_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMaterial.KeyPress
 If e.KeyChar = Chr(13) Then
    txtCdate.Focus()
 End If
End Sub

Private Sub txtMaterial_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMaterial.LostFocus
   txtMaterial.Text = txtMaterial.Text.ToUpper.Trim
End Sub

Private Sub txtCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCdate.GotFocus
 With mskCdate
      .BringToFront()
      txtCdate.SendToBack()
      .Focus()
 End With
End Sub

Private Sub txtCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCdate.KeyPress
 If e.KeyChar = Chr(13) Then
    txtBlockID.Focus()
 End If
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
                        txtMaterial.Focus()
                    End If
                Case Is = 38 '�١�â��
                    txtThk.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtBlockID.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtBlockID.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtBlockID.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub mskCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskCdate.KeyPress
 If e.KeyChar = Chr(13) Then
    txtBlockID.Focus()
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

Private Sub txtBlockID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBlockID.KeyDown
 Dim intChkPoint As Integer

        With txtBlockID
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

                    If .SelectionStart = 0 Then
                        txtCdate.Focus()
                    End If

                Case Is = 38 '�����١�â��    
                        txtCdate.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        txtSize.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtSize.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                          txtSetQty.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtBlockID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBlockID.KeyPress
 If e.KeyChar = Chr(13) Then
    txtSize.Focus()
 End If
End Sub

Private Sub txtBlockID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBlockID.LostFocus
 txtBlockID.Text = txtBlockID.Text.ToUpper.Trim
End Sub

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown
  Dim intChkPoint As Integer

        With txtSize
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

                    If .SelectionStart = 0 Then
                        txtBlockID.Focus()
                    End If

                Case Is = 38 '�����١�â��    
                     txtCdate.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        txtPart.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtPart.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtPrice.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress
  If e.KeyChar = Chr(13) Then
     txtPart.Focus()
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
                        txtPart.Focus()
                    End If

                Case Is = 38 '�����١�â��    
                        txtCdate.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        txtw.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtw.Focus()
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
    txtw.Focus()
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
                        mskLg.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtSize.Focus()

                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        mskSizeQty.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     '��� InChkPoint = ������Ǣͧ  mskLdate

                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            mskSizeQty.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ
                    txtRmk.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

End Sub

Private Sub txtw_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtw.GotFocus
 With mskw
      .BringToFront()
      txtw.SendToBack()
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
                        txtSizeDesc.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtBlockID.Focus()
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

  Select Case Asc(e.KeyChar)
            Case 48 To 57            ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13
                e.Handled = False
                txtLg.Focus()
            Case 8                   ' ���� Backspace
                e.Handled = False
            Case 32                   ' ���� Tab
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtw.Focus()
  End Select

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
      .BringToFront()
      txtLg.SendToBack()
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
                    txtBlockID.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSetQty.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtSetQty.Focus()
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
  Select Case Asc(e.KeyChar)

            Case 48 To 57            ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13                  '���� Enter
                e.Handled = False
                txtSetQty.Focus()
            Case 8                   ' ���� Backspace
                e.Handled = False
            Case 32                   ' ���� Tab
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtLg.Focus()
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
                    txtPart.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtPr.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtPr.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ
                    txtPr.Focus()
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
                txtPr.Focus()
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
                        txtPr.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtPr.Focus()
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

Private Sub txtPart_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPart.KeyDown
 Dim intChkPoint As Integer

      With txtPart
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtSize.Focus()
                    End If
                Case Is = 38   '�����١�â��
                    txtCdate.Focus()
                Case Is = 39   '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtSizeDesc.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length         '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
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

Private Sub txtPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPart.KeyPress
 If e.KeyChar = Chr(13) Then
    txtSizeDesc.Focus()
 End If
End Sub

Private Sub txtPart_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPart.LostFocus
 txtPart.Text = txtPart.Text.ToUpper.Trim
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
                    typePic = Mid(strFileName, lengTypePic + 1, 4)                  ' �Ѵ��� .jpg .png .gif 
                    lengPic = strFileName.Length - 4                                '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
                    strNamePic = Mid(strFileName, 1, lengPic)                       '�Ѵ���੾�Ъ����ٻ
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ

                    lblPicPath1.Text = strFileFullPath
                    lblPicName1.Text = strFileName             '�����������С�ȵ�������С���
                End If

           Catch ex As Exception
                 ClearBlankPicture1()
           End Try

        End With
        CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
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

Private Sub ClearBlankPicture3()
  picEqp3.Image = Nothing
  lblPicPath3.Text = ""
  lblPicName3.Text = ""
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
                    'Load ������ picturebox
                    strFileName = New System.IO.FileInfo(.FileName).Name '�Ѻ���੾�Ъ������
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) '�Ѻ���੾�оҸ���

                    img = ScaleImage(Image.FromFile(.FileName), picEqp2.Height, picEqp2.Width)
                    picEqp2.Image = img

                    '----------- �����ѹ������ҵ�ͪ������ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4)                  ' �Ѵ��� .jpg .png .gif 
                    lengPic = strFileName.Length - 4                                '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
                    strNamePic = Mid(strFileName, 1, lengPic)                       '�Ѵ���੾�Ъ����ٻ
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ

                    lblPicPath2.Text = strFileFullPath
                    lblPicName2.Text = strFileName             '�����������С�ȵ�������С���

                End If

            Catch ex As Exception
                  ClearBlankPicture2()
            End Try

        End With
        CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
End Sub

Private Sub btnDelEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp2.Click
   picEqp2.Image = Nothing
   lblPicPath2.Text = ""
   lblPicName2.Text = ""
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
                    ' Load ������ picturebox
                    strFileName = New System.IO.FileInfo(.FileName).Name '�Ѻ���੾�Ъ������
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) '�Ѻ���੾�оҸ���

                    img = ScaleImage(Image.FromFile(.FileName), picEqp3.Height, picEqp3.Width)
                    picEqp3.Image = img

                    '----------- �����ѹ������ҵ�ͪ������ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4)                  ' �Ѵ��� .jpg .png .gif 
                    lengPic = strFileName.Length - 4                                '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
                    strNamePic = Mid(strFileName, 1, lengPic)                       '�Ѵ���੾�Ъ����ٻ
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ

                    lblPicPath3.Text = strFileFullPath
                    lblPicName3.Text = strFileName             '�����������С�ȵ�������С���

                End If

            Catch ex As Exception
                  ClearBlankPicture3()
            End Try

        End With
        CallEditData()   '������¡�����ŷҧ෤�Ԥ������� 
End Sub

Private Sub btnDelEqp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp3.Click
  picEqp3.Image = Nothing
  lblPicPath3.Text = ""
  lblPicName3.Text = ""
End Sub

Private Sub txtSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeQty.GotFocus
 With mskSizeQty
      .BringToFront()
      txtSizeQty.SendToBack()
      .Focus()
 End With
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
                          txtPart.Focus()

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
                  Case Is = 38   '�����١�â��
                       txtSetQty.Focus()
                  Case Is = 39   '�����١�â��
                       If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                          txtRmk.Focus()
                       Else
                           intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                           If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                              txtRmk.Focus()
                           End If
                      End If
                Case Is = 40 '����ŧ
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
             txtRmk.Focus()

      Case Else
             e.Handled = True
             MessageBox.Show("��س��кآ������������ѧ��� - ����Ţ��ҹ��")
  End Select

End Sub

Private Sub txtPr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPr.LostFocus
  txtPr.Text = txtPr.Text.ToUpper.Trim
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
   tt.Hide(picEqp3)
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


End Class