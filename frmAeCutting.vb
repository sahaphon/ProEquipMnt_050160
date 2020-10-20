Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeCutting
  Dim IsShowSeek As Boolean   '������ʶҹ� gpbSeek
  Dim strDateDefault As String     '���������Ѻ�ѹ�������

  Public Const DrvName As String = "\\10.32.0.15\data1\EquipPicture\"
  Public Const PthName As String = "\\10.32.0.15\data1\EquipPicture"         '���������Ѻ�� part �ٻ�Ҿ 

  Private tt As ToolTip = New ToolTip '�ʴ����ŷԻ ��ٻ�Ҿ��������͹��������

Protected Overrides ReadOnly Property CreateParams() As CreateParams          '��ͧ�ѹ��ûԴ������� Close Button(�����ҡ�ҷ)
    Get
        Dim cp As CreateParams = MyBase.CreateParams
            Const CS_DBLCLKS As Int32 = &H8
            Const CS_NOCLOSE As Int32 = &H200
            cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
            Return cp
    End Get
End Property

Private Sub frmAeCutting_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearTmpTable(0, "")
  frmCutting.lblCmd.Text = "0"  '������ʶҹ�
  Me.Dispose()
End Sub

'������������ table tmp_eqptrn
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

Private Sub frmAeCutting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Dim dtComputer As Date = Now
  Dim strCurrentDate As String

      StdDateTimeThai()        '�Ѻ�ٷչ��ŧ�ѹ�����Ǵ�.�� ����� Module    * ��� Control �ʴ� Datetime �繻վط�
      strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")
      ClearDataGpbHead()
      PreTypeSeek()            '��Ŵ��������´���� cbo ������ǹ����Ե
      PreCutTypeSeek()          '��Ŵ��¡���մ�Ѵ

      Select Case frmCutting.lblCmd.Text.ToString

             Case Is = "0"   '�ó�����������
                  With txtBegin
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
                 'dgvSize.Columns(4).Visible = False  '��͹��������� 5

             Case Is = "1" '�ó���䢢�����

                  With Me
                       .Text = "���䢢�����"
                  End With

                  LockEditData()
                  txtEqp_id.ReadOnly = True   '�����ҹ���ҧ����
                  txtEqpnm.ReadOnly = True
                  txtShoe.ReadOnly = True
                  txtOrder.ReadOnly = True
                  txtRemark.ReadOnly = True

             Case Is = "2"

                  With Me
                       .Text = "����ͧ"
                  End With

                  LockEditData()
                  txtEqp_id.ReadOnly = True  '�����ҹ���ҧ����
                  btnSaveData.Enabled = False

       End Select

End Sub

Private Sub LockEditData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlSelc As String = ""

  Dim strCmd As String = ""
  Dim strLoadFilePicture As String     '�纤��ʵ�ԧ��Ŵ Picture
  Dim strPartPicture As String = "\\10.32.0.15\data1\EquipPicture\"   '�� part

  Dim blnHaveData As Boolean
  Dim strPart As String = ""
  Dim strCode As String = frmCutting.dgvShoe.Rows(frmCutting.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim

      With Conn
           If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      strSqlSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                           & " WHERE eqp_id = '" & strCode & "' "

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

              lblPicName1.Text = .Fields("pic_ctain").Value.ToString.Trim
              lblPicName2.Text = .Fields("pic_part").Value.ToString.Trim
              lblPicPath1.Text = PthName
              lblPicPath2.Text = PthName


              '-------------------------------Load �ٻ�Ҿ(�ٻ�մ�Ѵ)-----------------------

              strLoadFilePicture = strPartPicture & .Fields("pic_io").Value.ToString.Trim
              If File.Exists(strLoadFilePicture) Then
                 Dim img1 As Image           '��С�ȵ���� img1 �������Ҿ
                 img1 = Image.FromFile(strLoadFilePicture) 'img1 ��ҡѺpicture �����Ŵ�Ҩҡ db
                 picEqp1.Image = ScaleImage(img1, picEqp1.Height, picEqp1.Width)
              Else
                 picEqp1.Image = Nothing
              End If
              strLoadFilePicture = ""

               '-------------------------------Load �ٻ�Ҿ��Ե��� -----------------------
              strLoadFilePicture = strPartPicture & .Fields("pic_part").Value.ToString.Trim
              If File.Exists(strLoadFilePicture) Then
                 Dim img2 As Image
                 img2 = Image.FromFile(strLoadFilePicture)
                 picEqp2.Image = ScaleImage(img2, picEqp2.Height, picEqp2.Width)
              Else
                 picEqp2.Image = Nothing
              End If
              strLoadFilePicture = ""


              strCmd = frmCutting.lblCmd.Text.ToString.Trim    '��� strCmd ��ҡѺ���� lblcmd 㹿���� frmEqpSheet

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

              '-------------------------------------------------------------------------------------

              strSqlSelc = "INSERT INTO tmp_eqptrn " _
                                     & " SELECT user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                     & " FROM eqptrn " _
                                     & " WHERE eqp_id = '" & strCode & "'"

              Conn.Execute(strSqlSelc)

              blnHaveData = True  ' �բ�����

           Else
              blnHaveData = False  '����բ�����

           End If
      .ActiveConnection = Nothing
      .Close()
      End With

Conn.Close()
Conn = Nothing  '������ Connection
     If blnHaveData Then          '��� blnHavedata = true
        ShowScrapItem()
     End If
End Sub


Private Sub ShowScrapItem()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim sta As String = ""    '�纤�� status
 Dim dubQty As Double
 Dim dubAmt As Double
 Dim sngSetQty As Single  '�纨ӹǹ SET
 Dim user As String = frmMainPro.lblLogin.Text.ToString.Trim
 Dim mold_id As String
 Dim mold_size As String
 Dim strArr() As String

      With Conn
           If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      strSqlSelc = "SELECT * " _
                                 & "FROM v_tmp_eqptrn (NOLOCK)" _
                                 & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "' " _
                                 & "ORDER BY size_desc, size_id"


      With Rsd

           .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .Open(strSqlSelc, Conn, , , )


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
                                    .Fields("cut_id").Value.ToString.Trim, _
                                    .Fields("size_desc").Value.ToString.Trim, _
                                    .Fields("cut_detail").Value.ToString.Trim, _
                                    .Fields("backgup").Value.ToString.Trim, _
                                    Format(.Fields("set_qty").Value, "#0.0"), _
                                    Format(.Fields("size_qty").Value, "#0.0"), _
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

              sngSetQty = sngSetQty + .Fields("set_qty").Value       '���ἧ
              dubQty = dubQty + .Fields("ord_qty").Value             '�ӹǹ��� ŧ��Ե
              dubAmt = dubAmt + .Fields("price").Value               '�����Ť���ػ�ó�

              .MoveNext()
              Loop

               txtSet.Text = sngSetQty.ToString.Trim
               txtAmount.Text = Format(dubQty, "#,##0")
               lblAmt.Text = Format(dubAmt, "#,##0.00")

           Else
               txtSet.Text = "0.0"
               txtAmount.Text = "0"
               lblAmt.Text = "0.00"
           End If
           .ActiveConnection = Nothing
           .Close()
           Rsd = Nothing

           dgvSize.ScrollBars = ScrollBars.Both       '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ
      End With

Conn.Close()
Conn = Nothing
End Sub

Private Sub ClearDataGpbHead()
  txtEqp_id.Text = ""
  txtEqpnm.Text = ""
  txtShoe.Text = ""
  txtOrder.Text = ""
  txtSet.Text = ""
  txtRemark.Text = ""
End Sub

Private Sub PreTypeSeek()
  Dim strGpbSeek(4) As String
  Dim i As Integer

      strGpbSeek(0) = "US"
      strGpbSeek(1) = "UW"
      strGpbSeek(2) = "UY"
      strGpbSeek(3) = "UV"
      strGpbSeek(4) = "UB"

      For i = 0 To 4
          cmbTMaterial.Items.Add(strGpbSeek(i))
      Next i
End Sub

Private Sub PreCutTypeSeek()    '��Ŵ��������´���� Combo �մ�Ѵ
   Dim strCutTopic(3) As String
   Dim i As Byte

     strCutTopic(0) = "�մἧ"
     strCutTopic(1) = "�մ 2.5 x 19 MM"
     strCutTopic(2) = "�մ����"
     strCutTopic(3) = "�����"

        With cboCutdetail

            For i = 0 To 3
                .Items.Add(strCutTopic(i))
            Next i

        End With
End Sub

Private Sub txtEqp_id_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqp_id.KeyDown
 Dim intChkPoint As Integer
 With txtEqp_id
     Select Case e.KeyCode
            Case Is = 35 '���� End 
            Case Is = 36 '���� Home
            Case Is = 37 '�١�ë���
                 If .SelectionStart = 0 Then
                   End If
            Case Is = 38 '�����١�â��

            Case Is = 39   '�����١�â��
                 If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                    txtEqpnm.Focus()
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
                    txtShoe.Focus()
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
                       txtOrder.Focus()
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

Private Sub txtShoe_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShoe.KeyPress
   If e.KeyChar = Chr(13) Then
       txtOrder.Focus()
   End If
End Sub

Private Sub txtEqp_id_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqp_id.LostFocus
  txtEqp_id.Text = txtEqp_id.Text.ToString.ToUpper.Trim()
End Sub

Private Sub txtEqpnm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpnm.LostFocus
  txtEqpnm.Text = txtEqpnm.Text.ToString.ToUpper.Trim()
End Sub

Private Sub txtShoe_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShoe.LostFocus
  txtShoe.Text = txtShoe.Text.ToString.ToUpper.Trim()
End Sub

Private Sub txtOrder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrder.LostFocus
  txtOrder.Text = txtOrder.Text.ToString.ToUpper.Trim()
End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBfSave()
End Sub

Private Sub CheckDataBfSave()

  Dim IntListwc As Integer = dgvSize.Rows.Count
  Dim strProd As String = ""
  Dim strProdnm As String = ""

  Dim bytConSave As Byte

  If txtEqp_id.Text <> "" Then

        If txtEqpnm.Text <> "" Then

               If IntListwc > 0 Then

                           bytConSave = MsgBox("�س��ͧ��úѹ�֡���������������!" _
                                  , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                  If bytConSave = 6 Then

                                       Select Case Me.Text
                                              Case Is = "����������"

                                                  If CheckCodeDuplicate() Then   '�������ػ�ó���
                                                     SaveNewRecord()

                                                  Else
                                                     MessageBox.Show("�����ػ�ó��� ��سҡ�͡�����ػ�ó�����!....", _
                                                                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

                         ShowResvrd()       '�ʴ���������� gpbSeek 
                         gpbSeek.Text = "����������"

                         If CheckCodeDuplicate() Then
                            txtSize.ReadOnly = False

                         Else
                              MessageBox.Show("�����ػ�ó��� ��سҡ�͡�����ػ�ó�����!....", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                              txtEqp_id.Text = ""
                              txtEqp_id.Focus()

                         End If

                   End If

        Else
             MsgBox("�ô�кآ�������������´�ػ�ó�  " _
                          & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
             txtEqp_id.Focus()

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

  Dim strDocdate As String           '��ʵ�ԧ�ѹ����͡���
  Dim strGpType As String = ""       '�纻������ػ�ó�
  Dim strPartType As String = ""     '�纪����ǹ����Ե
  Dim strNull As String              '�纤����ҧ NULL
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

               strNull = "NULL"

                      '------------------------------------ �ѹ�֡�ٻ�մ ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath1.Text.ToString.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath1.Text = PthName

                       Else
                          lblPicName1.Text = ""
                          lblPicPath1.Text = ""
                          picEqp1.Image = Nothing

                       End If

                      '------------------------------------ �ѹ�֡�ٻ��鹧ҹ ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath2.Text.ToString.Trim, ReturnImageName(lblPicName2.Text.ToString.Trim), lblPicName2.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath2.Text = PthName

                       Else
                          lblPicName2.Text = ""
                          lblPicPath2.Text = ""
                          picEqp2.Image = Nothing

                       End If

                      '---------------------------------- UPDATE ������㹵��ҧ eqpmst ----------------------------------------------

                      strSqlCmd = "UPDATE  eqpmst SET eqp_name ='" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                 & "," & "pi ='" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                                 & "," & "shoe ='" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                                 & "," & "part ='" & strPartType & "'" _
                                 & "," & "eqp_type ='" & "LCA" & "'" _
                                 & "," & "set_qty =" & ChangFormat(txtSet.Text.ToString.Trim) _
                                 & "," & "pic_ctain ='" & "" & "'" _
                                 & "," & "pic_io ='" & ReplaceQuote(lblPicName1.Text.ToString.Trim) & "'" _
                                 & "," & "pic_part ='" & ReplaceQuote(lblPicName2.Text.ToString.Trim) & "'" _
                                 & "," & "remark ='" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                 & "," & "tech_desc = '" & "" & "'" _
                                 & "," & "tech_thk = '" & "" & "'" _
                                 & "," & "tech_lg = '" & "" & "'" _
                                 & "," & "tech_sht = '" & "" & "'" _
                                 & "," & "tech_eva = '" & "" & "'" _
                                 & "," & "tech_warm = '" & "" & "'" _
                                 & "," & "tech_time1 = '" & "" & "'" _
                                 & "," & "tech_time2 = '" & "" & "'" _
                                 & "," & "creat_date = " & strNull _
                                 & "," & "eqp_amt = " & RetrnAmount() _
                                 & "," & "last_date = '" & strDate & "'" _
                                 & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                 & "," & "exp_id ='" & "" & "'" _
                                 & "," & "tech_trait ='" & "" & " '" _
                                 & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)


                     '------------------------------------------------ź������㹵��ҧ eqptrn----------------------------------------------------------------------

                     strSqlCmd = "Delete FROM eqptrn" _
                                          & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlCmd)

                     '----------------------------------------- �ѹ�֡������㹵��ҧ eqptrn �� Select �ҡ tmp_eqptrn ------------------------------------------------

        strSqlCmd = "INSERT INTO eqptrn " _
                        & "SELECT [group] = 'E'" _
                        & ",eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'" _
                        & ",size_id,size_desc,size_qty,weight,dimns,backgup" _
                        & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date" _
                        & ",pr_doc,recv_date,ord_rep,ord_qty,fc_date,impt_id" _
                        & ",sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                        & " FROM tmp_eqptrn " _
                        & " WHERE user_id= '" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

        Conn.Execute(strSqlCmd)
        Conn.CommitTrans()  '��� Commit transection

        frmCutting.lblCmd.Text = txtEqp_id.Text.ToString.Trim   '�觺͡��Һѹ�֡�����������
        frmCutting.Activating()
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
   Dim strPRdate As String
   Dim strDateDoc, strINdate As String
   Dim strFCdate As String
   Dim strPartType As String = ""
   Dim Rsd As New ADODB.Recordset

   Dim strNull As String     '�纤����ҧ

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

                    '---------------------------------------- �ѹ�֡�ٻ�մ�Ѵ ----------------------------------------------------
                    blnRetuneCopyPic = CallCopyPicture(lblPicPath1.Text.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)

                    If blnRetuneCopyPic Then       '��� CallCopyPicture = true
                       lblPicPath1.Text = PthName
                    Else
                       lblPicPath1.Text = ""
                       lblPicName1.Text = ""
                       picEqp1 = Nothing

                    End If

                    '---------------------------------------- �ѹ�֡�ٻ��鹧ҹ -------------------------------------------------

                     blnRetuneCopyPic = CallCopyPicture(lblPicPath2.Text.Trim, ReturnImageName(lblPicName2.Text.Trim), lblPicName2.Text.Trim)

                    If blnRetuneCopyPic Then
                       lblPicPath2.Text = PthName

                    Else
                       lblPicPath2.Text = ""
                       lblPicName2.Text = ""
                       picEqp2 = Nothing
                    End If


                   strNull = "NULL"


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


                '-------------------------------- INSERT ������㹵��ҧ eqpmst --------------------------------------

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
                      & ",'" & "E" & "'" _
                      & ",'" & ReplaceQuote(txtEqp_id.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtEqpnm.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(txtSupplier.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ChangFormat(txtSet.Text.ToString.Trim) & "'" _
                      & ",'" & strPartType & "'" _
                      & ",'" & "LCA" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(lblPicName1.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(lblPicName2.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & " '" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & " '" _
                      & "," & strNull _
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
                                     & " SELECT [group] ='E'" _
                                     & ",eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'" _
                                     & ",size_id,size_desc,size_qty,weight,dimns,backgup,price,men_rmk" _
                                     & ",delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                     & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & " FROM tmp_eqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                Conn.Execute(strSqlCmd)
                Conn.CommitTrans()

                frmCutting.lblCmd.Text = txtEqp_id.Text.ToString.Trim   '�觺͡��Һѹ�֡�����������
                frmCutting.Activating()
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

        '---------------------------------- ����� SELCT SUM()AS ����������� ---------------------------------------

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

 '----------------------------- �ѧ���� CopyPicture ---------------------------------------------
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

Private Function CheckCodeDuplicate() As Boolean
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
                             & " WHERE eqp_id = '" & txtEqp_id.Text & "'"

      With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

           If .RecordCount > 0 Then
               CheckCodeDuplicate = False

           Else
               CheckCodeDuplicate = True
           End If

      .ActiveConnection = Nothing
      .Close()
      End With

Conn.Close()
Conn = Nothing

End Function

Private Sub ShowResvrd()
  tabMain.SelectedTab = tabSize  '�ʴ� TabSize
  IsShowSeek = Not IsShowSeek    '�ҡ  IsShowSeek �� False �������¹�� True

  If IsShowSeek Then

     With gpbSeek
          .Visible = True
          .Left = 8    '᡹ X
          .Top = 230   '᡹ Y 
          .Height = 500
          .Width = 990
     End With

      StateLockFindDept(False)                '�Ѻ�ٷչ��ͤ Control

  Else
      StateLockFindDept(True)
  End If

End Sub

Private Sub ShowResvrdEdit()
  tabMain.SelectedTab = tabSize  '�ʴ� TabSize
  IsShowSeek = Not IsShowSeek    '�ҡ  IsShowSeek �� False �������¹�� True

  If IsShowSeek Then

     With gpbSeek
          .Visible = True
          .Left = 8    '᡹ X
          .Top = 230   '᡹ Y 
          .Height = 500
          .Width = 990
     End With

      StateLockFindDept(False)                '�Ѻ�ٷչ��ͤ Control

  Else
      StateLockFindDept(True)
  End If

End Sub

Private Sub StateLockAEItem(ByVal sta As String)

  cmbTMaterial.Enabled = sta
  cboCutdetail.Enabled = sta

End Sub

Private Sub StateLockFindDept(ByVal sta As String)
 Dim strMod As String = frmCutting.lblCmd.Text.ToString

     btnAdd.Enabled = sta
     gpbHead.Enabled = sta

     tabMain.Enabled = sta
     btnSaveData.Enabled = sta

     Select Case strMod
            Case Is = "1"   '��䢢����� 
            Case Is = "2"   '����ͧ������
                  btnSaveData.Enabled = False
     End Select

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  On Error Resume Next    '�ҡ�Դ error ������ѧ�зӧҹ���������ʹ� error ����Դ���
  Dim strCode As String

     If MessageBox.Show("�س��ͧ����͡�ҡ����� ���������", "��س��׹�ѹ�͡�ҡ�����", MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                                                                         = Windows.Forms.DialogResult.Yes Then
        With frmCutting.dgvShoe
             If .Rows.Count > 0 Then
                 strCode = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim          '���strCode = ��������ǻѨ�غѹ Cell �á
                 lblComplete.Text = strCode  '��� label �ʴ�������� Cell �Ѩ�غѹ   
             End If
        End With
        Me.Close()

        frmMainPro.Show()
        frmCutting.Show()
     End If
End Sub

Private Sub btnSeekExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekExit.Click
    StateLockFindDept(True)
    gpbSeek.Text = ""
    gpbSeek.Visible = False
    IsShowSeek = False
End Sub

Private Sub ClearAllData()
 txtCutID.Text = ""
 txtSize.Text = ""
 txtPart.Text = ""
 txtSizeDesc.Text = ""
 txtSizeQty.Text = "0"
 txtSetQty.Text = "0"
 txtPrice.Text = "0.00"
 txtPr.Text = ""
 txtPrdate.Text = "__/__/____"
 txtFCdate.Text = "__/__/____"
 txtIndate.Text = "__/__/____"
 txtSupplier.Text = ""
 txtRmk.Text = ""
End Sub

Private Sub btnSeekSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekSave.Click
  CheckSubDataBfSave()
End Sub

Private Sub CheckSubDataBfSave()
Dim i As Integer

 If cmbTMaterial.Text <> "" Then

      If txtCutID.Text <> "" Then

              If txtSize.Text <> "" Then

                     If gpbSeek.Text = "����������" Then
                        SaveSubRecord()
                     Else
                        EditSubRecord()
                     End If

                     ShowScrapItem()   '�ʴ�����ŷ���ѹ��� dgvSize �� Select �ҡ v_tmp_eqptrn

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
                   MsgBox("�ô��͡�����������մ  " _
                          & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                   txtSize.Focus()

              End If

       Else
            MsgBox("�ô��͡�����������մ  " _
                          & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
            txtCutID.Focus()
       End If


 Else
      MsgBox("�ô���͡�����Ż������ѵ�شԺ  " _
                      & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")

      cmbTMaterial.DroppedDown = True
      cmbTMaterial.Focus()
 End If
End Sub

Private Function SaveSubRecord() As Boolean

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlSelec As String = ""
  Dim strCmd As String = ""

  Dim dateSave As Date = Now()
  Dim strDate As String
  Dim strEngYear As String
  Dim strDateDoc As String    '�ѹ���͡���

  Dim strPrdate As String   '�ѹ����Դ���觫���
  Dim strFcDate As String   '�ѹ���Ѵ���
  Dim strIndate As String   '�ѹ�����Ѻ���

  Dim strCutType As String = ""
  Dim strMateType As String = ""

     Try

      With Conn
           If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

     '------------------------------------�礢�������͹�����������������-------------------------------------------------

     strSqlSelec = "SELECT size_id FROM tmp_eqptrn " _
                      & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                      & " AND size_id = '" & txtSize.Text.ToString.Trim & "'" _
                      & " AND cut_id = '" & txtCutID.Text.ToString.Trim & "'" _
                      & " AND size_desc = '" & txtSizeDesc.Text.ToString.Trim & "'"


    With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelec, Conn, , , )

         If .RecordCount <> 0 Then
             MessageBox.Show("Size : " & txtSize.Text.ToString & _
                                                 "����к����� ��س��к� Size ����", "�����ū��!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
             SaveSubRecord = False

         Else

             strDate = dateSave.ToString("yyyy-MM-dd")
             strEngYear = SaveChangeEngYear(strDate)

             strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtBegin.Text.ToString, 1, 2)
             strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"

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

                 '----------------------- �������ѵ�شԺ ----------------------------------------

                 Select Case cmbTMaterial.Text.ToString.Trim

                    Case Is = "US"
                        strMateType = "US"
                    Case Is = "UW"
                        strMateType = "UW"
                    Case Is = "UY"
                        strMateType = "UY"
                    Case Is = "UV"
                        strMateType = "UV"
                    Case Is = "UB"
                        strMateType = "UB"
                 End Select

                 '--------------------------- ��¡���մ�Ѵ -------------------------------------

                 Select Case cboCutdetail.SelectedIndex

                    Case Is = 0
                        strCutType = "�մἧ"
                    Case Is = 1
                        strCutType = "�մ 2.5 x 19 MM"
                    Case Is = 2
                        strCutType = "�մ����"
                    Case Is = 3
                        strCutType = "�����"

                 End Select


             strCmd = "INSERT INTO tmp_eqptrn " _
                                & "(user_id,size_id,size_desc,size_qty,set_qty" _
                                & ",dimns,backgup,price,men_rmk,[group],eqp_id" _
                                & ",delvr_sta,sent_sta,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                & ",fc_date,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                & ")" _
                                & " VALUES (" _
                                & "'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                & ",'" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                & ",'" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                & "," & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                & "," & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                & ",'" & "0.00" & " x " & _
                                         "0.00" & " '" _
                                & ",'" & ReplaceQuote(txtPart.Text.ToString.Trim) & "'" _
                                & "," & ChangFormat(txtPrice.Text.ToString.Trim) _
                                & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                & ",'" & "E" & "'" _
                                & ",'" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
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
                                & ",'" & "" & "'" _
                                & ",'" & ReplaceQuote(txtCutID.Text.ToString.Trim) & "'" _
                                & ",'" & strMateType & "'" _
                                & ",'" & strCutType & "'" _
                                & "," & 0.0 _
                                & ")"

              Conn.Execute(strCmd)
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

Private Function EditSubRecord() As Boolean

 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim strPrDate As String = ""
 Dim strRecvDate As String = ""
 Dim strFcDate As String = ""

 Dim strCutType As String = ""
 Dim strMateType As String = ""

     Try

      With Conn
         If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

      End With

             '----------------------------------------�ѹ����Դ����---------------------------------------------------

             If txtPrdate.Text <> "__/__/____" Then

                strPrDate = Mid(txtPrdate.Text.ToString, 7, 4) & "-" _
                                & Mid(txtPrdate.Text.ToString, 4, 2) & "-" _
                                & Mid(txtPrdate.Text.ToString, 1, 2)
                                strPrDate = "'" & SaveChangeEngYear(strPrDate) & "'"

            Else
                strPrDate = "NULL"
            End If

           '----------------------------------------�ѹ����Ѻ���---------------------------------------------------

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

            '----------------------- �������ѵ�شԺ ---------------------------------------------

                 Select Case cmbTMaterial.Text.ToString.Trim

                    Case Is = "US"
                        strMateType = "US"
                    Case Is = "UW"
                        strMateType = "UW"
                    Case Is = "UY"
                        strMateType = "UY"
                    Case Is = "UV"
                        strMateType = "UV"
                    Case Is = "UB"
                        strMateType = "UB"
                 End Select

             '------------------------- ��¡���մ�Ѵ -----------------------------------------------

              Select Case cboCutdetail.SelectedIndex

                    Case Is = 0
                        strCutType = "�մἧ"
                    Case Is = 1
                        strCutType = "�մ 2.5 x 19 MM"
                    Case Is = 2
                        strCutType = "�մ����"
                    Case Is = 3
                        strCutType = "�����"

              End Select

          '------------------------------------��˹������������---------------------------------------------------------------------


             strSqlCmd = "UPDATE  tmp_eqptrn SET size_desc ='" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                            & "," & "size_qty =" & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                            & "," & "[group]= '" & "E" & "'" _
                            & "," & "dimns ='" & "0.00" & " x " & _
                                                 "0.00" & "'" _
                            & "," & "backgup = '" & ReplaceQuote(txtPart.Text.ToString.Trim) & "'" _
                            & "," & "price = " & ChangFormat(txtPrice.Text.ToString.Trim) _
                            & "," & "pr_doc ='" & ReplaceQuote(txtPr.Text.ToString.Trim) & "'" _
                            & "," & "men_rmk ='" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                            & "," & "set_qty =" & ChangFormat(txtSetQty.Text.ToString.Trim) _
                            & "," & "pr_date = " & strPrDate _
                            & "," & "recv_date = " & strRecvDate _
                            & "," & "fc_date = " & strFcDate _
                            & "," & "sup_name = '" & ReplaceQuote(txtSupplier.Text.ToString.Trim) & "'" _
                            & "," & "lp_type = '" & "LCA" & "'" _
                            & "," & "size_group = '" & "" & "'" _
                            & "," & "cut_id = '" & ReplaceQuote(txtCutID.Text.ToString.Trim) & "'" _
                            & "," & "mate_type = '" & strMateType & "'" _
                            & "," & "cut_detail = '" & strCutType & "'" _
                            & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                            & " AND size_id ='" & txtSize.Text.ToString.Trim & "'" _
                            & " AND cut_id = '" & txtCutID.Text.ToString.Trim & "'" _
                            & " AND size_desc = '" & txtSizeDesc.Text.ToString.Trim & "'"


       Conn.Execute(strSqlCmd)


  Conn.Close()
  Conn = Nothing

     Catch ex As Exception
           MsgBox("����ͼԴ��Ҵ��зӡ�úѹ�֡ �ô���Թ��������ա����", MsgBoxStyle.Critical, "�Դ��Ҵ")
           MsgBox(ex.Message)
     End Try

End Function


Private Sub txtCutID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCutID.KeyDown
  Dim intChkPoint As Integer
        With txtCutID
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                Case Is = 38 '�����١�â��     
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                       cmbTMaterial.DroppedDown = True
                       cmbTMaterial.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           cmbTMaterial.DroppedDown = True
                           cmbTMaterial.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtPart.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtCutID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCutID.KeyPress
   If e.KeyChar = Chr(13) Then
      cmbTMaterial.DroppedDown = True
      cmbTMaterial.Focus()

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
                         txtPart.Focus()
                     End If
                Case Is = 38 '�����١�â��     
                     txtCutID.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                       txtSizeDesc.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtSizeDesc.Focus()
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
                     If .SelectionStart = 0 Then
                         txtCutID.Focus()
                     End If

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
                    txtPr.Focus()
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

Private Sub txtSetQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSetQty.GotFocus
  With mskSetQty
       txtSetQty.SendToBack()
       .BringToFront()
       .Focus()
  End With
End Sub

Private Sub txtSetQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSetQty.KeyDown
 Dim intChkPoint As Integer
        With txtSetQty
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                     If .SelectionStart = 0 Then
                         txtSizeDesc.Focus()
                     End If
                Case Is = 38 '�����١�â��   
                     If .SelectionStart = 0 Then
                         txtCutID.Focus()
                     End If

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                       txtSizeQty.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtSizeQty.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtPrdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSetQty.KeyPress
   If e.KeyChar = Chr(13) Then
       txtSizeQty.Focus()
   End If
End Sub

Private Sub txtSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeQty.GotFocus
  With mskSizeQty
       txtSizeQty.SendToBack()
       .BringToFront()
       .Focus()
  End With
End Sub

Private Sub txtSizeQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSizeQty.KeyDown
  Dim intChkPoint As Integer
        With txtSizeQty
            Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                     If .SelectionStart = 0 Then
                         txtSetQty.Focus()
                     End If
                Case Is = 38 '�����١�â��   
                      If .SelectionStart = 0 Then
                         txtCutID.Focus()
                     End If

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
                    txtFCdate.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSizeQty.KeyPress
   If e.KeyChar = Chr(13) Then
       txtPrice.Focus()
   End If
End Sub

Private Sub txtPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrice.GotFocus

  With mskPrice
       txtPrice.SendToBack()
      .BringToFront()
      .Focus()
  End With

End Sub

Private Sub txtPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPr.Focus()
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
                            txtPrice.Focus()
                        End If
                   Case Is = 38 '�����١�â��     

                        If .SelectionStart = 0 Then
                           txtSize.Focus()
                        End If

                   Case Is = 39 '�����١�â��

                        If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                           txtPrdate.Focus()
                        Else
                          intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                           If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                              txtPrdate.Focus()
                           End If
                        End If

                Case Is = 40 '����ŧ
                    txtSupplier.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub txtPr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPr.KeyPress
   If e.KeyChar = Chr(13) Then
       txtPrdate.Focus()
   End If
End Sub

Private Sub txtPrdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrdate.GotFocus
 With mskPrdate
      txtPrdate.SendToBack()
      .BringToFront()
      .Focus()
 End With
End Sub

Private Sub txtPrdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrdate.KeyPress
 If e.KeyChar = Chr(13) Then
    txtFCdate.Focus()
 End If
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

           Select Case strMerg.ToString.Length        '�Ѻ�ӹǹ strMerg
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
                         txtSize.Focus()
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
                     txtIndate.Focus()
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
          x = .Text.Length     '�Ѻ��Ҥ������ Text

          For i = 1 To x
              strTmp = Mid(.Text.ToString.Trim, i, 1)
              Select Case strTmp
                     Case Is = "+"
                     Case Is = "-"
                     Case Is = ","
                     Case Else
                          If InStr("0123456789/", strTmp) > 0 Then
                             strMerg = strMerg & strTmp
                          End If
              End Select
              strTmp = ""    '�ó������ Case Else
          Next i

          Try   '��

             mskPrdate.Text = ""
             strMerg = "#" & strMerg & "#"
             z = CDate(strMerg)
             txtPrdate.Text = z.ToString("dd/MM/yyyy")

          Catch ex As Exception
                mskPrdate.Text = "__/__/____"
                txtPrdate.Text = "__/__/____"
          End Try
     .SendToBack()
     txtPrdate.BringToFront()

     End With
End Sub

Private Sub txtFCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFCdate.GotFocus
  With mskFCdate
       txtFCdate.SendToBack()
       .BringToFront()
       .Focus()
  End With
End Sub

Private Sub txtFCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFCdate.KeyPress
  If e.KeyChar = Chr(13) Then
     txtIndate.Focus()
  End If
End Sub

Private Sub mskFCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFCdate.GotFocus
  Dim i, x As Integer
  Dim strTmp As String = ""
  Dim strMerg As String = ""

      With mskFCdate
           If txtFCdate.Text <> "__/__/____" Then
              x = Len(txtFCdate.Text)

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
              .SelectedText = strMerg    'mskFCdate = strMerg
           End If
           .SelectAll()        'mskFCdate = ����ѡ�÷���������
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
                           txtPrdate.Focus()
                        End If
                   Case Is = 38   '�����١�â��
                         txtSizeQty.Focus()
                   Case Is = 39   '�����١�â��
                        If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                            txtIndate.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                            If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                               txtIndate.Focus()
                            End If
                        End If
                   Case Is = 40 '����ŧ
                     txtIndate.Focus()
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
 Dim strMerg As String = ""

     With mskFCdate
          x = .Text.Length

          For i = 1 To x
              strTmp = Mid(.Text.ToString.Trim, i, 1)
              Select Case strTmp
                     Case Is = "+"
                     Case Is = "-"
                     Case Is = ","
                     Case Else
                          If InStr("0123456789/", strTmp) > 0 Then
                             strMerg = strMerg & strTmp
                          End If
              End Select
              strTmp = ""         '�ó������� Case Else
          Next i

          Try

             mskFCdate.Text = ""
             strMerg = "#" & strMerg & "#"
             z = CDate(strMerg)
             txtFCdate.Text = z.ToString("dd/MM/yyyy")

          Catch ex As Exception
                mskFCdate.Text = "__/__/____"
                txtFCdate.Text = "__/__/____"
          End Try

     .SendToBack()
     txtFCdate.BringToFront()
     End With

End Sub

Private Sub txtIndate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndate.GotFocus
 With mskIndate
      txtIndate.SendToBack()
      .BringToFront()
      .Focus()
 End With
End Sub

Private Sub txtIndate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIndate.KeyPress
  If e.KeyChar = Chr(13) Then
     txtSupplier.Focus()
  End If
End Sub

Private Sub mskIndate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskIndate.GotFocus
  Dim i, x As Integer

  Dim strTmp As String = ""
  Dim strMerg As String = ""

      With mskIndate
           If txtIndate.Text <> "__/__/____" Then
              x = Len(txtIndate.Text)

              For i = 1 To x
                  strTmp = Mid(txtIndate.Text.Trim, i, 1)
                  Select Case strTmp
                         Case Is = "_"
                         Case Else
                              If InStr("0123456789/", strTmp) > 0 Then        '������ԧ�����ʵ�ԧ��ѡ �¨Ф׹��ҵ��˹觷�辺
                                 strMerg = strMerg & strTmp
                              End If
                  End Select
                  strTmp = ""
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
                         txtPrice.Focus()
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
          x = .Text.Length           '�Ҥ�����ǵ���ѡ��� mskIndate

          For i = 1 To x
              strTmp = Mid(.Text.ToString.Trim, i, 1)
              Select Case strTmp
                     Case Is = "+"
                     Case Is = "-"
                     Case Is = ","
                     Case Else
                          If InStr("0123456789/", strTmp) > 0 Then
                             strMerg = strMerg & strTmp
                          End If
              End Select
              strTmp = ""            '�ó������� Case Else
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

     .SendToBack()
     txtIndate.BringToFront()
     End With

End Sub

Private Sub txtSizeDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeDesc.LostFocus
  txtSizeDesc.Text = txtSizeDesc.Text.ToUpper.Trim
End Sub

Private Sub txtPr_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPr.LostFocus
  txtPr.Text = txtPr.Text.ToUpper.Trim
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
                   Case Is = 38   '�����١�â��
                         txtPr.Focus()
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

Private Sub txtSupplier_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSupplier.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRmk.Focus()
  End If
End Sub

'-------------------------- �Ѻ�ٷ�� AddTxtCutID ---------------------------------------------
Private Sub AddTxtCutID()
 Dim strTypeMate As String = ""

   Select Case cmbTMaterial.Text.ToString.Trim

        Case Is = "US"
              strTypeMate = "US" & txtCutID.Text
        Case Is = "UW"
              strTypeMate = "UW" & txtCutID.Text
        Case Is = "UY"
              strTypeMate = "UY" & txtCutID.Text
        Case Is = "UV"
              strTypeMate = "UV" & txtCutID.Text
        Case Is = "UB"
              strTypeMate = "UB" & txtCutID.Text
   End Select

        Select Case cboCutdetail.Text.ToString.Trim

               Case Is = "�մἧ"
                     txtCutID.Text = strTypeMate & "-1"

               Case Is = "�մ 2.5 x 19 MM"
                     txtCutID.Text = strTypeMate & "-2"

               Case Is = "�մ����"
                     txtCutID.Text = strTypeMate & "-3"

               Case Is = "�����"
                     txtCutID.Text = strTypeMate & "-4"

        End Select

End Sub

Private Sub txtCutnm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
   txtCutID.Text = txtCutID.Text.ToUpper.Trim
End Sub

Private Sub txtCutnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
   If e.KeyChar = Chr(13) Then
      txtSize.Focus()
   End If
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
                         txtCutID.Focus()
                   Case Is = 39   '�����١�â��
                        If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                            txtSizeQty.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                            If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                               txtSizeQty.Focus()
                            End If
                        End If
                   Case Is = 40 '����ŧ
                         txtPrdate.Focus()
                   Case Is = 113 '���� F2
                         .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub mskSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSetQty.KeyPress
   If e.KeyChar = Chr(13) Then
      txtSizeQty.Focus()
   End If
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

Private Sub mskSizeQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSizeQty.KeyDown
  Dim intChkPoint As Integer
  With mskSizeQty
       Select Case e.KeyCode
              Case Is = 35 '���� End 
              Case Is = 36 '���� Home
              Case Is = 37 '�١�ë���
                        If .SelectionStart = 0 Then
                            txtSizeDesc.Focus()
                        End If
              Case Is = 38   '�����١�â��
                         txtCutID.Focus()
              Case Is = 39   '�����١�â��
                        If .SelectionLength = .Text.Trim.Length Then  '��ҵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                            txtSizeQty.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                            If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                               txtSizeQty.Focus()
                            End If
                        End If
             Case Is = 40 '����ŧ
                      txtPrdate.Focus()
             Case Is = 113 '���� F2
                      .SelectionStart = .Text.Trim.Length
       End Select
  End With
End Sub


Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPrice.Focus()
   End If
End Sub

Private Sub mskSizeQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.LostFocus
  Dim i, x As Integer
  Dim z As Double

  Dim strTmp As String = ""
  Dim strMerge As String = ""

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
                txtSizeQty.Text = z.ToString("#,##0.0")


            Catch ex As Exception
                txtSizeQty.Text = "0.0"
                mskSizeQty.Text = ""
            End Try

            mskSizeQty.SendToBack()
            txtSizeQty.BringToFront()

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
                            txtPr.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                            If .SelectionStart = intChkPoint Then   '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                               txtPr.Focus()
                            End If
                        End If
             Case Is = 40 '����ŧ
                      txtIndate.Focus()
             Case Is = 113 '���� F2
                      .SelectionStart = .Text.Trim.Length
       End Select
  End With

End Sub

Private Sub mskPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrice.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPr.Focus()
   End If
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

Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
  If dgvSize.Rows.Count > 0 Then
     ShowResvrdEdit()       '�ʴ� gpbSeek 
     CallEditData()
     StateLockAEItem(False)   ' �Դ�����ҹ Combobox �������ѵ�شԺ, ��¡���մ�Ѵ
     gpbSeek.Text = "��䢢�����"
     txtCutID.ReadOnly = True
     txtSize.ReadOnly = True
     txtSizeDesc.ReadOnly = True
     txtPart.Focus()
  End If
End Sub

Private Sub CallEditData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlSelc As String
  Dim strWd As String = ""   '�纤�ҡ��ҧ
  Dim strLg As String = ""   '�纤������
  Dim strHg As String = ""   '�纤����٧

      If dgvSize.Rows.Count <> 0 Then

         Dim strSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(2).Value.ToString.Trim     'Size
         Dim strCutID As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(4).Value.ToString.Trim     '�����մ
         Dim strGpSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(5).Value.ToString.Trim      ' Group Size

         With Conn

              If .State Then Close()
                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

         End With

            strSqlSelc = " SELECT *  FROM v_tmp_eqptrn (NOLOCK)" _
                          & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                          & " AND size_id= '" & strSize & "'" _
                          & " AND cut_id = '" & strCutID & "'" _
                          & " AND size_desc = '" & strGpSize & "'" _
                          & " ORDER BY size_desc"


         With Rsd

              .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
              .LockType = ADODB.LockTypeEnum.adLockOptimistic
              .Open(strSqlSelc, Conn, , , )

              If .RecordCount <> 0 Then

                 cmbTMaterial.Text = .Fields("mate_type").Value.ToString.Trim
                 cboCutdetail.Text = .Fields("cut_detail").Value.ToString.Trim
                 txtPart.Text = .Fields("backgup").Value.ToString.Trim
                 txtCutID.Text = .Fields("cut_id").Value.ToString.Trim
                 txtSize.Text = .Fields("size_id").Value.ToString.Trim
                 txtSizeDesc.Text = .Fields("size_desc").Value.ToString.Trim
                 txtSetQty.Text = Format(.Fields("set_qty").Value, "#,##0.0")
                 txtSizeQty.Text = Format(.Fields("size_qty").Value, "#,##0.0")
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


             End If
            .ActiveConnection = Nothing    '����������������
            .Close()

    End With

End If

Conn.Close()
Conn = Nothing

End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
   ShowResvrd()
   ClearAllData()
   'CallEditData()    '�Ѻ�ٷչ�ʴ� Size ������䢢�����
   StateLockAEItem(True)
   gpbSeek.Text = "����������"
   txtCutID.Text = ""
   txtCutID.ReadOnly = False
   txtSize.ReadOnly = False
   txtSizeDesc.ReadOnly = False
   txtCutID.Focus()
  End Sub

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
  DeleteSubData()
End Sub

Private Sub DeleteSubData()
 Dim Conn As New ADODB.Connection
 Dim strCmd As String = ""

 Dim btyConsider As Byte
 Dim strSize As String = ""
 Dim strSizeAct As String = ""
 Dim strCutID As String = ""     '�����մ�Ѵ
 Dim strGpsize As String

   With dgvSize

        If .Rows.Count > 0 Then
             strSize = .Rows(.CurrentRow.Index).Cells(2).Value.ToString
             strSizeAct = .Rows(.CurrentRow.Index).Cells(3).Value.ToString
             strCutID = .Rows(.CurrentRow.Index).Cells(4).Value.ToString
             strGpsize = .Rows(.CurrentRow.Index).Cells(5).Value.ToString

              If strSizeAct <> "" Then

                    btyConsider = MsgBox("SIZE : " & strSize.ToString.Trim & vbNewLine _
                                                   & "�����մ : " & strCutID.ToString.Trim & vbNewLine _
                                                   & "����䫵� : " & strGpsize.ToString.Trim & vbNewLine _
                                                   & "�س��ͧ���ź���������!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                   + MsgBoxStyle.Exclamation, "Confirm Delete Data")

                     If btyConsider = 6 Then

                           If Conn.State Then Close()

                              Conn.ConnectionString = strConnAdodb
                              Conn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                              Conn.ConnectionTimeout = 90
                              Conn.Open()


                              strCmd = " DELETE FROM tmp_eqptrn" _
                                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                     & " AND size_id = '" & strSize.ToString.Trim & "'" _
                                                     & " AND cut_id = '" & strCutID.ToString.Trim & "'" _
                                                     & " AND size_desc = '" & strGpsize.ToString.Trim & "'"

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

Private Sub btnEditEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp1.Click

  Dim OpenFileDialog As New OpenFileDialog
  Dim strFileFullPath As String   '�纾������
  Dim strFileName As String       '�纪������
  Dim img As Image = Nothing

  Dim dateNow As Date = Now
  Dim typePic As String
  Dim strNamePic As String
  Dim lengPic, lengTypePic As Integer

      With OpenFileDialog
           .CheckFileExists = True  '��Ǩ�ͺ��������������к�
           .ShowReadOnly = False    '����ʴ��຺��ҹ���ҧ����
           .Filter = "All Files|*.*|����ٻ�Ҿ (*)|*.bmp;*.gif;*.jpg;*.png"
           .FilterIndex = 2

      Try

          If .ShowDialog = Windows.Forms.DialogResult.OK Then
              ' Load ������ picturebox
              strFileName = New System.IO.FileInfo(.FileName).Name               '�Ѻ���੾�Ъ������
              strFileFullPath = System.IO.Path.GetDirectoryName(.FileName)       '�Ѻ���੾�оҸ���

              img = ScaleImage(Image.FromFile(.FileName), picEqp1.Height, picEqp1.Width)      '��Ѻ��Ҵ�ٻ�Ҿ�����Ŵ�����ʹաѺ picbox
              picEqp1.Image = img                   '��ҹ����ٻ������ picBox

              '----------- �����ѹ������ҵ�ͪ������ ------------
              strFileName = Trim(strFileName)
              lengTypePic = strFileName.Length - 4
              typePic = Mid(strFileName, lengTypePic + 1, 4)                  ' �Ѵ��� .jpg .png .gif 
              lengPic = strFileName.Length - 4                                '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
              strNamePic = Mid(strFileName, 1, lengPic)                       '�Ѵ���੾�Ъ����ٻ
              strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ

              lblPicPath1.Text = strFileFullPath
              lblPicName1.Text = strFileName

          End If

      Catch ex As Exception
           ClearBlankPicture1()
      End Try

      End With
End Sub

Private Sub btnEditEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp2.Click

  Dim OpenFileDialog As New OpenFileDialog
  Dim strFileFullPath As String       '�纾������
  Dim strFileName As String           '�纪������
  Dim img As Image = Nothing

  Dim dateNow As Date = Now
  Dim typePic As String
  Dim strNamePic As String
  Dim lengPic, lengTypePic As Integer

      With OpenFileDialog

           .CheckFileExists = True        '��Ǩ�ͺ����������������к�
           .ShowReadOnly = False
           .Filter = "All Files|*.*|����ٻ�Ҿ (*)|*.bmp;*.gif;*.jpg;*.png"
           .FilterIndex = 2

      Try

         If .ShowDialog = Windows.Forms.DialogResult.OK Then   '����͵ͺ��ŧ

            '��Ŵ ������ picturebox
            strFileName = New System.IO.FileInfo(.FileName).Name    '�Ѻ���੾�Ъ������
            strFileFullPath = System.IO.Path.GetDirectoryName(.FileName)   '�Ѻ���੾�оҸ���


            img = ScaleImage(Image.FromFile(.FileName), picEqp2.Height, picEqp2.Width)   '��Ѻ��Ҵ�ٻ�Ҿ�����Ŵ�����ʹաѺ picbox
            picEqp2.Image = img               '��ҹ����ٻ������ picBox

                    '----------- �����ѹ������ҵ�ͪ������ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4) ' �Ѵ��� .jpg .png .gif 
                    lengPic = strFileName.Length - 4   '��Ҩӹǹ charactor ��������ź�͡����ʡ�� picture
                    strNamePic = Mid(strFileName, 1, lengPic)     '�Ѵ���੾�Ъ����ٻ
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           '�����ѹ����ͷ��ª�������ٻ


            lblPicPath2.Text = strFileFullPath
            lblPicName2.Text = strFileName

         End If

      Catch ex As Exception
            ClearBlankPicture2()
      End Try

      End With
End Sub

'---------------------------------------------- Clear PictureBox1 ------------------------------------------
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

Private Sub btnDelEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp1.Click
  ClearBlankPicture1()
End Sub

Private Sub btnDelEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp2.Click
  ClearBlankPicture2()
End Sub

Private Sub cboTypeCut_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCutdetail.GotFocus
  txtCutID.Text = txtCutID.Text.ToUpper.Trim
End Sub

Private Sub cboTypeCut_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCutdetail.LostFocus
 Dim strEqpid As String
     strEqpid = txtCutID.Text.Trim

         '--------------- ����ʵ�ԧ "-" ������������ ----------------------

         If InStr(1, strEqpid, "-") > 0 Then
             txtCutID.Text = strEqpid.ToUpper.Trim
         Else
             AddTxtCutID()  '����ʵ�ԧ��ͷ��� eqp_id
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
                        cboCutdetail.DroppedDown = True
                        cboCutdetail.Focus()
                     End If
                Case Is = 38 '�����١�â��     
                     txtCutID.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                       txtSize.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                           txtSize.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtSizeQty.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPart.KeyPress
   If e.KeyChar = Chr(13) Then
      txtSize.Focus()
   End If
End Sub

Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
    With txtRemark
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
                   Case Is = 40 '����ŧ
                   Case Is = 113 '���� F2
                             .SelectionStart = .Text.Trim.Length
            End Select
   End With
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