Imports ADODB
Imports System.IO
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmAeNotifyIssue
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

Private Sub frmAeNotifyIssue_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 ClearTmpTable(0, "")  'ź������ Table tmp_eqptrn where user_id..
 frmNotifyIssue.lblCmd.Text = "0"  '������ʶҹ�
 Me.Dispose()     '����¿���� �׹˹��¤�����

End Sub

Private Sub frmAeReqFxeqp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Dim dtComputer As Date = Now       '������纤���ѹ���Ѩ�غѹ
  Dim strCurrentDate As String       '�纤��ʵ�ԧ�ѹ���Ѩ�غѹ

  Me.WindowState = FormWindowState.Maximized  '��������������˹�Ҩ�
  StdDateTimeThai()        '�Ѻ�ٷչ��ŧ�ѹ�����Ǵ�.�� ����� Module
  strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

  ClearAlldata()
  PreDeptSeek()            '��Ŵ��ª��͡��Ἱ�
  PreGroupSeek()           '��Ŵ������ػ�ó�


     Select Case frmNotifyIssue.lblCmd.Text.ToString

            Case Is = "0" '����������

                With txtBegin
                     .Text = strCurrentDate
                     strDateDefault = strCurrentDate
                End With

                GenDocid()  '�ѹ�Ţ����͡���
                gpbReceive.Enabled = False
                txtDocid.ReadOnly = True


             Case Is = "1" '��䢢�����

               ' InputTmpData()  'Copy ������ŧ Tmpdata

                         If frmMainPro.lblLogin.Text.Trim = "SUTID" Then

                                      If chkFirstApprove() Then     '��Ǩ�ͺ��� ���.Ἱ��������͹��ѵ�����
                                         LockEditData()

                                         gpbNotify.Enabled = False   '�Դgroupbox ��ǹ�����
                                         gpbReceive.Enabled = True

                                      Else

                                         LockEditData()

                                         gpbNotify.Enabled = False   '�Դgroupbox ��ǹ�����
                                         gpbReceive.Enabled = False
                                         btnSaveData.Enabled = False

                                         MessageBox.Show("�͡����ѧ�����͹��ѵԨҡἹ�����駻ѭ�� �ô�Դ���Ἱ������!...", _
                                                                     "�������ö���Թ����� ", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                                      End If

                          Else
                                LockEditData()

                                gpbNotify.Enabled = True
                                gpbReceive.Enabled = False
                                txtDocid.ReadOnly = True

                          End If


             Case Is = "2"   '����ͧ������

                LockEditData()
                btnSaveData.Enabled = False
                txtDocid.ReadOnly = True

        End Select
End Sub

Function chkFirstApprove() As Boolean

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim strReqid As String
     strReqid = frmNotifyIssue.dgvIssue.Rows(frmNotifyIssue.dgvIssue.CurrentRow.Index).Cells(2).Value.ToString

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

        strSqlSelc = "SELECT person2 " _
                                 & " FROM notifyissue" _
                                 & " WHERE req_id = '" & strReqid & "'"

        Rsd = New ADODB.Recordset

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            If .Fields("person2").Value.ToString.Trim <> "" Then

                Return True

            Else
                Return False

            End If

            .ActiveConnection = Nothing   '������ Connection
            .Close()

        End With
        Rsd = Nothing   '������ RecordSet
  Conn.Close()    '�Դ�����������
  Conn = Nothing   '������ RecordSet

End Function

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

                           strSqlcmd = "DELETE tmp_notifyissue " _
                                 & "WHERE user_id = '" & frmMainPro.lblLogin.Text & "'"
                          .Execute(strSqlcmd)

                   Case Is = "1"

            End Select

     End With

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub LockEditData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim Rsdwc As New ADODB.Recordset

  Dim strCmd As String  ' ��ʵ�ԧ Command

  Dim strLoadFilePicture As String   '�纤��ʵ�ԧ��Ŵ Picture
  Dim strPathPicture As String = "H:\EquipPicture\"   '�� part

  Dim strSqlSelc As String = ""   '��ʵ�ԧ sql select
  Dim strPart As String = ""
  Dim strGpType As String = ""
  Dim strDocID As String = frmNotifyIssue.dgvIssue.Rows(frmNotifyIssue.dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim

      With Conn

           If .State Then Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()
      End With

        strSqlSelc = "SELECT * " _
                             & "FROM notifyissue (NOLOCK)" _
                             & " WHERE req_id = '" & strDocID & "'"

        Rsd = New ADODB.Recordset

        With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelc, Conn, , )

         If .RecordCount <> 0 Then

             Select Case .Fields("group").Value.ToString.Trim

                    Case Is = "A"
                         strGpType = "���촩մ EVA INJECTION"
                    Case Is = "B"
                         strGpType = "���촩մ PVC INJECTION"
                    Case Is = "C"
                         strGpType = "������ʹ PU"
                    Case Is = "D"
                         strGpType = "����ἧ�Ѵ���˹ѧ˹��,���"
                    Case Is = "E"
                         strGpType = "�մ�Ѵ"
                    Case Is = "F"
                         strGpType = "���͡ʡ�չ"
                    Case Is = "G"
                         strGpType = "���͡����"

             End Select

                txtDocid.Text = .Fields("req_id").Value.ToString.Trim
                cboDepto.Text = .Fields("to_dep").Value.ToString.Trim
                txtFrom.Text = .Fields("from_notify").Value.ToString.Trim
                cboDepfrom.Text = .Fields("dep_notify").Value.ToString.Trim
                cboGroup.Text = strGpType
                txtName.Text = .Fields("person1").Value.ToString.Trim

                txtOrder.Text = .Fields("order").Value.ToString.Trim
                txtShoe.Text = .Fields("shoe").Value.ToString.Trim
                txtSize.Text = .Fields("size").Value.ToString.Trim
                txtSizeQty.Text = .Fields("amount").Value.ToString.Trim
                txtEqpnm.Text = .Fields("eqpnm").Value.ToString.Trim
                txtIssue.Text = .Fields("issue").Value.ToString.Trim
                txtCause.Text = .Fields("cause").Value.ToString.Trim

                If Mid(.Fields("needdate").Value.ToString.Trim, 1, 10) = "" Then
                   txtNeedDate.Text = "__/__/____"

                Else
                   txtNeedDate.Text = Mid(.Fields("needdate").Value.ToString.Trim, 1, 10)

                End If

                txtNeedtime.Text = .Fields("needtime").Value.ToString.Trim
                txtRemark.Text = .Fields("remark").Value.ToString.Trim
                txtFxIssue.Text = .Fields("fxissue").Value.ToString.Trim

                If Mid(.Fields("wantdate").Value.ToString.Trim, 1, 10) = "" Then
                   txtWantDate.Text = "__/__/____"

                Else
                  txtWantDate.Text = Mid(.Fields("wantdate").Value.ToString.Trim, 1, 10)

                End If

                txtWantTime.Text = .Fields("wanttime").Value.ToString.Trim

                '------------------------------- Load �ٻ�Ҿ�ػ�ó� -----------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_issue").Value.ToString.Trim
                ' �紾�������������ԧ
                If File.Exists(strLoadFilePicture) Then

                   Dim img1 As Image      '��С�ȵ���� img1 �������Ҿ
                   img1 = Image.FromFile(strLoadFilePicture)  'img1 ��ҡѺpicture �����Ŵ�Ҩҡ db
                   Dim s1 As String = ImageToBase64(img1, System.Drawing.Imaging.ImageFormat.Jpeg)  '��С�ȵ���� s1 �纤��ʵ�ԧ����ŧ����
                   img1.Dispose()     '����µ���� img1
                   Piceqp.Image = Base64ToImage(s1)

                Else
                    Piceqp.Image = Nothing  '���������ٻ�Ҿ��� picEqp1 ��ҧ����

                End If


                                  strCmd = frmNotifyIssue.lblCmd.Text.ToString.Trim    '��� strCmd ��ҡѺ���� lblcmd 㹿���� frmEqpSheet

                                  Select Case strCmd

                                         Case Is = "1"   '�����ͤ�͹���

                                         Case Is = "2"   '�����ͤ�͹����ͧ
                                                btnSaveData.Enabled = False  '�Դ���� "�ѹ�֡������"

                                  End Select
   
        End If

            .ActiveConnection = Nothing   '��觻Դ�����������
            .Close()

        End With

   Rsd = Nothing   '�������� RecordSet
   Conn.Close()    '��觵Ѵ�����������
   Conn = Nothing  '������ Connection

End Sub

Private Sub ClearAlldata()

 txtDocid.Text = ""
 cboDepto.Text = ""
 txtFrom.Text = ""
 txtName.Text = ""

 cboDepfrom.Text = ""
 txtOrder.Text = ""
 txtShoe.Text = ""
 txtSize.Text = ""
 cboGroup.Text = ""

 txtEqpnm.Text = ""
 txtIssue.Text = ""
 txtCause.Text = ""

 txtNeedDate.Text = "__/__/____"
 txtNeedtime.Text = ""
 txtRemark.Text = ""

 txtFxIssue.Text = ""
 txtWantDate.Text = "__/__/____"
 txtWanttime.Text = ""

End Sub

Private Sub PreDeptSeek()

 Dim strDept(5) As String
 Dim strDeptTo(0) As String
 Dim i As Integer

    strDept(0) = "121000 Ἱ���Ե��"
    strDept(1) = "122000 Ἱ��Ѵ�����ǹ"
    strDept(2) = "123000 Ἱ����"
    strDept(3) = "124000 Ἱ��մ PVC"
    strDept(4) = "125000 Ἱ��մ EVA INJECTION"
    strDept(5) = "126000 Ἱ��մ PU"


    strDeptTo(0) = "Ἱ���鹵͹����ػ�ó��ü�Ե"

    With cboDepfrom

         For i = 0 To 5
               .Items.Add(strDept(i))

         Next

    End With

       With cboDepto

           .Items.Add(strDeptTo(0))

       End With

End Sub

Private Sub PreGroupSeek()

 Dim strGroup(6) As String
 Dim i As Integer

    strGroup(0) = "���촩մ EVA INJECTION"
    strGroup(1) = "���촩մ PVC"
    strGroup(2) = "������ʹ PU"
    strGroup(3) = "����ἧ�Ѵ���˹ѧ˹��,���"
    strGroup(4) = "�մ�Ѵ"
    strGroup(5) = "���͡ʡ�չ"
    strGroup(6) = "���͡����"

    With cboGroup

         For i = 0 To 6
               .Items.Add(strGroup(i))

         Next

    End With

End Sub

Private Sub GenDocid()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim LastNumber As Integer
 Dim LastYear As Integer
 Dim DateCom As Date = Now
 Dim strCurrentDate As String
 Dim Thayear As String

     strCurrentDate = DateCom.Date.ToString("yyyy-MM-dd")
     Thayear = Mid(SaveChangeThaYear(strCurrentDate), 3, 2) '�鴻��� 5X

      With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

      End With

      strSqlSelc = "SELECT * FROM notifyissue (NOLOCK) "

      Rsd = New ADODB.Recordset

      With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

             .MoveLast()              '����͹��ѧ Record �ش����

             LastNumber = CInt(Mid(.Fields("req_id").Value.ToString.Trim, 5))  '�Ѵʵ�ԧ ��� 4 ��Ƿ���  000x
             LastYear = Mid(.Fields("req_id").Value.ToString.Trim, 3, 2)  '�Ѵ��һ�  5x ੾�� 2����á

               If String.Compare(LastYear, Thayear) = 0 Then       '���º��º ʵ�ԧ�� 5x
                  LastYear = LastYear
                  LastNumber += 1


               Else
                  LastYear += 1  ' ������� LestRec �ա 1.
                  LastNumber = 1

               End If

          Else
               LastYear = Thayear
               LastNumber = 1

          End If

          txtDocid.Text = "DN" & LastYear & LastNumber.ToString("0000")

      .ActiveConnection = Nothing
      .Close()
      End With

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
 CheckDataBfSave()

End Sub

Private Sub CheckDataBfSave()
   MsgBox("�� checkDatabfsave")
 Dim strProd As String = ""
 Dim strProdnm As String = ""

     If cboDepto.Text.ToString.Trim <> "" Then

         If txtFrom.Text.ToString.Trim <> "" Then

               If txtName.Text.ToString.Trim <> "" Then

                    If cboDepfrom.Text.ToString.Trim <> "" Then

                          If txtShoe.Text.ToString.Trim <> "" Then

                                  If cboGroup.Text <> "" Then

                                           If txtSize.Text.ToString.Trim <> "" Then

                                                     If txtIssue.Text.ToString.Trim <> "" Then

                                                         Select Case frmNotifyIssue.lblCmd.Text.Trim

                                                                Case Is = "0"      '�ó�����������

                                                                     If CheckCodeDuplicate() Then
                                                                        MsgBox("�� SaveNewdata")
                                                                        SaveNewData()

                                                                     Else
                                                                         MessageBox.Show("��س��͡�ҡ����� ���Ǵ��Թ�������!....", "***�����͡��ë��***" _
                                                                            , MessageBoxButtons.OK, MessageBoxIcon.Error)

                                                                         ClearAlldata()    '��ҧ˹�Ҩ�                                                            
                                                                         cboDepto.Focus()
                                                                      End If


                                                                 Case Is = "1"         '�ó���䢢�����
                                                                       SaveEditData()


                                                            End Select

                                                         Else

                                                             MsgBox("�ô�кػѭ�ҷ�辺  " _
                                                                   & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                                             txtIssue.Focus()
                                                         End If

                                                Else
                                                    MsgBox("�ô�к� Size �ػ�ó�  " _
                                                      & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                                    txtSize.Focus()

                                                 End If

                                        Else
                                            MsgBox("�ô���͡������ػ�ó�  " _
                                                 & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                            cboGroup.DroppedDown = True
                                            cboGroup.Focus()

                                         End If

                                  Else
                                       MsgBox("�ô�к�����ػ�ó�  " _
                                           & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                       txtShoe.Focus()

                                  End If


                            Else

                                 MsgBox("�ô���͡Ἱ� / ���¼���駻ѭ��  " _
                                            & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                 cboDepfrom.DroppedDown = True
                                 cboDepfrom.Focus()
                           End If

                  Else
                       MsgBox("�ô�кت��ͼ���駻ѭ��  " _
                                            & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                       txtName.Focus()

                  End If

            Else

                 MsgBox("�ô�к���ǹ�ҹ  " _
                                    & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                 txtFrom.Focus()

            End If

    Else

        MsgBox("�ô�к�Ἱ��Ѻ����ͧ��駻ѭ��  " _
                               & " ��͹�ѹ�֡������", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
        cboDepto.DroppedDown = True
        cboDepto.Focus()

    End If

End Sub

Private Sub SaveEditData()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

  Dim dateSave As Date = Now()
  Dim strDate As String = ""

  Dim strWantDate As String
  Dim strNeedDate As String
  Dim strDocdate As String           '��ʵ�ԧ�ѹ����͡���
  Dim strGpType As String = ""       '�纻������ػ�ó�
  Dim strPartType As String = ""     '�纪����ǹ����Ե
  Dim strDocid As String
  Dim blnReturnCopyPic As Boolean

      With Conn
            If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()

     End With

               strDocid = frmNotifyIssue.dgvIssue.Rows(frmNotifyIssue.dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim()

               'Conn.BeginTrans()      '�ش������� Transection

               strDate = dateSave.Date.ToString("yyyy-MM-dd")
               strDate = "'" & SaveChangeEngYear(strDate) & "'"

              '------------------------- �ѹ����͡��� ----------------------------------------------------

               strDocdate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocdate = "'" & SaveChangeEngYear(strDocdate) & "'"


              '---------------------------------------- Ǵ�.����Դ --------------------------------------------

               If txtNeedDate.Text <> "__/__/____" Then

                  strNeedDate = Mid(txtNeedDate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtNeedDate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtNeedDate.Text.ToString, 1, 2)
                  strNeedDate = "'" & SaveChangeEngYear(strNeedDate) & "'"

               Else
                  strNeedDate = "NULL"

               End If


               '---------------------------------------- Ǵ�.����Դ --------------------------------------------

               If txtWantDate.Text <> "__/__/____" Then

                  strWantDate = Mid(txtWantDate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtWantDate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtWantDate.Text.ToString, 1, 2)
                  strWantDate = "'" & SaveChangeEngYear(strWantDate) & "'"

               Else
                  strWantDate = "NULL"

               End If

                      '------------------------------------ �ѹ�֡�ٻ�ػ�ó�  ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath.Text.ToString.Trim, lblPicName.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath.Text = PthName

                       Else
                          lblPicName.Text = ""
                          lblPicPath.Text = ""
                          Piceqp.Image = Nothing

                       End If


                           Select Case cboGroup.Text.ToString.Trim

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

                            If frmMainPro.lblLogin.Text = "MALIWAN" Or frmMainPro.lblLogin.Text = "SUTID" Then       '��Ǩ�ͺ user login

                                 strSqlCmd = "UPDATE notifyissue SET fxissue = '" & ReplaceQuote(txtFxIssue.Text.ToString.Trim) & "'" _
                                              & "," & "wantdate  = " & strWantDate _
                                              & "," & "wanttime  = '" & ReplaceQuote(txtWanttime.Text.ToString.Trim) & "'" _
                                              & "," & "lastby  = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                              & "," & "last_Date  = " & strDate _
                                              & " WHERE req_id = '" & strDocid & "'"

                                Conn.Execute(strSqlCmd)

                            Else

                               strSqlCmd = "UPDATE notifyissue SET [group] = '" & strGpType & "'" _
                                              & "," & "to_dep = '" & ReplaceQuote(cboDepto.Text.ToString.Trim) & "'" _
                                              & "," & "from_notify = '" & ReplaceQuote(txtFrom.Text.ToString.Trim) & "'" _
                                              & "," & "dep_notify = '" & ReplaceQuote(cboDepfrom.Text.Trim) & "'" _
                                              & "," & "[order]  = '" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                                              & "," & "eqpnm  = '" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                              & "," & "shoe  = '" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                                              & "," & "size  = '" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                              & "," & "amount  = " & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                              & "," & "issue  = '" & ReplaceQuote(txtIssue.Text.ToString.Trim) & "'" _
                                              & "," & "cause  = '" & ReplaceQuote(txtCause.Text.ToString.Trim) & "'" _
                                              & "," & "needdate  = " & strNeedDate _
                                              & "," & "needtime  = '" & ReplaceQuote(txtNeedtime.Text.ToString.Trim) & "'" _
                                              & "," & "pic_Issue  = '" & ReplaceQuote(lblPicName.Text.ToString.Trim) & "'" _
                                              & "," & "person1_sta  = '" & True & "'" _
                                              & "," & "person1  = '" & ReplaceQuote(txtName.Text.ToString.Trim) & "'" _
                                              & "," & "person1_date  = " & strDate _
                                              & "," & "lastby  = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                              & "," & "last_date  = " & strDate _
                                              & "," & "remark  = '" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                              & " WHERE req_id = '" & strDocid & "'"

                              Conn.Execute(strSqlCmd)
                              'Conn.CommitTrans()  '��� Commit transection

                            End If

        lblComplete.Text = txtDocid.Text.ToString.Trim  '�觺͡��Һѹ�֡�����������

        Me.Hide()
        frmMainPro.Show()
        frmNotifyIssue.Show()

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub SaveNewData()
 MsgBox("�� Sub SaveNewData")
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim dateSave As Date = Now()    '�纤���ѹ���Ѩ�غѹ
 Dim strDate As String

 Dim blnRetuneCopyPic As Boolean
 Dim strNeedDate As String
 Dim strDateNull As String = "Null"
 Dim strWantDate As String
 Dim strDateDoc As String
 Dim strGpType As String = ""
 Dim strType As String = ""


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
                strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"

                '---------------------------------------- �ѹ�֡�ٻ�ػ�ó� ----------------------------------------------

                blnRetuneCopyPic = CallCopyPicture(lblPicPath.Text.Trim, lblPicName.Text.Trim)

                    If blnRetuneCopyPic Then          '��� CallCopyPicture = true
                       lblPicPath.Text = PthName
                    Else
                       lblPicPath.Text = ""
                       lblPicName.Text = ""
                       picEqp = Nothing

                    End If


                    '---------------------------------------- Ǵ�.����ͧ��� ----------------------------------------------------
                    If txtNeedDate.Text <> "__/__/____" Then

                       strNeedDate = Mid(txtNeedDate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtNeedDate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtNeedDate.Text.ToString, 1, 2)
                       strNeedDate = "'" & SaveChangeEngYear(strNeedDate) & "'"

                   Else
                       strNeedDate = "NULL"
                   End If


                    '---------------------------------------- Ǵ�.��˹����� -------------------------------------------------
                   If txtWantDate.Text <> "__/__/____" Then

                      strWantDate = Mid(txtWantDate.Text.ToString, 7, 4) & "-" _
                                    & Mid(txtWantDate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtWantDate.Text.ToString, 1, 2)
                      strWantDate = "'" & SaveChangeEngYear(strWantDate) & "'"

                   Else
                      strWantDate = "NULL"
                   End If

                   '------------------------------------��˹�������ͧ��鹧ҹ--------------------------------------------------

                   Select Case cboGroup.Text.ToString.Trim

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

                 strSqlCmd = "INSERT INTO notifyissue" _
                       & "(req_id,req_sta,[group],to_dep,from_notify,dep_notify" _
                       & ",[order],eqpnm,shoe,size,amount" _
                       & ",issue,cause,needdate,needtime,fxissue,wantdate" _
                       & ",wanttime,pic_issue,person1_sta,person1,person1_date,person2_sta" _
                       & ",person2,person2_date,person3_sta,person3,person3_date,person4_sta,person4,person4_date" _
                       & ",recordby,record_date,lastby,last_date,remark" _
                       & ")" _
                       & " VALUES (" _
                       & "'" & ReplaceQuote(txtDocid.Text.Trim) & "'" _
                       & ",'" & "0" & "'" _
                       & ",'" & strGpType & "'" _
                       & ",'" & ReplaceQuote(cboDepto.Text.Trim) & "'" _
                       & ",'" & ReplaceQuote(txtFrom.Text.Trim) & "'" _
                       & ",'" & ReplaceQuote(cboDepfrom.Text.Trim) & "'" _
                       & ",'" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                       & ",'" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                       & ",'" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                       & ",'" & ReplaceQuote(txtSize.Text.Trim) & "'" _
                       & ",'" & ChangFormat(txtSizeQty.Text.ToString.Trim) & "'" _
                       & ",'" & ReplaceQuote(txtIssue.Text.ToString.Trim) & "'" _
                       & ",'" & ReplaceQuote(txtCause.Text.ToString.Trim) & "'" _
                       & "," & strNeedDate _
                       & ",'" & ReplaceQuote(txtNeedtime.Text.ToString.Trim) & "'" _
                       & ",'" & "" & "'" _
                       & "," & strDateNull _
                       & ",'" & "" & "'" _
                       & ",'" & ReplaceQuote(lblPicName.Text.ToString.Trim) & "'" _
                       & ",'" & True & "'" _
                       & ",'" & ReplaceQuote(txtName.Text.ToString.Trim) & "'" _
                       & "," & strDateDoc _
                       & ",'" & False & "'" _
                       & ",'" & "" & "'" _
                       & "," & strDateNull _
                       & ",'" & False & "'" _
                       & ",'" & "" & "'" _
                       & "," & strDateNull _
                       & ",'" & False & "'" _
                       & ",'" & "" & "'" _
                       & "," & strDateNull _
                       & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                       & "," & strDateDoc _
                       & ",'" & "" & "'" _
                       & "," & strDateNull _
                       & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                       & ")"

                Conn.Execute(strSqlCmd)
                Conn.CommitTrans()


           lblComplete.Text = txtDocid.Text.ToString.Trim     '�觺͡��Һѹ�֡�����������
           Me.Hide()    '����������Ѩ�غѹ

           frmMainPro.Show()
           frmNotifyIssue.Show()

   Conn.Close()
   Conn = Nothing

End Sub

Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String) As Boolean

  Dim fname As String = String.Empty  '��ҡѺ ""
  Dim dFile As String = String.Empty
  Dim dFilePath As String = String.Empty

  Dim fServer As String = String.Empty
  Dim intResult As Integer   '�׹����繨ӹǹ���

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

                '------------------------------------��Ҥ���� 0 �ʴ������Ŵ��������� �������ö Copy �����------------------------------

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

        strSqlSelc = "SELECT req_id " _
                    & " FROM notifyissue" _
                    & " WHERE req_id = '" & txtDocid.Text.Trim & "'"

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

Private Sub btnEditEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp1.Click

 Dim OpenFileDialog1 As New OpenFileDialog
 Dim strFileFullPath As String   '�纾������
 Dim strFileName As String       '�� filenaem

     With OpenFileDialog1

            .CheckFileExists = True
            .ShowReadOnly = False
            .Filter = "All Files|*.*|����ٻ�Ҿ (*)|*.bmp;*.gif;*.jpg;*.png"
            .FilterIndex = 2

            Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then

                    ' Load ������ picturebox
                    Piceqp.Image = Image.FromFile(.FileName)

                    strFileName = New System.IO.FileInfo(.FileName).Name '�Ѻ���੾�Ъ������
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) '�Ѻ���੾�оҸ���

                    lblPicPath.Text = strFileFullPath
                    lblPicName.Text = strFileName             '�����������С�ȵ�������С���
                    'lblPicName.Text = .FileName.Substring(.FileName.LastIndexOf("\") + 1)

                End If

            Catch ex As Exception
                  ClearBlankPicture()
            End Try

     End With

End Sub

Private Sub ClearBlankPicture()
  Piceqp.Image = Nothing
  lblPicPath.Text = ""
  lblPicName.Text = ""
End Sub

Private Sub btnDelEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp1.Click
 ClearBlankPicture()
End Sub

Private Sub cboDepto_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

  Dim intChkPoint As Integer

      With cboDepto

         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                Case Is = 38 '�����١�â��    
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtFrom.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtFrom.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                      txtFrom.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSendTo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

  Select Case e.KeyChar
         Case Is = Chr(13)  '��� Enter
               txtFrom.Focus()
  End Select

End Sub

Private Sub txtFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFrom.KeyDown

  Dim intChkPoint As Integer

      With txtFrom

         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                      If .SelectionLength = .Text.Trim.Length Then
                           cboDepto.Focus()

                      End If

                Case Is = 38 '�����١�â��    
                         cboDepto.Focus()
                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        txtName.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtName.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     cboDepfrom.DroppedDown = True
                     cboDepfrom.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtFrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFrom.KeyPress

 Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
              txtName.Focus()
 End Select

End Sub

Private Sub cboDepfrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboDepfrom.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
               txtOrder.Focus()
  End Select

End Sub

Private Sub txtOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrder.KeyDown
 Dim intChkPoint As Integer

      With txtOrder

         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                        If .SelectionLength = .Text.Trim.Length Then
                          cboDepfrom.DroppedDown = True

                        End If
                Case Is = 38 '�����١�â��    
                        cboDepfrom.DroppedDown = True
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                       txtShoe.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                           txtShoe.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtShoe.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrder.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
             txtShoe.Focus()
  End Select

End Sub

Private Sub cboGroup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboGroup.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
              txtSize.Focus()
  End Select

End Sub

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown

  Dim intChkPoint As Integer

      With txtSize

         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                        If .SelectionLength = .Text.Trim.Length Then
                          cboGroup.DroppedDown = True

                        End If
                Case Is = 38 '�����١�â��    
                       txtShoe.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                       txtEqpnm.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtEqpnm.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtEqpnm.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
             txtEqpnm.Focus()
  End Select

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
                       If .SelectionLength = .Text.Trim.Length Then
                          txtSize.Focus()

                       End If
                Case Is = 38 '�����١�â��    
                      cboGroup.DroppedDown = True

                Case Is = 39 '�����١�â��
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

Private Sub txtSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSizeQty.KeyPress

     Select Case Asc(e.KeyChar)

            Case 48 To 57 ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False
            Case 13
                e.Handled = False
                txtEqpnm.Focus()

            Case 8                  '���� Backspace
                e.Handled = False
            Case 32                   '��� spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Select


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

                    Case Is = -2
                        .SelectionStart = 0
                    Case Is = -1
                        .SelectionStart = 1
                    Case Is = 1
                        .SelectionStart = 2
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
                        txtSize.Focus()
                    End If
                Case Is = 38 '�����١�â��  
                     cboGroup.DroppedDown = True
                     cboGroup.Focus()

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtEqpnm.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkpoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtEqpnm.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtEqpnm.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
             txtEqpnm.Focus()
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
                    txtSizeQty.Text = z.ToString("0")
                End If
            Catch ex As Exception
                txtSizeQty.Text = "0"
                mskSizeQty.Text = ""
            End Try

            mskSizeQty.SendToBack()
            txtSizeQty.BringToFront()

        End With
End Sub

Private Sub txtEqpnm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqpnm.KeyDown

 Dim intChkPoint As Integer

      With txtEqpnm

         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtSizeQty.Focus()

                       End If
                Case Is = 38 '�����١�â��    
                      txtSize.Focus()

                Case Is = 39 '�����١�â��
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

Private Sub txtEqpnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqpnm.KeyPress

  Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
             txtIssue.Focus()

   End Select

End Sub

Private Sub txtIssue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtIssue.KeyDown

  Dim intChkPoint As Integer

      With txtIssue

         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtEqpnm.Focus()

                       End If
                Case Is = 38 '�����١�â��    
                      txtEqpnm.Focus()

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                       txtCause.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtCause.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtCause.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtIssue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIssue.KeyPress

  Select Case e.KeyChar
        Case Is = Chr(13)  '��� Enter
             txtCause.Focus()

   End Select
End Sub

Private Sub txtCause_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCause.KeyDown

  Dim intChkPoint As Integer

      With txtCause

         Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtIssue.Focus()

                       End If
                Case Is = 38 '�����١�â��    
                      txtIssue.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                       txtNeedDate.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtNeedDate.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtNeedDate.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

        End Select

     End With

End Sub

Private Sub txtCause_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCause.KeyPress
   Select Case e.KeyChar
          Case Is = Chr(13)  '��� Enter
             txtNeedDate.Focus()

   End Select
End Sub

Private Sub txtNeedDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNeedDate.GotFocus
  With mskNeedDate
       txtNeedDate.SendToBack()
       .BringToFront()
       .Focus()
  End With

End Sub

Private Sub mskNeedDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskNeedDate.GotFocus

 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerg As String = ""

     With mskNeedDate

            If txtNeedDate.Text.Trim <> "__/__/____" Then
                x = Len(txtNeedDate.Text.Trim)

                For i = 1 To x
                    strTmp = Mid(txtNeedDate.Text.Trim, i, 1)
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

Private Sub mskNeedDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskNeedDate.KeyDown

  Dim intChkPoint As Integer

        With mskNeedDate

         Select Case e.KeyCode

                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtCause.Focus()
                    End If

                Case Is = 38 '�����١�â��
                    txtCause.Focus()
                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then
                        txtNeedtime.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtNeedtime.Focus()
                        End If
                    End If
                Case Is = 40 '����ŧ           
                    txtNeedtime.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

          End Select

        End With

End Sub

Private Sub mskNeedDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskNeedDate.KeyPress
   If e.KeyChar = Chr(13) Then
      txtNeedtime.Focus()
   End If

End Sub

Private Sub mskNeedDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskNeedDate.LostFocus
  Dim i, x As Integer
  Dim z As Date

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskNeedDate

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

                mskNeedDate.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)

                If Year(z) < 2500 Then '�դ��ʵ� < 2100                        
                    txtNeedDate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))
                Else
                    txtNeedDate.Text = z.ToString("dd/MM/yyyy")
                End If

            Catch ex As Exception
                txtNeedDate.Text = "__/__/____"
                mskNeedDate.Text = ""

            End Try

            mskNeedDate.SendToBack()
            txtNeedDate.BringToFront()

        End With

End Sub

Private Sub txtNeedtime_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNeedtime.KeyDown

  Dim intChkPoint As Integer

      With txtNeedtime

         Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtNeedDate.Focus()

                       End If
                Case Is = 38 '�����١�â��    
                      txtNeedDate.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                       txtRemark.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
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

Private Sub txtNeedtime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNeedtime.KeyPress

   Select Case Asc(e.KeyChar)

            Case 48 To 57 ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False

            Case 13
                e.Handled = False
                 txtRemark.Focus()

            Case 8                  '���� Backspace
                e.Handled = False

            Case 32                 '��� spacebar
                e.Handled = False

            Case 58
                e.Handled = False   '��� :

            Case 46
                e.Handled = False   ' ��� .

            Case 44
                e.Handled = False   ' ��� ,

            Case Else
                e.Handled = True

                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)

     End Select

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  On Error Resume Next    '�ҡ�Դ error ������ѧ�зӧҹ���������ʹ� error ����Դ���
  Dim strCode As String

     If MessageBox.Show("��ͧ����͡�ҡ����� �������", "��س��׹�ѹ�͡�ҡ�����", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

           With frmNotifyIssue.dgvIssue
                If .Rows.Count > 0 Then   '����բ������ Grid
                    strCode = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim          '���strCode = ��������ǻѨ�غѹ Cell �á
                    lblComplete.Text = strCode  '��� label �ʴ�������� Cell �Ѩ�غѹ   
                End If

           End With
           Me.Close()

       frmMainPro.Show()
       frmNotifyIssue.Show()

     End If

End Sub

Private Sub txtShoe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShoe.KeyDown

 Dim intChkPoint As Integer

      With txtShoe

         Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtOrder.Focus()

                       End If
                Case Is = 38 '�����١�â��    
                      txtOrder.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                       txtSizeQty.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtSizeQty.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     txtSizeQty.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

        End Select

     End With

End Sub

Private Sub txtSeries_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShoe.KeyPress
   If e.KeyChar = Chr(13) Then
      cboGroup.DroppedDown = True
      cboGroup.Focus()

   End If

End Sub

Private Sub txtShoe_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtShoe.TextChanged
  txtShoe.Text = txtShoe.Text.ToUpper

End Sub

Private Sub txtName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyDown

 Dim intChkPoint As Integer

      With txtName

         Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë��� 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtFrom.Focus()

                       End If
                Case Is = 38 '�����١�â��    
                      txtFrom.Focus()

                Case Is = 39 '�����١�â��

                    If .SelectionLength = .Text.Trim.Length Then
                        cboDepfrom.DroppedDown = True
                        cboDepfrom.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            cboDepfrom.DroppedDown = True
                            cboDepfrom.Focus()
                        End If

                    End If

                Case Is = 40 '����ŧ    
                     cboDepfrom.DroppedDown = True
                     cboDepfrom.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

        End Select

     End With
End Sub

Private Sub txtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress
  If e.KeyChar = Chr(13) Then
     cboDepfrom.DroppedDown = True
     cboDepfrom.Focus()

   End If
End Sub

Private Sub txtWantDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWantDate.GotFocus

 With mskWantDate
      .BringToFront()
      txtWantDate.SendToBack()
      .Focus()
 End With

End Sub

Private Sub mskWantDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskWantDate.GotFocus

 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerg As String = ""

     With mskWantDate

            If txtWantDate.Text.Trim <> "__/__/____" Then
                x = Len(txtWantDate.Text.Trim)

                For i = 1 To x
                    strTmp = Mid(txtWantDate.Text.Trim, i, 1)

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

Private Sub mskWantDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskWantDate.KeyDown

  Dim intChkpoint As Integer

        With mskWantDate

            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtFxIssue.Focus()
                    End If

                Case Is = 38 '�����١�â��  
                     txtFxIssue.Focus()

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        txtWanttime.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkpoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            txtWanttime.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    txtWanttime.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub mskWantDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskWantDate.KeyPress
   If e.KeyChar = Chr(13) Then
      txtWanttime.Focus()

   End If

End Sub

Private Sub mskWantDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskWantDate.LostFocus

  Dim i, x As Integer
  Dim z As Date

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskWantDate

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

                mskWantDate.Text = ""
                strMerge = "#" & strMerge & "#"
                z = CDate(strMerge)

                If Year(z) < 2500 Then '�դ��ʵ� < 2100                        
                    txtWantDate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))
                Else
                    txtWantDate.Text = z.ToString("dd/MM/yyyy")
                End If

            Catch ex As Exception
                txtWantDate.Text = "__/__/____"
                mskWantDate.Text = ""

            End Try

            mskWantDate.SendToBack()
            txtWantDate.BringToFront()

        End With

End Sub

Private Sub txtWanttime_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtWanttime.KeyDown
 Dim intChkpoint As Integer

        With txtWanttime

            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                    If .SelectionStart = 0 Then
                        txtWantDate.Focus()
                    End If

                Case Is = 38 '�����١�â��  
                     txtFxIssue.Focus()

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        btnSaveData.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkpoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            btnSaveData.Focus()
                        End If
                    End If

                Case Is = 40 '����ŧ
                    btnSaveData.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub txtWanttime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWanttime.KeyPress
    Select Case Asc(e.KeyChar)

            Case 48 To 57 ' key �� �ͧ����Ţ�����������ҧ48-57��Ѻ 48����Ţ0 57����Ţ9����ӴѺ
                e.Handled = False

            Case 13
                e.Handled = False
                 btnSaveData.Focus()

            Case 8                  '���� Backspace
                e.Handled = False

            Case 32                 '��� spacebar
                e.Handled = False

            Case 58
                e.Handled = False   '��� :

            Case 46
                e.Handled = False   ' ��� .

            Case 44
                e.Handled = False   ' ��� ,

            Case Else
                e.Handled = True

                MessageBox.Show("��س��кآ������繵���Ţ", "����͹", MessageBoxButtons.OK, MessageBoxIcon.Warning)

     End Select
End Sub

End Class