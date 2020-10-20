Imports System.Data
Imports System.IO
Imports System.Drawing

Public Class frmSendPdf

 Dim strDateDefault As String     '���������Ѻ�ѹ�������
 Dim fileBinary() As Byte
 Dim strStatus As String

 Public Const DrvName As String = "\\10.32.0.15\data1\Eqpdocument\"
 Public Const PthName As String = "\\10.32.0.15\data1\Eqpdocument"

Private Sub frmSendPdf_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
  LoadData()
End Sub

Private Sub frmSendPdf_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 strStatus = ""
 ClearTmpTable(0, "")
 Me.Dispose()
End Sub

Private Sub frmSendPdf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
 Dim dtComputer As Date = Now       '������纤���ѹ���Ѩ�غѹ
 Dim strCurrentDate As String       '�纤��ʵ�ԧ�ѹ���Ѩ�غѹ

     Me.WindowState = FormWindowState.Maximized  '��������������˹�Ҩ�
     StdDateTimeThai()        '�Ѻ�ٷչ��ŧ�ѹ�����Ǵ�.�� ����� Module
     strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

          With txtBegin
               .Text = strCurrentDate
               strDateDefault = strCurrentDate

          End With

     PreDeptSend()  '��ŴἹ��Ѻ�͡���
     LoadData()
     Importdata()
     ClearAlldata()  '�����������
     Timer1.Enabled = True

End Sub

Private Sub PreDeptSend()
 Dim strDept(5) As String
 Dim i As Integer
     strDept(0) = "- Ἱ��Ѵ�����ǹ (�س��д�ɰ� �ѧ��ͧ)"
     strDept(1) = "- Ἱ��մ EVA(�س�Է���ѡ��� �ҹ�ش��Ե���)"
     strDept(2) = "- Ἱ��մ PU(�س�Ⱦ� �����ع�ø���)"
     strDept(3) = "- Ἱ����(�سʶԵ�� �ʹ�ѡ)"
     strDept(4) = "- Ἱ���Ե��(�س൪Թ�� ����ŧ)"
     strDept(5) = "- Ἱ��մ PVC(�س���� �ʧ��س����ط���)"

     With cboDepRecv

          For i = 0 To 5
              .Items.Add(strDept(i))
          Next i

     End With

End Sub

Private Sub ClearAlldata()
 lblDocid.Text = ""
 lblFilename.Text = ""
 cboDepRecv.Text = ""

 txtName.Text = ""
 txtRemark.Text = ""
End Sub

Private Sub LoadData()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim strSta As String = ""
 Dim strDepnm As String = ""

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

     strSqlSelc = "SELECT * " _
                          & "FROM docsend (NOLOCK)" _
                          & "ORDER BY doc_id"

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , )

          If .RecordCount <> 0 Then

             dgvDoc.Rows.Clear()
             dgvDoc.ScrollBars = ScrollBars.None      '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ

             dgvDoc.Rows.Clear() '������� Data grid

             Do While Not .EOF()

                  Select Case .Fields("doc_sta").Value.ToString.Trim

                         Case Is = "0"
                                strSta = "�ʹ��Թ���"

                         Case Is = "1"
                                strSta = "�Ѻ�͡�������"

                         Case Else

                  End Select

                              Select Case .Fields("recv_dept").Value.ToString.Trim

                                     Case Is = "D1"
                                        strDepnm = "- Ἱ��Ѵ�����ǹ (�س��д�ɰ� �ѧ��ͧ)"

                                     Case Is = "D2"
                                        strDepnm = "- Ἱ��մ EVA(�س�Է���ѡ��� �ҹ�ش��Ե���)"

                                     Case Is = "D3"
                                        strDepnm = "- Ἱ��մ PU(�س�Ⱦ� �����ع�ø���)"

                                     Case Is = "D4"
                                        strDepnm = "- Ἱ����(�سʶԵ�� �ʹ�ѡ)"

                                     Case Is = "D5"
                                        strDepnm = "- Ἱ���Ե��(�س൪Թ�� ����ŧ)"

                                     Case Is = "D6"
                                        strDepnm = "- Ἱ��մ PVC(�س���� �ʧ��س����ط���)"

                              End Select



                  dgvDoc.Rows.Add( _
                                      IIf(.Fields("doc_sta").Value.ToString.Trim = "0", My.Resources._16x16_ledred, My.Resources._16x16_ledgreen), _
                                      strSta, _
                                      "DC" & .Fields("doc_id").Value.ToString.Trim, _
                                      .Fields("content_type").Value.ToString.Trim, _
                                      strDepnm, _
                                      Mid(.Fields("send_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("send_by").Value.ToString.Trim, _
                                      Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                      .Fields("recv_by").Value.ToString.Trim, _
                                      .Fields("remark").Value.ToString.Trim _
                                 )
                  .MoveNext()

             Loop
              StateLockEdit(False)

          Else

              dgvDoc.Rows.Clear() '������� Data grid

              StateLockbtn(False)
              StateLockEdit(False)

              btnAdd.Enabled = True

          End If

            dgvDoc.ScrollBars = ScrollBars.Both '�ѹ ScrollBars �ͧ DataGrid Refresh ���ѹ

     .ActiveConnection = Nothing
     .Close()
     End With

 Conn.Close()
 Conn = Nothing

End Sub

Private Sub StateLockbtn(ByVal sta As Boolean)
  btnAdd.Enabled = sta
  btnEdit.Enabled = sta
  btnDelete.Enabled = sta

  dgvDoc.Enabled = sta

End Sub

Private Sub stateLockAdd(ByVal sta As String)
 btnUpload.Enabled = sta
 txtName.Enabled = sta
 txtRemark.Enabled = sta
 cboDepRecv.Enabled = sta

 btnSave.Enabled = False
 btnCancle.Enabled = True

End Sub

Private Sub StateLockEdit(ByVal sta As Boolean)
 btnSave.Enabled = sta
 btnCancle.Enabled = sta
 btnUpload.Enabled = sta
 cboDepRecv.Enabled = sta

 txtName.Enabled = sta
 txtRemark.Enabled = sta

End Sub

Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
 If txtName.Text <> "" Then

    If cboDepRecv.Text <> "" Then
       Loadfiles() 'Upload ����͡���
       StatelockUpload(True)

    Else
       MessageBox.Show("��س����͡Ἱ�����Ѻ�͡��� ��͹Ṻ��������", "�Դ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Error)
       cboDepRecv.DroppedDown = True

    End If

 Else
       MessageBox.Show("��س��к���Ǣ���͡��� ��͹Ṻ�������", "�Դ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Error)
       txtName.Focus()

 End If

End Sub

Private Sub StatelockUpload(ByVal sta As String)
 btnSave.Enabled = sta
 btnCancle.Enabled = sta

 btnUpload.Enabled = False

End Sub

Private Sub StatelockCancle(ByVal sta As String)
 btnUpload.Enabled = sta

 btnSave.Enabled = False
 btnCancle.Enabled = False
End Sub

Private Sub Loadfiles()  '�����ͤ Import files
 Dim OpenFileDialog As New OpenFileDialog
 Dim strFullpart As String
 Dim strFilename As String

     With OpenFileDialog

          .CheckFileExists = True
          .ShowReadOnly = False
          .Filter = "All Files|*.*|����ٻ�Ҿ (*)|*.bmp;*.gif;*.jpg;*.png;*.pdf;*.xls;*.doc"
          .FilterIndex = 2

          Try

             If .ShowDialog = Windows.Forms.DialogResult.OK Then

                strFullpart = System.IO.Path.GetDirectoryName(.FileName) '�Ѻ���੾�оҸ���
                strFilename = New System.IO.FileInfo(.FileName).Name '�Ѻ���੾�Ъ������

                lblPath.Text = strFullpart
                lblFilename.Text = strFilename

                strStatus = "����Ң�����"

             End If

          Catch ex As Exception
                lblFilename.Text = ""
                txtName.Text = ""

          End Try

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

      strSqlSelc = "SELECT * FROM docsend (NOLOCK) "

      Rsd = New ADODB.Recordset

      With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

             .MoveLast()              '����͹��ѧ Record �ش����

             LastNumber = CInt(Mid(.Fields("doc_id").Value.ToString.Trim, 3))  '�Ѵʵ�ԧ ��� 4 ��Ƿ���  000x
             LastYear = Mid(.Fields("doc_id").Value.ToString.Trim, 1, 2)  '�Ѵ��һ�  5x ੾�� 2����á

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

          lblDocid.Text = "DC" & LastYear & LastNumber.ToString("0000")

      .ActiveConnection = Nothing
      .Close()
      End With

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub SaveData()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim strDateDoc As String
 Dim blnReturnCopyPic As Boolean
 Dim i As Integer
 Dim strDocid As String
 Dim strDepid As String = ""

     With Conn

          If .State Then Close()

             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

       strDocid = lblDocid.Text.ToString.Trim

       If chkDocidExist(strDocid) Then                                               '������ docid ���
          strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                   & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                   & Mid(txtBegin.Text.Trim, 1, 2)

             strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"                  '��� "" ���͡ѹ error


             '--------------------------- Copy Document ----------------------------------------------

             blnReturnCopyPic = CallCopyPicture(lblPath.Text.ToString.Trim, lblFilename.Text.ToString.Trim)

             If blnReturnCopyPic Then
                lblPath.Text = PthName

             Else
                txtName.Text = ""
                lblFilename.Text = ""

             End If

                   Select Case cboDepRecv.SelectedIndex

                          Case Is = 0
                               strDepid = "D1"

                          Case Is = 1
                               strDepid = "D2"

                          Case Is = 2
                               strDepid = "D3"

                          Case Is = 3
                               strDepid = "D4"

                          Case Is = 4
                               strDepid = "D5"

                          Case Is = 5
                               strDepid = "D6"

                   End Select


                  lblDocid.Text = Mid(lblDocid.Text, 3)

             strSqlCmd = "INSERT INTO docsend" _
                                & "(doc_sta,doc_id,filename,content_type" _
                                & ",send_date,send_by,recv_date,recv_by,remark,recv_dept" _
                                & ")" _
                                & " VALUES (" _
                                & "'" & "0" & "'" _
                                & ",'" & ReplaceQuote(lblDocid.Text.ToString.Trim) & "'" _
                                & ",'" & ReplaceQuote(lblFilename.Text.ToString.Trim) & "'" _
                                & ",'" & ReplaceQuote(txtName.Text.ToString.Trim) & "'" _
                                & "," & strDateDoc _
                                & ",'" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                                & "," & "NULL" _
                                & ",'" & "" & "'" _
                                & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                & ",'" & ReplaceQuote(strDepid) & "'" _
                                & ")"

             Conn.Execute(strSqlCmd)
             btnCancle.Enabled = False
             LoadData()              '�ʴ�������


                     '------------------------------�������ʷ��������������------------------------------------------

                     For i = 1 To dgvDoc.Rows.Count - 1

                             If dgvDoc.Rows(i).Cells(2).Value.ToString = lblDocid.Text.ToString.Trim Then    '��Ҥ������ Size � dgvSize �դ����ҡѺ txtSize
                                dgvDoc.CurrentCell = dgvDoc.Item(5, i)
                                dgvDoc.Focus()

                             End If
                     Next i
                             StateLockbtn(True)
                             StateLockEdit(False)

                     strStatus = ""  '�����������
                     ClearAlldata()

       Else
          MsgBox("�����͡��ë�� �ô���͡�������� ")
          ClearAlldata()
          txtName.Focus()

       End If

  Conn.Close()
  Conn = Nothing

End Sub

Private Function chkDocidExist(ByVal txtDocid As String) As Boolean
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

     strSqlSelc = "SELECT doc_id FROM docsend (NOLOCK)" _
                                & "WHERE doc_id = '" & txtDocid & "'"

     Rsd = New ADODB.Recordset

     With Rsd
          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
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

Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String) As Boolean
  Dim fname As String = String.Empty  '��ҡѺ ""
  Dim dFile As String = String.Empty
  Dim dFilePath As String = String.Empty

  Dim fServer As String = String.Empty
  Dim intResult As Integer   '�׹����繨ӹǹ���

  On Error GoTo Err70

        fname = strPicPath & "\" & strPicName  '���� \\10.32.0.15\data1\EquipDocument\"�����ٻ�Ҿ"
        fServer = PthName & "\" & strPicName    'partServer \\10.32.0.15\data1\EquipDocument\"�����ٻ�Ҿ"
        If File.Exists(fServer) Then    '�������������ԧ
           CallCopyPicture = True      '���׹��� true

        Else

            If File.Exists(fname) Then
               dFile = Path.GetFileName(fname)
               dFilePath = DrvName + dFile


               intResult = String.Compare(fname.ToString.Trim, dFilePath.ToString.Trim)

                    '-------------------------- ��Ҥ���� 0 �ʴ������Ŵ��������� �������ö Copy ����� ------------------------------

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

Private Sub btnCancle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancle.Click
 ClearAlldata()

 StateLockbtn(True)
 StateLockEdit(False)
End Sub

Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

     If dgvDoc.Rows.Count <> 0 Then

           If lblDocid.Text <> "" And txtName.Text <> "" Then

                   If ChkStaRecv() Then     '������͡��ö١�Ѻ������������ѧ
                      StateLockEdit(True)
                      strStatus = "��䢢�����"
                      LockEditData()

                   Else

                      MessageBox.Show("�͡��ö١�Ѻ������� �������ö�����", "�Դ��Ҵ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                   End If

            Else
                MsgBox("��س����͡��¡�� ��͹��䢢�����")
            End If

        StateLockbtn(True)

     Else
         MsgBox("����բ���������Ѻ���")

     End If
End Sub

Private Function ChkStaRecv() As Boolean           '������͡��ö١�Ѻ���������ѧ
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim docid As String
     docid = Mid(lblDocid.Text.ToString.Trim, 3)

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .CommandTimeout = 90
             .Open()
     End With

            strSqlSelc = "SELECT doc_sta FROM docsend (NOLOCK)" _
                                           & " WHERE doc_id = '" & docid & "'" _
                                           & " AND doc_sta = '1'"


     Rsd = New ADODB.Recordset

     With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

            If .RecordCount <> 0 Then

               Return False           '�͡��ö١�Ѻ�������

            Else

               Return True            '�͡����ѧ������Ѻ���

            End If

     .ActiveConnection = Nothing
     .Close()
     End With


  Conn.Close()
  Conn = Nothing

End Function

Private Sub SaveEditdata()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim strDateDoc As String
 Dim blnReturnCopyPic As Boolean

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With


              strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                   & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                   & Mid(txtBegin.Text.Trim, 1, 2)

              strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"    '��� "" ���͡ѹ error


             '--------------------------- Copy �͡��� ----------------------------------------------

             blnReturnCopyPic = CallCopyPicture(lblPath.Text.ToString.Trim, lblFilename.Text.ToString.Trim)

             If blnReturnCopyPic Then
                lblPath.Text = PthName

             Else
                txtName.Text = ""
                lblFilename.Text = ""

             End If

                     '---------------------------- UPDATE ������㹵��ҧ eqpmst ------------------------------

                      strSqlCmd = "UPDATE docsend SET filename ='" & ReplaceQuote(lblFilename.Text.ToString.Trim) & "'" _
                                           & "," & "content_type ='" & ReplaceQuote(txtName.Text.ToString.Trim) & "'" _
                                           & "," & "remark = '" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                           & " WHERE doc_id ='" & lblDocid.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)

                      LoadData()
                      strStatus = "" '��������ҵ����

 Conn.Close()
 Conn = Nothing

End Sub

Private Sub LockEditData()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim strDepnm As String = ""
 Dim strDoc As String
     strDoc = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim
     strDoc = Mid(strDoc, 3)

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

     strSqlSelc = "SELECT * FROM docsend (NOLOCK)" _
                                   & "WHERE doc_id = '" & strDoc & "'"

     Rsd = New ADODB.Recordset

     With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

           If .RecordCount <> 0 Then

               Select Case .Fields("recv_dept").Value.ToString.Trim

                          Case Is = "D1"
                               strDepnm = "- Ἱ��Ѵ�����ǹ (�س��д�ɰ� �ѧ��ͧ)"

                          Case Is = "D2"
                               strDepnm = "- Ἱ��մ EVA(�س�Է���ѡ��� �ҹ�ش��Ե���)"

                          Case Is = "D3"
                               strDepnm = "- Ἱ��մ PU(�س�Ⱦ� �����ع�ø���)"

                          Case Is = "D4"
                               strDepnm = "- Ἱ����(�سʶԵ�� �ʹ�ѡ)"

                          Case Is = "D5"
                               strDepnm = "- Ἱ���Ե��(�س൪Թ�� ����ŧ)"

                          Case Is = "D6"
                               strDepnm = "- Ἱ��մ PVC(�س���� �ʧ��س����ط���)"

                   End Select

              lblDocid.Text = "DC" & .Fields("doc_id").Value.ToString.Trim
              lblFilename.Text = .Fields("filename").Value.ToString.Trim
              txtName.Text = .Fields("content_type").Value.ToString.Trim
              cboDepRecv.Text = strDepnm
              txtRemark.Text = .Fields("remark").Value.ToString.Trim

              StateLockbtn(True)
              StateLockEdit(True)

           End If

     .ActiveConnection = Nothing
     .Close()
     End With

 Conn.Close()
 Conn = Nothing
End Sub

Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
 Dim strDocid As String
 Dim strContent As String
 Dim btyConsider As Byte

   If dgvDoc.Rows.Count <> 0 Then
      strDocid = Mid(dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim, 3)
      strContent = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(3).Value.ToString.Trim

      btyConsider = MsgBox("�����͡��� : DC" & strDocid.ToString.Trim & vbNewLine _
                                            & "��Ǣ���͡��� : " & strContent.ToString.Trim & vbNewLine _
                                            & "�س��ͧ���ź���������!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                            + MsgBoxStyle.Exclamation, "Confirm Delete Data")

            If btyConsider = 6 Then
               DeleteFileImport()      'ź�͡����  \\10.32.0.15\data1
               DeleteData(strDocid)
               ClearAlldata()

            End If

   Else
       MsgBox("����բ����������Թ���")

   End If
End Sub

Private Sub DeleteData(ByVal strDocid As String)
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

           Conn.BeginTrans()

           strSqlCmd = "DELETE docsend " _
                                & "WHERE doc_id = '" & strDocid & "'"

           Conn.Execute(strSqlCmd)

           Conn.CommitTrans()

               LoadData() '���ê������
               'dgvDoc.Focus()

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub dgvDoc_CellMouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvDoc.CellMouseUp
 Dim strDocid As String

  If dgvDoc.Rows.Count <> 0 Then
     strDocid = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim
     LockEditData()

     StateLockbtn(True)
     StateLockEdit(False)

  End If
End Sub

Private Sub dgvDoc_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvDoc.KeyUp
  Dim strDocid As String

  If dgvDoc.Rows.Count <> 0 Then
     strDocid = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim
     LockEditData()

     StateLockbtn(True)
     StateLockEdit(False)

  End If
End Sub

Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
  Me.Close()

End Sub

Private Sub txtName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyDown
   Dim intChkPoint As Integer
        With txtName
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���

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
                Case Is = 40  '����ŧ  
                    btnUpload.Focus()
                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRemark.Focus()
  End If
End Sub

Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
 Dim intChkPoint As Integer

        With txtRemark
            Select Case e.KeyCode
                Case Is = 35 '���� End 
                Case Is = 36 '���� Home
                Case Is = 37 '�١�ë���
                        If .SelectionStart = 0 Then
                           txtName.Focus()
                        End If
                Case Is = 38 '�����١�â��   
                        If .SelectionStart = 0 Then
                           txtName.Focus()
                        End If

                Case Is = 39 '�����١�â��
                    If .SelectionLength = .Text.Trim.Length Then  '��Ҥ�����ǵ��˹觻Ѩ�غѹ = ������Ǣͧ mskLdate
                        btnUpload.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     '��� InChkPoint = ������Ǣͧ  mskLdate
                        If .SelectionStart = intChkPoint Then    '��� Pointer ���价����˹��ش���¢ͧ mskLdate
                            btnUpload.Focus()
                        End If
                    End If

                Case Is = 40  '����ŧ  
                    btnUpload.Focus()

                Case Is = 113 '���� F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub txtRemark_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRemark.KeyPress
 If e.KeyChar = Chr(13) Then
    btnUpload.Focus()
 End If
End Sub

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  LoadData() '���ê������
End Sub

Private Sub Importdata()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim strDocid As String

     With Conn

          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()

     End With

         If dgvDoc.RowCount <> 0 Then
            strDocid = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim

            Conn.BeginTrans()
            strSqlCmd = "INSERT INTO tmp_docsend " _
                                & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                & " FROM docsend " _
                                & " WHERE doc_id = '" & strDocid & "' "

            Conn.Execute(strSqlCmd)
            Conn.CommitTrans()

         End If

  Conn.Close()
  Conn = Nothing
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
                    strSqlcmd = "DELETE tmp_docsend " _
                                & "WHERE user_id = '" & frmMainPro.lblLogin.Text & "'"
                    .Execute(strSqlcmd)

          End Select

       End With

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
  If txtName.Text <> "" Then

      Select Case strStatus

             Case Is = "����Ң�����"
                   SaveData() '�ѹ�֡����������

             Case Is = "��䢢�����"
                   SaveEditdata() '�ѹ�֡��䢢�����

      End Select

     ClearAlldata()
     btnUpload.Enabled = False

  Else
     MsgBox("��س��к� ��Ǣ���͡���")
     txtName.Focus()
  End If

End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
  ClearAlldata()
  GenDocid()
  strStatus = "����������"

  stateLockAdd(True)
  StateLockbtn(False)
  txtName.Focus()


End Sub

'----------------------- �Դ����͡��÷��Ṻ -----------------------------------------------------------
Private Sub dgvDoc_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDoc.CellDoubleClick
 Dim Path As String = DrvName
 Dim strDocid As String
 Dim strFilename As String

      strDocid = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim
      strFilename = lblFilename.Text.Trim
       Path = System.IO.Path.GetDirectoryName(Path) & "\" & strFilename
       Dim myProcess As System.Diagnostics.Process = New Process
       myProcess.StartInfo.FileName = Path
       myProcess.Start()

End Sub

Private Sub DeleteFileImport()   'ź�͡���� 10.32.0.15
 Dim Path As String = DrvName
 Dim strDocid As String
 Dim strFilename As String
      strDocid = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim
      strFilename = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(4).Value.ToString.Trim

       Path = System.IO.Path.GetDirectoryName(Path) & "\" & strFilename

       '--------------- ź�͡����  \\10.32.0.15\data1\Eqpdocument\

       If System.IO.File.Exists(Path) = True Then
          System.IO.File.Delete(Path)

       End If

End Sub
End Class