Imports ADODB
Imports System.IO
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmApproveIssue
  Dim intBkPageCount As Integer
  Dim blnHaveFilter As Boolean    '�óա�ͧ������

  Dim dubNumberStart As Double   '�١��˹� = 1
  Dim dubNumberEnd As Double     '�١��˹� = 2100

  Dim strSqlFindData As String
  Dim strDocCode As String = "F11"

Private Sub frmApproveIssue_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

     InputDeptData()

     If FormCount("frmRptIssueNotify") > 0 Then

        With frmRptIssueNotify
             strSearch = .lblComplete.Text

             If strSearch <> "" Then

                SearchData(0, strSearch)

             End If

             .Close()

        End With

    End If


    Timer1.Enabled = True       '��� Timer1 ���ê˹�Ҩ�

    Me.Height = Int(lblHeight.Text)
    Me.Width = Int(lblWidth.Text)

    Me.Top = Int(lblTop.Text)
    Me.Left = Int(lblLeft.Text)

End Sub

Private Sub frmApproveIssue_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
 lblHeight.Text = Me.Height.ToString.Trim
 lblWidth.Text = Me.Width.ToString.Trim

 lblTop.Text = Me.Top.ToString.Trim
 lblLeft.Text = Me.Left.ToString.Trim
End Sub

Private Sub frmApproveIssue_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  Me.Dispose()

End Sub

Private Sub frmApproveIssue_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Me.WindowState = FormWindowState.Maximized     '���¢�Ҵ���˹�Ҩ�
  StdDateTimeThai()                           '���¡ �Ѻ�ٷչ StdDateTimeThai
  tlsBarFmr.Cursor = Cursors.Hand             '����������ç Toolstripbar ���ٻ���

  dubNumberStart = 1                          '�������á� Recordset = 1
  dubNumberEnd = 2100                         '�������á� Recordset = 2100

  PreviewUser(frmMainPro.lblLogin.Text.Trim)
  PreGroupType()
  InputDeptData()
  tabCmd.Focus()

End Sub

Private Function FormCount(ByVal frmName As String) As Long
 Dim frm As Form

     For Each frm In My.Application.OpenForms

         If frm Is My.Forms.frmRptIssueNotify Then
            FormCount = FormCount + 1
         End If

     Next

End Function

Sub PreviewUser(ByVal strUsr As String)

   Select Case strUsr

               Case Is = "PRADIST"
                    lblName.Text = "�س��д�ɰ� �ѧ��ͧ"
                    lblDept.Text = "122000 �Ѵ�����ǹ"

               Case Is = "ITHISAK"
                    lblName.Text = "�س�Է���ѡ��� �ҹ�ش��Ե���"
                    lblDept.Text = "125000 EVA INJECTION"

               Case Is = "TODSAPORN"
                    lblName.Text = "�س�Ⱦ� �����ع�ø���"
                    lblDept.Text = "126000 �մ PU"

               Case Is = "SATHID"
                    lblName.Text = "�سʶԵ�� �ʹ�ѡ"
                    lblDept.Text = "123000 ���"

               Case Is = "TECHIN"
                    lblName.Text = "�س൪Թ�� ����ŧ"
                    lblDept.Text = "121000 ��Ե��"

               Case Is = "PEERA"
                    lblName.Text = "�س���� �ʧ��س����ط���"
                    lblDept.Text = "124000 �մ PVC"

               Case Is = "BOONTUM"
                    lblName.Text = "�س�ح���� �������ʴ��"
                    lblDept.Text = "෤�Ԥ�ػ�ó�"

    End Select
End Sub

Private Sub PreGroupType()
 Dim strGpTopic(4) As String
 Dim i As Byte

     strGpTopic(0) = "�����͡���"
     strGpTopic(1) = "��������´�ػ���"           '�����ػ�ó�
     strGpTopic(2) = "Ἱ�����駻ѭ��"
     strGpTopic(3) = "��������´�����"            'DescThai
     strGpTopic(4) = "����ػ�ó�"                '�����ػ�ó�

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

Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchTxt As String)

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
                                 strSqlFind = "req_id "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 1
                                 strSqlFind = "eqp_nm "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 2
                                 strSqlFind = "from_dep "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 3
                                 strSqlFind = "desc_thai "
                                 strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                          Case Is = 4
                                  strSqlFind = "series "
                                  strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTxt) & "%'"

                    End Select


        strSqlCmdSelc = "SELECT * FROM v_IssueNotify (NOLOCK)" _
                                      & " WHERE " & strSqlFind _
                                      & " ORDER BY req_id"

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

                                        '---------------------------------------���Ң�����-------------------------------
                                        .MoveFirst()
                                        .Find(strSqlFind)
                                         lblPage.Text = Str(.AbsolutePage)
                                        '-----------------------------------------------------------------------------

                                        If .Fields("RowNumber").Value >= 2100 Then

                                            dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                                            dubNumberEnd = .Fields("RowNumber").Value + 1000

                                        Else

                                             dubNumberStart = 1
                                             dubNumberEnd = 2100

                                        End If

                                        strSqlFindData = strSqlFind

                                        InputDeptData()


                                                For i = 0 To dgvIssue.Rows.Count - 1

                                                        If InStr(UCase(dgvIssue.Rows(i).Cells(2).Value), strSearchTxt.Trim.ToUpper) <> 0 Then
                                                                dgvIssue.CurrentCell = dgvIssue.Item(bytColNumber, i)
                                                                dgvIssue.Focus()
                                                                Exit For
                                                        End If
                                                Next i

                          Else

                               MsgBox("����բ����� : " & strSearchTxt & " ��к�" & vbNewLine _
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

Private Sub InputDeptData()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String = ""
 Dim strDateAdd As String = ""
 Dim strDateEdit As String = ""

 Dim strInDate As String = ""

 Dim intPageCount As Integer
 Dim intPageSize As Integer
 Dim intCounter As Integer

 Dim strSearch As String = txtFilter.Text.ToString.Trim
 Dim strFieldFilter As String = ""

 Dim dteComputer As Date = Now()
 Dim imgStaReq As Image

 Dim strDateFilter As String = ""
 Dim strYearCnvt As String = ""
 Dim strApprovesta As String = ""

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
                                strFieldFilter = "req_id like '%" & ReplaceQuote(strSearch) & "'"

                           Case Is = 1
                                strFieldFilter = "eqpnm like '%" & ReplaceQuote(strSearch) & "'"

                           Case Is = 2
                                strFieldFilter = "from_dep like '%" & ReplaceQuote(strSearch) & "'"

                           Case Is = 3
                                strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "'"

                           Case Is = 4
                                strFieldFilter = "shoe like '%" & ReplaceQuote(strSearch) & "'"

                    End Select


                              Select Case frmMainPro.lblLogin.Text

                                     Case Is = "PRADIST"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND dep_notify = '" & "122000 Ἱ��Ѵ�����ǹ" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                                     Case Is = "ITHISAK"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND dep_notify = '" & "125000 Ἱ��մ EVA INJECTION" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                                     Case Is = "TODSAPORN"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND dep_notify = '" & "126000 Ἱ��մ PU" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                                     Case Is = "SATHID"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND dep_notify = '" & "123000 Ἱ����" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                                     Case Is = "TECHIN"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND dep_notify = '" & "121000 Ἱ���Ե��" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"
                                     Case Is = "PEERA"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND dep_notify = '" & "124000 Ἱ��մ PVC" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                                     Case Is = "BOONTUM"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE " & strFieldFilter _
                                                                 & " AND person3_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                              End Select

              Else

                    Select Case frmMainPro.lblLogin.Text

                           Case Is = "PRADIST"

                                       strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND dep_notify = '" & "122000 Ἱ��Ѵ�����ǹ" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                           Case Is = "ITHISAK"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND dep_notify = '" & "125000 Ἱ��մ EVA INJECTION" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                           Case Is = "TODSAPORN"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND dep_notify = '" & "126000 Ἱ��մ PU" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                           Case Is = "SATHID"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND dep_notify = '" & "123000 Ἱ����" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                           Case Is = "TECHIN"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND dep_notify = '" & "121000 Ἱ���Ե��" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                           Case Is = "PEERA"

                                         strSqlCmdSelc = "SELECT  * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND dep_notify = '" & "124000 Ἱ��մ PVC" & "'" _
                                                                 & " AND person1_sta = 'True' " _
                                                                 & " ORDER BY req_id"

                           Case Is = "BOONTUM"

                                         strSqlCmdSelc = "SELECT * FROM v_notifyissue (NOLOCK)" _
                                                                 & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                                 & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                                 & " AND person3_sta = 'True' " _
                                                                 & " ORDER BY req_id"


                   End Select


             End If

              intPageSize = 30   '����á�˹���Ҵ��д��

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

                                    '--------------------------����ա�ä���----------------------------------------

                                     If strSqlFindData <> "" Then

                                        .MoveFirst()
                                        .Find(strSqlFindData)

                                             If Not .EOF Then
                                                lblPage.Text = Str(.AbsolutePage)
                                             End If

                                            strSqlFindData = ""

                                    End If

                                    '------------------------------------------------------------------------------------

                                    If Int(lblPage.Text.ToString) > intPageCount Then
                                        lblPage.Text = intPageCount.ToString
                                    End If

                                    txtPage.Text = lblPage.Text.ToString
                                    intBkPageCount = .PageCount
                                    lblPageAll.Text = "/ " & .PageCount.ToString
                                    .AbsolutePage = Int(lblPage.Text.ToString)

                                    dgvIssue.Rows.Clear()

                                    intCounter = 0
                                    lblDocnull.Visible = False

                                    Do While Not .EOF

                                     '-------------------------------------------ʶҹ����ͺ���¼�Ե---------------------------

                                                 If frmMainPro.lblLogin.Text = "BOONTUM" Then

                                                            Select Case .Fields("person4_sta").Value

                                                                   Case Is = False
                                                                       imgStaReq = My.Resources._16x16_ledred

                                                                   Case Else
                                                                       imgStaReq = My.Resources._16x16_ledgreen

                                                            End Select

                                                                  '----------- ��Ǩ�ͺ��� ���͹��ѵ��Ѻ�� �׹�ѹ�����ѧ -----------------------

                                                                  Select Case .Fields("person4_sta").Value

                                                                         Case Is = True
                                                                              strApprovesta = "��͹��ѵ�����"

                                                                         Case Else
                                                                             strApprovesta = "����͹��ѵ�"

                                                                  End Select


                                                 Else

                                                          Select Case .Fields("person2_sta").Value

                                                                 Case Is = False
                                                                     imgStaReq = My.Resources._16x16_ledred

                                                                 Case Else
                                                                      imgStaReq = My.Resources._16x16_ledgreen

                                                          End Select


                                                                 '----------- ��Ǩ�ͺ��� ���͹��ѵ��Ѻ�� �׹�ѹ�����ѧ -----------------------

                                                                  Select Case .Fields("person2_sta").Value

                                                                         Case Is = True
                                                                              strApprovesta = "��͹��ѵ�����"

                                                                         Case Else
                                                                             strApprovesta = "����͹��ѵ�"

                                                                  End Select

                                       End If


                                            dgvIssue.Rows.Add( _
                                                                  imgStaReq, strApprovesta, _
                                                                  .Fields("req_id").Value.ToString.Trim, _
                                                                  .Fields("dep_notify").Value.ToString.Trim, _
                                                                  .Fields("eqpnm").Value.ToString.Trim, _
                                                                  .Fields("shoe").Value.ToString.Trim & " / " & .Fields("size").Value.ToString.Trim, _
                                                                  .Fields("amount").Value, _
                                                                  .Fields("issue").Value.ToString.Trim & " / " & .Fields("cause").Value.ToString.Trim, _
                                                                  .Fields("fxissue").Value.ToString.Trim, _
                                                                  Mid(.Fields("person1_date").Value.ToString.Trim, 1, 10), _
                                                                  .Fields("person1").Value.ToString.Trim, _
                                                                  Mid(.Fields("last_date").Value.ToString.Trim, 1, 10), _
                                                                  .Fields("lastby").Value.ToString.Trim, _
                                                                  .Fields("remark").Value.ToString.Trim _
                                                            )
                                            intCounter = intCounter + 1

                                              If intCounter = intPageSize Then
                                                    Exit Do
                                              End If

                                         .MoveNext()    '����价������¹����

                                     Loop

                            Else

                                intBkPageCount = 1
                                txtPage.Text = "1"

                                lblDocnull.Visible = True

                            End If

                   .Close()

              End With

  Rsd = Nothing

 Conn.Close()
 Conn = Nothing

End Sub

Private Sub StateLockFind(ByVal Sta As Boolean)
 tabCmd.Enabled = Sta
 dgvIssue.Enabled = Sta
 tlsBarFmr.Enabled = Sta

End Sub

Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
 lblPage.Text = "1"
 InputDeptData()
End Sub

Private Sub btnPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPre.Click
  If Int(lblPage.Text) > 1 Then
     lblPage.Text = Str(Int(lblPage.Text) - 1)
     InputDeptData()
  End If

End Sub

Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
 lblPage.Text = Str(Int(lblPage.Text) + 1)
 InputDeptData()
End Sub

Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
 lblPage.Text = Str(intBkPageCount)
 InputDeptData()
End Sub

Private Sub txtPage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPage.GotFocus
 txtPage.SelectAll()
End Sub

Private Sub txtPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPage.KeyPress
 If e.KeyChar = Chr(13) Then
     dgvIssue.Focus()
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
                                    InputDeptData()

                            Else

                                    lblPage.Text = intMovePage.ToString.Trim
                                    txtPage.Text = lblPage.Text
                                    InputDeptData()

                            End If
                    Else

                        If intMovePage > 0 Then
                            lblPage.Text = intMovePage.ToString.Trim
                            txtPage.Text = lblPage.Text
                        Else
                            lblPage.Text = "1"
                            txtPage.Text = lblPage.Text
                        End If

                        InputDeptData()

                    End If

                Catch ex As Exception
                    txtPage.Text = lblPage.Text
                End Try

End Sub

Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click
 Dim blnReturn As Boolean
 Dim strEqpId As String = ""

     With tabCmd

          Select Case .SelectedIndex

                    Case Is = 0 '͹��ѵ� / �����

                            blnReturn = CheckUserEntry(strDocCode, "act_print")
                            If blnReturn Then

                               If dgvIssue.Rows.Count <> 0 Then

                                  ClearTmpTableUser("tmp_notifyissue")
                                  frmMainPro.lblRptDesc.Text = Me.dgvIssue.Rows(Me.dgvIssue.CurrentRow.Index).Cells(2).Value.ToString

                                  With frmRptIssueNotify
                                       .Show()

                                  End With
                                  Me.Hide()
                                  frmMainPro.Hide()

                               End If

                            Else
                                MsnAdmin()
                            End If

                    Case Is = 1    '��ͧ������

                            If dgvIssue.Rows.Count > 0 Then

                                With gpbFilter

                                     .Top = 230
                                     .Left = 210
                                     Width = 348
                                     .Height = 125

                                     .Visible = True

                                     cmbFilter.SelectedItem = cmbFilter.Items(0)
                                     txtFilter.Text = _
                                              dgvIssue.Rows(dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                                     StateLockFind(False)
                                     txtFilter.Focus()

                                 End With

                            End If

                    Case Is = 2 '���Ң�����

                            If dgvIssue.Rows.Count > 0 Then

                               With gpbSearch

                                    .Top = 230
                                    .Left = 210
                                    .Width = 348
                                    .Height = 125

                                    .BringToFront()
                                    .Visible = True

                                    cmbType.SelectedItem = cmbType.Items(0)
                                    txtSeek.Text = _
                                             dgvIssue.Rows(dgvIssue.CurrentRow.Index).Cells(2).Value.ToString.Trim()

                                    StateLockFind(False)
                                    txtSeek.Focus()

                               End With

                            End If

                    Case Is = 3 '��鹿٢�����
                              InputDeptData()

                    Case Is = 4 '�͡
                            Me.Close()

          End Select

  End With
End Sub

Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
  FilterData()
End Sub

Sub FilterData()    '��ͧ������

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlCmdSelc As String = ""

  Dim strFieldFilter As String = ""
  Dim blnHaveData As Boolean
  Dim strSearch As String = txtFilter.Text.ToUpper.Trim

  Dim strDateFilter As String = ""
  Dim strYearCnvt As String = ""

      If strSearch <> "" Then


           With Conn

              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

           End With

                    Select Case cmbFilter.SelectedIndex()

                               Case Is = 0
                                     strFieldFilter = "req_id like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 1
                                      strFieldFilter = "eqpnm like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 2
                                      strFieldFilter = "dep_notify like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 3
                                      strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "'"

                               Case Is = 4
                                      strFieldFilter = "shoe like '%" & ReplaceQuote(strSearch) & "'"


                    End Select


                    strSqlCmdSelc = "SELECT * FROM v_notifyissue (NOLOCK)" _
                                                  & " WHERE " & strFieldFilter _
                                                  & " ORDER BY req_id"


                    Rsd = New ADODB.Recordset

                    With Rsd

                            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                            .LockType = ADODB.LockTypeEnum.adLockOptimistic
                            .Open(strSqlCmdSelc, Conn, , , )

                            If .RecordCount <> 0 Then
                                blnHaveData = True
                            Else
                                blnHaveData = False
                            End If

                           .ActiveConnection = Nothing
                           .Close()


                   Rsd = Nothing
                   'Rsd.Close()

                   End With

                         If blnHaveData Then

                            blnHaveFilter = True
                            InputDeptData()

                            StateLockFind(True)
                            gpbFilter.Visible = True

                         Else

                             MsgBox("����բ����ŷ���ͧ��á�ͧ������!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
                             txtFilter.Focus()

                         End If

        Else

           MsgBox("�ô�кآ����ŷ���ͧ��á�ͧ��͹!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
           txtFilter.Focus()

        End If

    Conn.Close()
    Conn = Nothing

End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click
 If blnHaveFilter Then
    blnHaveFilter = False
    InputDeptData()

 End If

    StateLockFind(True)
    gpbFilter.Visible = False

End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
  FindDocID()
End Sub

Private Sub FindDocID() '�����͡���

 Dim strSearch As String = txtSeek.Text.ToUpper.Trim

   If strSearch <> "" Then

        Select Case cmbType.SelectedIndex()

               Case Is = 0 '����
                    SearchData(0, strSearch)     '�觵����͹� ,Text ��� �Ѻ�ٷչ SearchData

               Case Is = 1 '��������´�ػ�ó�
                    SearchData(1, strSearch)

               Case Is = 2 'Ἱ�����駻ѭ��
                    SearchData(2, strSearch)

               Case Is = 3 '��������´�����
                    SearchData(3, strSearch)

               Case Is = 4 '����ػ�ó�
                    SearchData(4, strSearch)

        End Select

   Else

    MsgBox("�ô�кآ��������ͤ���!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
    txtSeek.Focus()

End If

End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
  StateLockFind(True)
  gpbSearch.Visible = False

End Sub

Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
 InputDeptData()
End Sub

Private Sub dgvIssue_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvIssue.RowsAdded
  dgvIssue.Rows(e.RowIndex).Height = 27
End Sub
End Class