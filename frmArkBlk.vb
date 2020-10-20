Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmArkBlk
Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean    '�óա�ͧ������

Dim dubNumberStart As Double   '�١��˹� = 1
Dim dubNumberEnd As Double     '�١��˹� = 2100

Dim strSqlFindData As String
Dim strDocCode As String = "F4"

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim dsTn As New DataSet

Private Sub frmArkBlk_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

    If FormCount("frmAeArkBlk") > 0 Then

        With frmAeArkBlk

               strSearch = .lblComplete.Text

                If strSearch <> "" Then
                    SearchData(0, strSearch)
                End If

              .Close()

        End With

     Timer1.Enabled = True       '��� Timer1 ���ê˹�Ҩ�m�ء 5 �ҷ�

    End If

    Me.Height = Int(lblHeight.Text)
    Me.Width = Int(lblWidth.Text)

    Me.Top = Int(lblTop.Text)
    Me.Left = Int(lblLeft.Text)
End Sub

Private Sub frmArkBlk_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

  lblHeight.Text = Me.Height.ToString.Trim
  lblWidth.Text = Me.Width.ToString.Trim

  lblTop.Text = Me.Top.ToString.Trim
  lblLeft.Text = Me.Left.ToString.Trim

End Sub

Public Sub Activating()
   Dim strSearch As String

       strSearch = lblCmd.Text.Trim

       If strSearch <> "" Then
          SearchData(0, strSearch)
       End If

    Timer1.Enabled = False
End Sub

Private Sub frmArkBlk_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  Me.Dispose()
End Sub

    Private Sub frmArkBlk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        StdDateTimeThai()
        tlsBarFmr.Cursor = Cursors.Hand             '����������ç Toolstripbar ���ٻ���

        dubNumberStart = 1                          '�������á� Recordset = 1
        dubNumberEnd = 2100                         '�������á� Recordset = 2100

        PreGroupType()

        InputDeptData()
        tabCmd.Focus()

    End Sub

    Private Sub PreGroupType()

        Dim strGpTopic(2) As String
        Dim i As Integer

        strGpTopic(0) = "�����ػ�ó�"
        strGpTopic(1) = "��������´�ػ���"
        strGpTopic(2) = "ʶҹ��觽��¼�Ե"

        With cmbFilter

            For i = 0 To 2
                .Items.Add(strGpTopic(i))
            Next i

            .SelectedItem = .Items(0)
        End With

        With cmbType

            .Items.Add(strGpTopic(0))
        End With

    End Sub

    Private Function FormCount(ByVal fname As String) As Long

        Dim frm As Form

        For Each frm In My.Application.OpenForms

           If frm Is My.Forms.frmAeArkBlk Then

              FormCount = FormCount + 1     'return FormCount

           End If
      Next

End Function

    Private Sub InputDeptData()

        Dim Conn As New ADODB.Connection
        Dim Rsd As New Recordset
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
        Dim imgStaPrd As Image
        Dim imgStaFix As Image

        Dim strDateFilter As String = ""
        Dim strYearCnvt As String = ""

        With Conn

            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

        End With

        If blnHaveFilter Then          '�ó����͡ ��ͧ������

            Select Case cmbFilter.SelectedIndex()

                Case Is = 0
                    strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                Case Is = 1
                    strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                Case Is = 2
                    strFieldFilter = "sta_pd like '%" & ReplaceQuote(strSearch) & "%'"

            End Select

            strSqlCmdSelc = "SELECT  * FROM v_moldinj_hd (NOLOCK)" _
                                             & " WHERE " & strFieldFilter _
                                             & " AND [group] = 'G'" _
                                             & " ORDER BY eqp_id"
        Else


            strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                        & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                        & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                        & " AND [group] = 'G'" _
                                        & " ORDER BY eqp_id"

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

                dgvBlock.Rows.Clear()

                intCounter = 0

                Do While Not .EOF

                    '-------------------------------------------ʶҹ����ͺ���¼�Ե----------------------------------------------------------

                    Select Case .Fields("prod_sta").Value.ToString.Trim

                        Case Is = "1"
                            imgStaPrd = My.Resources._16x16_ledlightblue
                        Case Is = "2"
                            imgStaPrd = My.Resources._16x16_ledgreen
                        Case Is = "0"
                            imgStaPrd = My.Resources._16x16_ledred
                        Case Else
                            imgStaPrd = My.Resources.blank

                    End Select

                    '-------------------------------------------ʶҹ��觫���----------------------------------------------------------------

                    Select Case .Fields("fix_sta").Value.ToString.Trim

                        Case Is = "1"
                            imgStaFix = My.Resources.sign_deny
                        Case Is = "2"
                            imgStaFix = My.Resources.Chk
                        Case Else
                            imgStaFix = My.Resources.blank

                    End Select



                    dgvBlock.Rows.Add(
                                                  .Fields("eqp_id").Value.ToString.Trim,
                                                  .Fields("desc_thai").Value.ToString.Trim,
                                                  .Fields("eqp_name").Value.ToString.Trim,
                                                 imgStaPrd, .Fields("sta_pd").Value.ToString.Trim,
                                                   Mid(.Fields("pre_date").Value.ToString.Trim, 1, 10),
                                                  .Fields("pre_by").Value.ToString.Trim,
                                                   Mid(.Fields("last_date").Value.ToString.Trim, 1, 10),
                                                  .Fields("last_by").Value.ToString.Trim,
                                                  .Fields("Remark").Value.ToString.Trim
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

            End If

            .Close()

        End With

        Rsd = Nothing

        Conn.Close()
        Conn = Nothing

    End Sub

    Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchTXT As String)

        Dim Conn As New ADODB.Connection
        Dim Rsd As New ADODB.Recordset

        Dim intPageCount As Integer
        Dim intPageSize As Integer
        Dim strSqlCmdSelc As String = ""
        Dim i As Integer
        Dim strSqlFind As String = ""
        Dim strDateFilter As String = ""
        Dim strYearCnvt As String = ""

        With Conn

            If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

        End With

        Select Case bytColNumber


            Case Is = 0

                strSqlFind = "eqp_id "
                strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTXT) & "%'"

                'Case Is = 2
                '       strSqlFind = "eqp_name "
                '       strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTXT) & "%'"

                'Case Is = 3
                '       strSqlFind = "part_nw "
                '       strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTXT) & "%'"

                'Case Is = 4
                '       strSqlFind = "desc_eng "
                '       strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTXT) & "%'"

                'Case Is = 5
                '        strSqlFind = "desc_thai "
                '        strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTXT) & "%'"

                'Case Is = 7
                '        strSqlFind = "sta_pd "
                '        strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTXT) & "%'"

                'Case Is = 9
                '        strSqlFind = "sta_fx "
                '        strSqlFind = strSqlFind & "Like '%" & ReplaceQuote(strSearchTXT) & "%'"

        End Select

        strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                         & " WHERE " & strSqlFind _
                         & " AND [group]= 'G'" _
                         & " ORDER BY eqp_id"

        intPageSize = 30

        With Rsd

            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .Open(strSqlCmdSelc, Conn, , )

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
                '-------------------------------------------------------------------------------------------------------------------

                If .Fields("RowNumber").Value >= 2100 Then

                    dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                    dubNumberEnd = .Fields("RowNumber").Value + 1000

                Else

                    dubNumberStart = 1
                    dubNumberEnd = 2100

                End If

                strSqlFindData = strSqlFind

                InputDeptData()

                For i = 0 To dgvBlock.Rows.Count - 1
                    If InStr(UCase(dgvBlock.Rows(i).Cells(bytColNumber).Value), strSearchTXT.Trim.ToUpper) <> 0 Then
                        dgvBlock.CurrentCell = dgvBlock.Item(bytColNumber, i)
                        dgvBlock.Focus()
                        Exit For
                    End If
                Next i

            Else

                MsgBox("����բ����� : " & strSearchTXT & " ��к�" & vbNewLine _
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
        dgvBlock.Enabled = sta
        tlsBarFmr.Enabled = sta

    End Sub

    Private Sub dgvBlock_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
        dgvBlock.Rows(e.RowIndex).Height = 27
    End Sub

    Private Sub SearchDT()

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

    Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

        Dim btnReturn As Boolean

        With tabCmd

            Select Case tabCmd.SelectedIndex

                Case Is = 0    '����������

                    btnReturn = CheckUserEntry(strDocCode, "act_add")
                    If btnReturn Then

                        ClearTmpTableUser("tmp_eqptrn")
                        lblCmd.Text = "0"                     '�觺͡�������������

                        With frmAeArkBlk
                            .ShowDialog()
                            .Text = "����������"

                        End With
                    Else
                        MsnAdmin()
                    End If

                Case Is = 1    '��䢢�����

                    If dgvBlock.Rows.Count <> 0 Then

                        btnReturn = CheckUserEntry(strDocCode, "act_edit")
                        If btnReturn Then

                            ClearTmpTableUser("tmp_eqptrn")
                            lblCmd.Text = "1"                     '���͡�˹�����繡�����

                            With frmAeArkBlk
                                .ShowDialog()
                                .Text = "��䢢�����"

                            End With

                        Else
                            MsnAdmin()
                        End If

                    End If

                Case Is = 2    '����ͧ

                    If dgvBlock.Rows.Count <> 0 Then

                        btnReturn = CheckUserEntry(strDocCode, "act_view")
                        If btnReturn Then
                            ViewShoeData()

                        Else
                            MsnAdmin()
                        End If

                    End If

                Case Is = 3   '��ͧ������

                    If dgvBlock.Rows.Count <> 0 Then

                        With gpbFilter

                            .Top = 230
                            .Left = 210
                            .Width = 348
                            .Height = 125

                            .Visible = True

                            cmbFilter.Text = cmbFilter.Items(0)
                            txtFilter.Text =
                                      dgvBlock.Rows(dgvBlock.CurrentRow.Index).Cells(0).Value.ToString.Trim

                            StateLockFind(False)
                            txtFilter.Focus()

                        End With

                    End If

                Case Is = 4   '���Ң�����

                    If dgvBlock.Rows.Count <> 0 Then

                        With gpbSearch

                            .Top = 230
                            .Left = 210
                            .Width = 348
                            .Height = 125

                            .Visible = True

                            cmbType.Text = cmbType.Items(0)
                            txtSeek.Text =
                                    dgvBlock.Rows(dgvBlock.CurrentRow.Index).Cells(0).Value.ToString.Trim

                            StateLockFind(False)
                            txtSeek.Focus()

                        End With

                    End If

                Case Is = 5        '����������

                    If dgvBlock.Rows.Count > 0 Then

                        ClearTmpTableUser("tmp_eqptrn")

                        With gpbOptPrint
                            .Top = 230
                            .Left = 210
                            .Width = 374
                            .Height = 125

                            .Visible = True

                            InputEqpDataPrint()
                            cmbOptPrint.Text = dgvBlock.Rows(dgvBlock.CurrentRow.Index).Cells(0).Value.ToString.Trim()

                            StateLockFind(False)
                            cmbOptPrint.Focus()

                        End With
                    End If

                Case Is = 6           '��鹿٢�����

                    blnHaveFilter = False
                    InputDeptData()

                Case Is = 7          'ź

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

    Private Sub InputEqpDataPrint()
  Dim Conn As New ADODB.Connection
  Dim RsdPnt As New ADODB.Recordset

  Dim strSqlSelc As String

        With Conn

             If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .Open()

        End With

        strSqlSelc = "SELECT DISTINCT eqp_id FROM v_mst_trn (NOLOCK)" _
                                           & " WHERE [group] ='G'" _
                                           & " ORDER BY eqp_id"


       RsdPnt = New ADODB.Recordset

       With RsdPnt

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlSelc, Conn, , , )

               If .RecordCount <> 0 Then

                  ds.Clear()
                  da.Fill(ds, RsdPnt, "eqpid")
                  cmbOptPrint.DataSource = ds.Tables("eqpid").DefaultView
                  cmbOptPrint.DisplayMember = "eqp_id"
                  cmbOptPrint.ValueMember = "eqp_id"

                End If

               .ActiveConnection = Nothing
             ' .Close()

     End With

     RsdPnt = Nothing


End Sub

Private Sub DeleteData()
 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim btyConsider As Byte
 Dim strEqpid As String
 Dim strEqpname As String

    With Conn

              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

   End With

   With dgvBlock

        If .Rows.Count > 0 Then

             strEqpid = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
             strEqpname = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim

             btyConsider = MsgBox("�����ػ�ó� : " & strEqpid & vbNewLine _
                                                & "��������´�ػ�ó� : " & strEqpname & vbNewLine _
                                                & "�س��ͧ���ź���������!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then

                        Conn.BeginTrans()

                                '------------------------------------ź���ҧ eqpmst--------------------------------------------

                                strSqlCmd = "DELETE FROM eqpmst" _
                                                      & " WHERE eqp_id ='" & strEqpid & "'"

                                Conn.Execute(strSqlCmd)

                                '------------------------------------ź���ҧ eqptrn--------------------------------------------

                                strSqlCmd = "DELETE FROM eqptrn" _
                                                     & " WHERE eqp_id ='" & strEqpid & "'"

                         Conn.Execute(strSqlCmd)
                         Conn.CommitTrans()

                        .Rows.RemoveAt(.CurrentRow.Index)
                         InputDeptData()

                End If

        End If

      .Focus()

    End With

Conn.Close()
Conn = Nothing

End Sub

Private Sub ViewShoeData()

 If dgvBlock.Rows.Count <> 0 Then

     ClearTmpTableUser("tmp_eqptrn")
     lblCmd.Text = "2"

     With frmAeArkBlk
          .ShowDialog()
          .Text = "����ͧ������"

     End With

     'Me.Hide()
     'frmMainPro.Hide()

  Else
     MsnAdmin()
  End If

End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
 StateLockFind(True)
 gpbFilter.Visible = False
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

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        SearchDT()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        StateLockFind(True)
        gpbSearch.Visible = False
    End Sub

    Private Sub btnPrntPrevw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntPrevw.Click

        Dim strDocId As String = cmbOptPrint.Text.ToUpper.Trim

        If strDocId <> "" Then

            PrePrintData(strDocId)

            frmMainPro.lblRptCentral.Text = "E"          ' �觺͡�������§ҹ ���ͤ����

            '-------------------------�觤��������� lblRptDesc �ͧ����� MainPro ���� Userid �Ѻ Eqpid ----------------------------- 

            frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim _
                                                                    & "' AND eqp_id ='" & strDocId & "'"

            frmRptCentral.Show()

            StateLockFind(True)
            gpbOptPrint.Visible = False
            frmMainPro.Hide()

        Else
            MsgBox("�ô�кآ����š�͹�����", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Equipment Empty!!")
            cmbOptPrint.Focus()

        End If
    End Sub

    Private Sub PrePrintData(ByVal strSelectCode As String)
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim RsdPic As New ADODB.Recordset

 Dim strSqlSelc As String
 Dim strSqlCmdPic As String

 Dim strPicPath As String = "\\10.32.0.15\data1\EquipPicture\"
 Dim strLoadFilePic1 As String
 Dim strLoadFilePic2 As String
 Dim strLoadFilePic3 As String

 Dim blnHavePic1 As Boolean
 Dim blnHavePic2 As Boolean
 Dim blnHavePic3 As Boolean

 Dim inImg As Image

    With Conn

              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()

    End With

    strSqlSelc = "SELECT * " _
                        & " FROM v_mst_trn (NOLOCK)" _
                        & " WHERE eqp_id = '" & strSelectCode.ToString.Trim & "'"

     With Rsd

             .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             .LockType = ADODB.LockTypeEnum.adLockOptimistic
             .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then
                  For i As Integer = 1 To .RecordCount

                                       '----------------------------------------LoadPicture ���ͤʡ�չ ------------------------------------------------

                                       strLoadFilePic1 = strPicPath & .Fields("pic_ctain").Value.ToString.Trim
                                       If strLoadFilePic1 <> "" Then

                                               If File.Exists(strLoadFilePic1) Then '�ٻ�ѧ��������к�
                                                      blnHavePic1 = True
                                               Else
                                                      blnHavePic1 = False
                                                End If

                                       Else
                                            blnHavePic1 = False
                                       End If



                                         '----------------------------------------LoadPicture �ٻ�����ǹ ------------------------------------------------

                                        strLoadFilePic2 = strPicPath & .Fields("pic_io").Value.ToString.Trim
                                       If strLoadFilePic2 <> "" Then

                                               If File.Exists(strLoadFilePic2) Then '�ٻ�ѧ��������к�
                                                      blnHavePic2 = True
                                               Else
                                                      blnHavePic2 = False
                                                End If

                                       Else
                                            blnHavePic3 = False
                                       End If



                                          '----------------------------------------LoadPicture �ٻ�ͧ���������ٻ ------------------------------------------------

                                        strLoadFilePic3 = strPicPath & .Fields("pic_part").Value.ToString.Trim
                                       If strLoadFilePic3 <> "" Then

                                               If File.Exists(strLoadFilePic3) Then '�ٻ�ѧ��������к�
                                                      blnHavePic3 = True
                                               Else
                                                      blnHavePic3 = False
                                                End If

                                       Else
                                            blnHavePic3 = False
                                       End If


                                       strSqlCmdPic = "SELECT * " _
                                                                  & " FROM tmp_mst_trn (NOLOCK)"

                                       RsdPic = New ADODB.Recordset
                                       RsdPic.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                                       RsdPic.LockType = ADODB.LockTypeEnum.adLockOptimistic
                                       RsdPic.Open(strSqlCmdPic, Conn, , , )

                                                    RsdPic.AddNew()
                                                    RsdPic.Fields("user_id").Value = frmMainPro.lblLogin.Text.ToString.Trim
                                                    RsdPic.Fields("prod_sta").Value = .Fields("prod_sta").Value
                                                    RsdPic.Fields("fix_sta").Value = .Fields("fix_sta").Value
                                                    RsdPic.Fields("group").Value = .Fields("group").Value
                                                    RsdPic.Fields("eqp_id").Value = .Fields("eqp_id").Value
                                                    RsdPic.Fields("eqp_name").Value = .Fields("eqp_name").Value
                                                    RsdPic.Fields("pi").Value = .Fields("pi").Value
                                                    RsdPic.Fields("shoe").Value = .Fields("shoe").Value
                                                    RsdPic.Fields("part").Value = .Fields("part").Value
                                                    RsdPic.Fields("eqp_type").Value = .Fields("eqp_type").Value
                                                    RsdPic.Fields("ap_desc").Value = .Fields("ap_desc").Value
                                                    RsdPic.Fields("doc_ref").Value = .Fields("doc_ref").Value
                                                    RsdPic.Fields("set_qty").Value = .Fields("set_qty").Value
                                                    RsdPic.Fields("pic_ctain").Value = .Fields("pic_ctain").Value
                                                    RsdPic.Fields("pic_io").Value = .Fields("pic_io").Value
                                                    RsdPic.Fields("pic_part").Value = .Fields("pic_part").Value
                                                    RsdPic.Fields("remark").Value = .Fields("remark").Value
                                                    RsdPic.Fields("tech_desc").Value = .Fields("tech_desc").Value
                                                    RsdPic.Fields("tech_thk").Value = .Fields("tech_thk").Value
                                                    RsdPic.Fields("tech_trait").Value = .Fields("tech_trait").Value
                                                    RsdPic.Fields("tech_sht").Value = .Fields("tech_sht").Value
                                                    RsdPic.Fields("tech_eva").Value = .Fields("tech_eva").Value
                                                    RsdPic.Fields("tech_warm").Value = .Fields("tech_warm").Value
                                                    RsdPic.Fields("tech_time1").Value = .Fields("tech_time1").Value
                                                    RsdPic.Fields("tech_time2").Value = .Fields("tech_time2").Value
                                                    RsdPic.Fields("creat_date").Value = .Fields("creat_date").Value
                                                    RsdPic.Fields("size_desc").Value = .Fields("size_desc").Value
                                                    RsdPic.Fields("size_id").Value = .Fields("size_id").Value
                                                    RsdPic.Fields("size_qty").Value = .Fields("size_qty").Value
                                                    RsdPic.Fields("dimns").Value = .Fields("dimns").Value
                                                    RsdPic.Fields("sup_name").Value = .Fields("backgup").Value
                                                    RsdPic.Fields("men_rmk").Value = .Fields("men_rmk").Value
                                                    RsdPic.Fields("pre_date").Value = .Fields("pre_date").Value
                                                    RsdPic.Fields("pre_by").Value = .Fields("pre_by").Value
                                                    RsdPic.Fields("last_date").Value = .Fields("last_date").Value
                                                    RsdPic.Fields("last_by").Value = .Fields("last_by").Value
                                                    RsdPic.Fields("pi_qty").Value = .Fields("pi_qty").Value
                                                    RsdPic.Fields("eqp_amt").Value = .Fields("eqp_amt").Value


                                                    '----------------------------�����������ٻ���͡ʡ�չ---------------------------------------------------

                                                    If blnHavePic1 Then '������ٻ�Ҿ����ŧ�� Binary ������������������

                                                            Dim RsdSteam1 As New MemoryStream
                                                            Dim bytes1 = File.ReadAllBytes(strLoadFilePic1)

                                                            inImg = Image.FromFile(strLoadFilePic1)  '�֧�ٻ�����
                                                            inImg = ScaleImage(inImg, 161, 230)     '��Ѻ��Ҵ size
                                                            inImg.Save(RsdSteam1, ImageFormat.Bmp)  '����¹���ʡ�� .Bmp
                                                            bytes1 = RsdSteam1.ToArray
                                                            RsdPic.Fields("bob_ctain").Value = bytes1   '�ҷ�����������Ǵ� bob_ctain

                                                            RsdSteam1.Close()              '�Դ RecordSet
                                                            RsdSteam1 = Nothing            '������ RecordSet

                                                    End If


                                                    '---------------------------- �����������ٻ�Ҿ ��鹧ҹ ---------------------------------------------------

                                                    If blnHavePic2 Then '������ٻ�Ҿ����ŧ�� Binary ������������������

                                                            Dim RsdSteam2 As New MemoryStream
                                                            Dim bytes2 = File.ReadAllBytes(strLoadFilePic2)

                                                            inImg = Image.FromFile(strLoadFilePic2)
                                                            inImg = ScaleImage(inImg, 161, 230)
                                                            inImg.Save(RsdSteam2, ImageFormat.Bmp)
                                                            bytes2 = RsdSteam2.ToArray
                                                            RsdPic.Fields("bob_io").Value = bytes2

                                                            RsdSteam2.Close()
                                                            RsdSteam2 = Nothing

                                                    End If


                                                    '---------------------------- �����������ٻ�Ҿ��Ե�ѳ�� ---------------------------------------------------

                                                    If blnHavePic3 Then '������ٻ�Ҿ����ŧ�� Binary ������������������

                                                            Dim RsdSteam3 As New MemoryStream
                                                            Dim bytes3 = File.ReadAllBytes(strLoadFilePic3)

                                                            inImg = Image.FromFile(strLoadFilePic3)
                                                            inImg = ScaleImage(inImg, 161, 230)
                                                            inImg.Save(RsdSteam3, ImageFormat.Bmp)
                                                            bytes3 = RsdSteam3.ToArray
                                                            RsdPic.Fields("bob_part").Value = bytes3

                                                            RsdSteam3.Close()
                                                            RsdSteam3 = Nothing

                                                    End If

                                                    RsdPic.Update()

                                        RsdPic.ActiveConnection = Nothing
                                        RsdPic.Close()
                                        RsdPic = Nothing
                                  .MoveNext()     '����͹价�� Record �Ѵ�
                  Next i


                End If

            .ActiveConnection = Nothing
            .Close()

    End With
    Rsd = Nothing


Conn.Close()
Conn = Nothing

End Sub

Private Sub btnPrntCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrntCancel.Click
  StateLockFind(True)
  gpbOptPrint.Visible = False
End Sub

    Private Sub FilterData()    '��ͧ������

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
                    strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                Case Is = 1
                    strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                Case Is = 2
                    strFieldFilter = "sta_pd like '%" & ReplaceQuote(strSearch) & "%'"

            End Select


            strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                                  & " WHERE " & strFieldFilter _
                                                  & " ORDER BY eqp_id"


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
            End With

            Rsd = Nothing


            Conn.Close()
            Conn = Nothing

            If blnHaveData Then

                blnHaveFilter = True
                InputDeptData()

                StateLockFind(True)
                gpbFilter.Visible = False

            Else

                MsgBox("����բ����ŷ���ͧ��á�ͧ������!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
                txtFilter.Focus()

            End If

        Else

            MsgBox("�ô�кآ����ŷ���ͧ��á�ͧ��͹!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
            txtFilter.Focus()

        End If

End Sub

    Private Sub btnFilterCancel_Click_(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click

        If blnHaveFilter Then
            blnHaveFilter = False
            InputDeptData()
        End If

        StateLockFind(True)
        gpbFilter.Visible = False

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        InputDeptData()   '��觿�鹿٢�����
    End Sub

    Private Sub dgvBlock_RowsAdded1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvBlock.RowsAdded
        dgvBlock.Rows(e.RowIndex).Height = 27
    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click

        If txtFilter.Text.Length > 0 Then
            FilterData()
        End If

    End Sub

    Private Sub txtSeek_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSeek.KeyPress

        If e.KeyChar = Chr(13) And txtSeek.Text.Length > 0 Then
            SearchDT()
        End If

    End Sub

    Private Sub txtFilter_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFilter.KeyPress

        If e.KeyChar = Chr(13) And txtFilter.Text.Length > 0 Then
            FilterData()
        End If

    End Sub

End Class