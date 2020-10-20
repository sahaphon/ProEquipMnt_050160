Imports ADODB
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Image
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmFixEqpmnt

Dim intBkPageCount As Integer
Dim blnHaveFilter As Boolean    'กรณีกรองข้อมูล
Dim IsShowSeek As Boolean

Dim dubNumberStart As Double   'ถูกกำหนด = 1
Dim dubNumberEnd As Double     'ถูกกำหนด = 2100

Dim strSqlFindData As String
Dim strDocCode As String = "F6"

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim dsTn As New DataSet

Dim staPrint As String = ""          'สถานะการพิมพ์
Dim strOperation As String           'กำหนดว่าเป็นการ Search หรือ Filter

Private Sub frmFixEqpmnt_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

 Dim strSearch As String

     If FormCount("frmAeFixEqpmnt") > 0 Then

        With frmAeFixEqpmnt

             strSearch = .lblComplete.Text     'รหัสซ่อม
             strOperation = "2"   'ใช้บอกว่ามาจากการเพิ่มข้อมูล

             If strSearch <> "" Then
                SearchData(0, strSearch)
             End If

           .Close()
        End With

       Timer1.Enabled = True          'สั่งรีเฟรชข้อมูลทุก 1 นาที

     End If

    Me.Height = Int(lblHeight.Text)
    Me.Width = Int(lblWidth.Text)

    Me.Top = Int(lblTop.Text)
    Me.Left = Int(lblLeft.Text)

End Sub

Private Function FormCount(ByVal frmName As String) As Long

  Dim frm As Form
      For Each frm In My.Application.OpenForms

          If frm Is My.Forms.frmAeFixEqpmnt Then
             FormCount = FormCount + 1
          End If

      Next

End Function

Public Sub Activating()
   Dim strSearch As String
       strSearch = lblCmd.Text.Trim

       If strSearch <> "" Then
          SearchData(0, strSearch)
       End If

    Timer1.Enabled = False
End Sub

Private Sub frmFixEqpmnt_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
  lblHeight.Text = Me.Height.ToString
  lblWidth.Text = Me.Width.ToString

  lblTop.Text = Me.Top.ToString
  lblLeft.Text = Me.Left.ToString
End Sub

Private Sub frmFixEqpmnt_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTableUser("tmp_v_fixeqptrn")
   Me.Dispose()
End Sub

    Private Sub frmFixEqpmnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized     'ขยายขนาดเต็มหน้าจอ
        StdDateTimeThai()                           'เรียก ซับรูทีน StdDateTimeThai
        tlsBarFmr.Cursor = Cursors.Hand             'ให้คอร์เซอร์ตรง Toolstripbar เป็นรูปมือ

        dubNumberStart = 1                          'ให้แถวเเรกใน Recordset = 1
        dubNumberEnd = 2100                         'ให้แถวเเรกใน Recordset = 2100

        PreGroupType()
        InputData()
        tabCmd.Focus()

    End Sub

    Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchtxt As String)

        Dim Conn As New ADODB.Connection
        Dim Rsd As New ADODB.Recordset

        Dim intPageCount As Integer
        Dim intPageSize As Integer
        Dim strSqlCmdSelc As String = ""

        Dim strSqlFind As String = ""
        Dim i As Integer

        Dim strDateFilter As String = ""
        Dim strYearCnvt As String = ""
        Dim BytCellPoint As Byte  'ตัวเเปรชี้ไปที่สิ่งที่ต้องการค้นหา

        With Conn
            If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
        End With

        Select Case bytColNumber

            Case Is = 0

                strSqlFind = "fix_id "
                      strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

            Case Is = 1

                strSqlFind = "fix_desc"
                      strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchtxt) & "%'"

           End Select

           strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                        & " WHERE " & strSqlFind _
                                        & " ORDER BY eqp_id"

        intPageSize = 30

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

                      '---------------- ค้นหาข้อมูล ------------------

                      .MoveFirst()
                      .Find(strSqlFind)
                      lblPage.Text = Str(.AbsolutePage)

                      '--------------------------------------------

                      If .Fields("RowNumber").Value >= 2100 Then
                         dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                         dubNumberEnd = .Fields("RowNumber").Value + 1000

                      Else
                          dubNumberStart = 1
                          dubNumberEnd = 2100

                      End If

                         strSqlFindData = strSqlFind


                '---------------- เลือก Cell ที่จะ point เมื่อค้นหา --------------------

                If strOperation = "0" Then     'ค้นหาข้อมูล

                    Select Case cmbType.SelectedIndex

                        Case Is = 0
                            BytCellPoint = 2

                        Case Is = 1
                            BytCellPoint = 1

                    End Select

                Else     'กรณ๊มาจากการเพิ่มข้อมูล
                    BytCellPoint = "16"
                             InputData()

                         End If
                             InputData()

                                      For i = 0 To dgvFix.Rows.Count - 1

                                            If InStr(UCase(dgvFix.Rows(i).Cells(BytCellPoint).Value), strSearchtxt.Trim.ToUpper) <> 0 Then      'UCase เป็นฟังก์ชั่น เเปลงสตริงเป็น ตัวพิมพ์เล็ก พิมพ์ใหญ่
                                               dgvFix.CurrentCell = dgvFix.Item(BytCellPoint, i)
                                               dgvFix.Focus()
                                               Exit For

                                            End If

                                       Next i
                                       strOperation = ""  'เคลี่่ยร์ค่าตัวเเปร
                                       BytCellPoint = 0

                Else

                     MsgBox("ไม่พบข้อมูล  " & cmbFilter.Text & " = " & strSearchtxt & " ในระบบ" & vbNewLine _
                                        & "โปรดระบุการค้นหาข้อมูลใหม่!", vbExclamation, "Not Found Data")

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

Sub FilterData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset
Dim strSqlCmdSelc As String = ""
Dim strFieldFilter As String = ""

Dim blnHaveData As Boolean
Dim strSearch As String = txtFilter.Text.ToUpper.Trim
Dim strDateFilter As String = ""
Dim strYearCnvt As String = ""

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
   End With

    If strSearch <> "" Then
       lblFree1.Text = cmbFilter.SelectedIndex()   'เก็บตัวเลือกการกรอง
       lblFree2.Text = txtFilter.Text.Trim    'รายละเอียดกรองข้อมูล

        Select Case cmbFilter.SelectedIndex()

               Case Is = 0
                    strFieldFilter = "fix_id like '%" & ReplaceQuote(strSearch) & "%'"

               Case Is = 1
                    strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                Case Is = 2
                    strFieldFilter = "fix_desc like '%" & ReplaceQuote(strSearch) & "%'"

        End Select


                    strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                            & " WHERE " & strFieldFilter _
                                            & " ORDER BY eqp_id"


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

        Conn.Close()
        Conn = Nothing

        If blnHaveData Then
           blnHaveFilter = True
           InputData()

           StateLockFind(True)
           gpbFilter.Visible = False
        Else
              MsgBox("ไม่มีข้อมูลที่ต้องการกรองข้อมูล!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
              txtFilter.Focus()
        End If

Else
    MsgBox("โปรดระบุข้อมูลที่ต้องการกรองก่อน!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
    txtFilter.Focus()
End If

End Sub

Private Sub StateLockFind(ByVal sta As Boolean)
  tabCmd.Enabled = sta
  dgvFix.Enabled = sta
  tlsBarFmr.Enabled = sta
End Sub

Private Sub PreGroupType()

        Dim strGpTopic(2) As String
        Dim i As Byte

        strGpTopic(0) = "รหัสส่งซ่อม"
        strGpTopic(1) = "รหัสอุปกรณ์"
        strGpTopic(2) = "สถานะส่งซ่อม"

        With cmbType

            For i = 0 To 2

                If i <> 1 Then
                    .Items.Add(strGpTopic(i))
                End If

            Next i

            .SelectedItem = .Items(0)

      End With

          With cmbFilter

            For i = 0 To 2
                .Items.Add(strGpTopic(i))
            Next i

            .SelectedItem = .Items(0)

         End With

End Sub

Private Sub InputData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String = ""
Dim strDateAdd As String = ""
Dim strDateEdit As String = ""

Dim strInDate As String = ""

Dim intPageCount As Integer          'จำนวนหน้าทั้งหมด
Dim intPageSize As Integer           'จำนวนรายการใน 1 หน้า
Dim intCounter As Integer

Dim strSearch As String = txtFilter.Text.ToString.Trim
Dim strFieldFilter As String = ""

Dim dteComputer As Date = Now()
Dim imgStaFix As Image               'รูปสถานะส่งซ่อม

Dim strDateFilter As String = ""
Dim strYearCnvt As String = ""

    With Conn
        If .State Then .Close()
           .ConnectionString = strConnAdodb
           .CursorLocation = ADODB.CursorLocationEnum.adUseClient
           .ConnectionTimeout = 90
           .Open()
   End With

       If blnHaveFilter Then          'กรณีเลือก กรองข้อมูล
          strSearch = lblFree2.Text

           Select Case CInt(lblFree1.Text)

                Case Is = 0
                    strFieldFilter = "fix_id like '%" & ReplaceQuote(strSearch) & "%'"

                Case Is = 1
                    strFieldFilter = "fix_desc like '" & ReplaceQuote(strSearch) & "%'"

           End Select

                   strSqlCmdSelc = "SELECT * FROM v_fixeqptrn " _
                                            & " WHERE " & strFieldFilter _
                                            & " ORDER BY eqp_id"

        Else

                           strSqlCmdSelc = "SELECT * FROM v_fixeqptrn " _
                                                  & " WHERE RowNumber >= " & dubNumberStart.ToString.Trim _
                                                  & " AND RowNumber <= " & dubNumberEnd.ToString.Trim _
                                                  & " ORDER BY eqp_id"

       End If

              intPageSize = 30   'ตัวแปรกำหนดขนาดกระดาษ

        With Rsd
            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlCmdSelc, Conn, , , )

            If .RecordCount <> 0 Then
                If intPageSize > .RecordCount Then    'ถ้าจำนวนรายการใน 1 page(30) > จำนวนเรคคอร์ดที่ qurey มา
                    intPageSize = .RecordCount
                End If

                If intPageSize = 0 Then
                    intPageSize = 30
                End If

                .PageSize = intPageSize              '.PageSize ใช้กำหนดว่าแต่ละหน้าจะให้มีกี่รายการ ในการแสดงผล
                intPageCount = .PageCount            '.PageCount นับจำนวนหน้าทั้งหมด ที่ได้จากการกำหนดขนาดของหน้า

                '-------------------------- กรณีมีการค้นหา ----------------------------------------

                If strSqlFindData <> "" Then
                    .MoveFirst()
                    .Find(strSqlFindData)

                    If Not .EOF Then
                        lblPage.Text = Str(.AbsolutePage)    '.AbsolutePage ใช้อ้างอิงไปยังหน้าที่ต้องการ
                    End If

                    strSqlFindData = ""
                End If

                '------------------------ กำหนดปุ่ม ใน tlsBarFmr --------------------------------

                If Int(lblPage.Text.ToString) > intPageCount Then
                    lblPage.Text = intPageCount.ToString
                End If

                txtPage.Text = lblPage.Text.ToString
                intBkPageCount = .PageCount
                lblPageAll.Text = "/ " & .PageCount.ToString
                .AbsolutePage = Int(lblPage.Text.ToString)

                intCounter = 0
                dgvFix.Rows.Clear()

                Do While Not .EOF

                    '--------------------------- สถานะส่งซ่อม ----------------------

                    Select Case .Fields("fix_sta").Value.ToString.Trim

                        Case Is = "1"     'ส่งซ่อม
                            imgStaFix = My.Resources._16x16_ledred

                        Case Is = "2"     'รับคืนส่งซ่อม
                            imgStaFix = My.Resources._16x16_ledgreen

                        Case Is = "3"     'รับคืนบางส่วน
                            imgStaFix = My.Resources._16x16ledyellow

                        Case Else         'ปกติ
                            imgStaFix = My.Resources.blank

                    End Select

                    dgvFix.Rows.Add(
                                                                 imgStaFix,
                                                                 .Fields("fix_desc").Value.ToString.Trim,
                                                                 .Fields("fix_id").Value.ToString.Trim,
                                                                 .Fields("eqp_id").Value.ToString.Trim,
                                                                 "#" & .Fields("size_id").Value.ToString.Trim,
                                                                 .Fields("desc_thai").Value.ToString.Trim,
                                                                 .Fields("amt_out").Value.ToString.Trim,
                                                                 .Fields("amt_in").Value.ToString.Trim,
                                                                 Format(.Fields("price").Value, "#,##0.00"),
                                                                 .Fields("issue").Value.ToString.Trim,
                                                                 .Fields("pr_doc").Value.ToString.Trim,
                                                                  Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10),
                                                                 .Fields("fix_by").Value.ToString.Trim,
                                                                  Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10),
                                                                 .Fields("recv_by").Value.ToString.Trim,
                                                                  Mid(.Fields("pre_date").Value.ToString.Trim, 1, 10),
                                                                 .Fields("pre_by").Value.ToString.Trim
                                                            )

                    intCounter = intCounter + 1
                    If intCounter = intPageSize Then
                        Exit Do
                    End If

                    .MoveNext()    'ข้ามไปที่ระเบียนใหม่
                Loop

            Else

                dgvFix.Rows.Clear()  'เคลียร์ grid
                intBkPageCount = 1
                txtPage.Text = "1"

            End If

            blnHaveFilter = False
            strSearch = ""

            .Close()

        End With
        Rsd = Nothing

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub StateLockFindDept(ByVal sta As Boolean)
   tabCmd.Enabled = sta
   dgvFix.Enabled = sta
   tlsBarFmr.Enabled = sta
End Sub

Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

 Dim btnReturn As Boolean

    With tabCmd

         Select Case tabCmd.SelectedIndex

                Case Is = 0     'เพิ่มข้อมูล

                     btnReturn = CheckUserEntry(strDocCode, "act_add")  'ฟังก์ชั่นตรวจสอบสิทธิ์ strDocCode = 'F0', act_add คือสิทธิ์การ เพิ่มข้อมูล
                     If btnReturn Then

                        ClearTmpTableUser("tmp_fixeqptrn")              'เคลียร์ข้อมูลใน ตาราง tmp
                        lblCmd.Text = "0"

                        With frmAeFixEqpmnt
                             .ShowDialog()
                             .Text = "เพิ่มข้อมูล"

                        End With

                     Else
                         MsnAdmin()     'message คุณไม่มีสิทธิใช้ข้อมูลส่วนนี้
                     End If

               Case Is = 1  'แก้ไขข้อมูล

                    If dgvFix.Rows.Count <> 0 Then

                         btnReturn = CheckUserEntry(strDocCode, "act_edit")
                         If btnReturn Then

                            ClearTmpTableUser("tmp_fixeqptrn")
                            lblCmd.Text = "1"                     'เพื่อกำหนดว่าเป็นการแก้ไข

                            With frmAeFixEqpmnt
                                 .ShowDialog()
                                 .Text = "แก้ไขข้อมูล"

                            End With

                        Else
                            MsnAdmin()
                         End If

                    End If


                Case Is = 2    'มุมมอง

                   If dgvFix.Rows.Count <> 0 Then

                      btnReturn = CheckUserEntry(strDocCode, "act_view")
                      If btnReturn Then
                         lblCmd.Text = "2"
                         ViewShoeData()

                      Else
                          MsnAdmin()
                      End If
                   End If

                Case Is = 3   'กรองข้อมูล

                     If dgvFix.Rows.Count <> 0 Then

                        With gpbFilter

                             .Top = 230
                             .Left = 210
                             .Width = 348
                             .Height = 125

                             .Visible = True

                             cmbFilter.Text = cmbFilter.Items(0)
                             txtFilter.Text = _
                                         dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(2).Value.ToString.Trim

                             StateLockFind(False)
                             txtFilter.Focus()

                        End With

                     End If


                Case Is = 4   'ค้นหาข้อมูล

                     If dgvFix.Rows.Count <> 0 Then

                        With gpbSearch

                             .Top = 230
                             .Left = 210
                             .Width = 348
                             .Height = 125

                             .Visible = True

                             cmbType.Text = cmbType.Items(0)
                             txtSeek.Text = _
                             dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(2).Value.ToString.Trim

                             StateLockFind(False)
                             txtSeek.Focus()

                        End With

                     End If

               Case Is = 5   'ฟื้นฟูข้อมูล
                    blnHaveFilter = False
                    InputData()

               Case Is = 6  'พิมพ์ข้อมูล

                    If dgvFix.Rows.Count <> 0 Then

                        With gpbPrint

                            .Top = 160
                            .Left = 300
                            .Width = 450
                            .Height = 290

                            .Visible = True

                            rdo1.Select()
                            'InputDataTmpPrint()     'นำเข้าข้อมูลเพื่อพิมพ์
                            inputdataprint()        'ใส่ข้อมูลใน cbo รหัสอุปกรณ์

                            StateLockFind(False)

                            btnPrnCancle.Enabled = False
                            btnExit.Enabled = True

                        End With

                   End If

               Case Is = 7 'ลบข้อมูล

                    btnReturn = CheckUserEntry(strDocCode, "act_delete")

                    If btnReturn Then
                        DeleteData()
                        frmAeFixEqpmnt.UpdateFixsta("0")        'ลบสถานะส่งซ่อมให้เป็น สถานะปกติ

                    Else
                        MsnAdmin()
                    End If

                Case Is = 8   'รับเข้าส่งซ่อม

                  If CheckHaveData() Then

                     btnReturn = CheckUserEntry(strDocCode, "act_edit")
                     If btnReturn Then

                        IsShowSeek = Not IsShowSeek
                        lblCmd.Text = "3"

                        If IsShowSeek Then

                           With gpbFx             'groupbox รับเข้าอุปกรณ์

                               .Visible = True
                               .Left = 285
                               .Top = 200
                               .Height = 395
                               .Width = 868

                               dgvShow.Rows.Clear()
                               InputTmpdata()                'เป็น Tmptable สำหรับออกรายงาน
                               InputGpbRecv()                'นำเข้าข้อมูลการซ่อม
                               StateLockFindDept(False)

                          End With

                        Else
                              StateLockFindDept(True)

                        End If

                    Else
                         MsnAdmin()
                    End If

                End If

               Case Is = 9 'ออก
                    Me.Close()

            End Select

    End With

End Sub

    'Private Sub InputDataTmpPrint()

    ' Dim Conn As New ADODB.Connection
    ' Dim Rsd As New ADODB.Recordset

    ' Dim strSqlSelc As String = ""
    ' Dim strPart As String = ""

    '     With Conn

    '          If .State Then Close()
    '             .ConnectionString = strConnAdodb
    '             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '             .ConnectionTimeout = 90
    '             .Open()

    '              '------------------------------- ลบข้อมูลใน tmp_v_fixeqptrn2 ----------------------------

    '              strSqlSelc = "DELETE FROM tmp_v_fixeqptrn" _
    '                                  & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

    '              .Execute(strSqlSelc)

    '              '------------------------------- บันทึกข้อมูลงในตาราง tmp_v_fixeqptrn2 ----------------------------

    '              strSqlSelc = "INSERT INTO tmp_v_fixeqptrn" _
    '                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
    '                                  & " FROM v_fixeqptrn "

    '             .Execute(strSqlSelc)

    '     End With

    '  Conn.Close()
    '  Conn = Nothing

    'End Sub

    Private Sub InputTmpdata()

        Dim Conn As New ADODB.Connection
        Dim Rsd As New ADODB.Recordset

        Dim strSqlSelc As String = ""                          'เก็บสตริง sql select
        Dim strPart As String = ""

        With Conn

            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

            '------------------------------- บันทึกข้อมูลงในตาราง tmp_eqptrn ----------------------------

            strSqlSelc = "INSERT INTO tmp_fixeqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "', *" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE fix_sta = '1'" _
                                  & " OR fix_sta = '3'"

            .Execute(strSqlSelc)

        End With

        Conn.Close()
        Conn = Nothing

    End Sub

    Private Function CheckHaveData() As Boolean            'เช็คส่งซ่อมอุปกรณ์

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
                                 & " WHERE fix_sta = '1'" _
                                 & " OR fix_sta = '3'"

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

    Private Sub inputdataprint()                'ใส่ข้อมูลใน cbo ตัวเลือกการพิมพ์
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmdSelc As String

 Dim da As New System.Data.OleDb.OleDbDataAdapter
 Dim ds As New DataSet
 Dim strUser As String

      Me.WindowState = FormWindowState.Maximized
      strUser = frmMainPro.lblLogin.Text.Trim.ToString 'ใช้ User

     With Conn
          If .State Then Close()

             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .CommandTimeout = 30
             .Open()

     End With

          strSqlCmdSelc = "SELECT DISTINCT eqp_id FROM v_fixeqptrn "

          Rsd = New ADODB.Recordset

             With Rsd
                     .CursorType = CursorTypeEnum.adOpenKeyset
                     .LockType = LockTypeEnum.adLockOptimistic
                     .Open(strSqlCmdSelc, Conn, , )

                     If .RecordCount <> 0 Then

                        '------------------- ใส่รหัสอุปกรณ์ลงใน cboEqpid -------------------

                        ds.Clear()
                        da.Fill(ds, Rsd, "eqpid")
                        cboEqpid.DataSource = ds.Tables("eqpid").DefaultView
                        cboEqpid.DisplayMember = "eqp_id"
                        cboEqpid.ValueMember = "eqp_id"
                     End If

                   .ActiveConnection = Nothing
              ' .Close()
             End With

 btnPrnCancle.Enabled = False
 btnExit.Enabled = False

 Conn.Close()
 Conn = Nothing

End Sub

Private Sub ViewShoeData()
  If dgvFix.Rows.Count <> 0 Then
     ClearTmpTableUser("tmp_fixeqptrn")
     lblCmd.Text = "2"

     With frmAeFixEqpmnt
          .ShowDialog()
          .Text = "มุมมองข้อมูล"

     End With

  Else
     MsnAdmin()
  End If
End Sub

Private Sub InputEqpDataPrint()
  With gpbPrint
       .Visible = False
  End With

     Select Case staPrint           'เลือกรายงานก่อนพิมพ์

            Case Is = "1"             'รายงานส่งซ่อม
                frmMainPro.lblRptCentral.Text = "F"

            Case Is = "2"             'รายงานรับเข้า
                frmMainPro.lblRptCentral.Text = "G"

            Case Is = "3"             'รายงานส่งซ่อม - รับเข้า
                frmMainPro.lblRptCentral.Text = "H"

     End Select

     '-------------------------ส่งค่าให้ตัวแปร lblRptDesc ของฟอร์ม MainPro โดยส่ง Userid กับ Eqpid ----------------------------- 

     frmMainPro.lblRptDesc.Text = " user_id ='" & frmMainPro.lblLogin.Text.ToString.Trim & "'"
                                                               ' & "' AND eqp_id ='" & strDocId & "'"

     frmRptCentral.Show()

     StateLockFind(True)
     'gpbOptPrint.Visible = False
     frmMainPro.Hide()

End Sub

Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
  lblPage.Text = "1"
  InputData()
End Sub

Private Sub btnPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPre.Click
   If Int(lblPage.Text) > 1 Then
        lblPage.Text = Str(Int(lblPage.Text) - 1)
        InputData()
   End If
End Sub

Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
  lblPage.Text = Str(Int(lblPage.Text) + 1)
  InputData()
End Sub

Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
  lblPage.Text = Str(intBkPageCount)
  InputData()
End Sub

Private Sub txtPage_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPage.GotFocus
  txtPage.SelectAll()
End Sub

Private Sub txtPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPage.KeyPress
   If e.KeyChar = Chr(13) Then
      dgvFix.Focus()
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
                                    InputData()

                            Else

                                    lblPage.Text = intMovePage.ToString.Trim
                                    txtPage.Text = lblPage.Text
                                    InputData()

                            End If
                    Else

                        If intMovePage > 0 Then
                            lblPage.Text = intMovePage.ToString.Trim
                            txtPage.Text = lblPage.Text
                        Else
                            lblPage.Text = "1"
                            txtPage.Text = lblPage.Text
                        End If

                        InputData()

                    End If

                Catch ex As Exception
                    txtPage.Text = lblPage.Text
                End Try
End Sub

Private Sub dgvFix_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvFix.CellDoubleClick
   Dim blnReturn As Boolean
       blnReturn = CheckUserEntry(strDocCode, "act_view")

       If blnReturn Then
          ViewShoeData()
       End If
End Sub

Private Sub dgvFix_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvFix.KeyDown
  If e.KeyCode = Keys.Enter Then
     e.Handled = True
  End If
End Sub

Private Sub DeleteData()

 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

 Dim btyConsider As Byte
 Dim strFixID As String
 Dim strEqpID As String
 Dim strSizeID As String

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()


             If dgvFix.Rows.Count <> 0 Then

             strFixID = dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(16).Value.ToString.Trim
             strEqpID = dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(2).Value.ToString.Trim
             strSizeID = Mid(dgvFix.Rows(dgvFix.CurrentRow.Index).Cells(3).Value.ToString.Trim, 2)

             btyConsider = MsgBox("รหัสส่งซ่อม: " & strFixID & vbNewLine _
                                               & "รหัสอุปกรณ์ : " & strEqpID & vbNewLine _
                                               & "SIZE : " & strSizeID & vbNewLine _
                                               & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then


                       If chkFixData(strEqpID) Then   'ถ้าข้อมูลใน fixeqptrn ถ้า eqp_id มีเพียง 1 เรคคอร์ด

                          '--------------------------------- ลบตาราง fixeqptrn ------------------------------------------

                          strSqlCmd = "DELETE FROM fixeqptrn" _
                                              & " WHERE fix_id ='" & strFixID & "'" _
                                              & " AND size_id = '" & strSizeID & "'"

                          .Execute(strSqlCmd)
                           dgvFix.Rows.RemoveAt(dgvFix.CurrentRow.Index)

                          '------------------------------------ ลบตาราง fixeqpmst ----------------------------------------

                          strSqlCmd = "DELETE FROM fixeqpmst" _
                                                 & " WHERE fix_id ='" & strFixID & "'"

                          Conn.Execute(strSqlCmd)

                          '------------------------------------ Update fix_sta ใน eqpmst  ให้เป็น  0 = ปกติ  ----------------

                          strSqlCmd = "UPDATE eqpmst SET fix_sta = '" & "0" & "'" _
                                                 & " WHERE eqp_id  = '" & strEqpID & "'"

                          .Execute(strSqlCmd)
                          InputData()

                       Else

                          '--------------------------------- ลบตาราง fixeqptrn -------------------------------------------

                           strSqlCmd = "DELETE FROM fixeqptrn" _
                                                & " WHERE fix_id ='" & strFixID & "'" _
                                                & " AND size_id = '" & strSizeID & "'"

                           .Execute(strSqlCmd)

                          '------------------------------------ ลบตาราง fixeqpmst ----------------------------------------

                          'strSqlCmd = "DELETE FROM fixeqpmst" _
                          '                       & " WHERE fix_id ='" & strFixID & "'"


                          '.Execute(strSqlCmd)
                         ' dgvFix.Rows.RemoveAt(dgvFix.CurrentRow.Index)

                          InputData()

                       End If

                End If

         End If

     End With

  Conn.Close()
  Conn = Nothing

End Sub

Private Function chkFixData(ByVal txtEqpid As String) As Boolean        'เช็คข้อมูลใน fixeqptrn ว่าเหลือ เรคคอร์ดสุดท้ายหรือไม่

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
                                        & " WHERE eqp_id = '" & txtEqpid & "'"

              Rsd = New ADODB.Recordset

              With Rsd

                   .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                   .LockType = ADODB.LockTypeEnum.adLockOptimistic
                   .Open(strSqlSelc, Conn, , , )

                   If .RecordCount = 1 Then
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

Private Sub dgvFix_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dgvFix.KeyPress

  Dim blnReturn As Boolean

      If e.KeyChar = Chr(13) Then

         blnReturn = CheckUserEntry(strDocCode, "act_view")
         If blnReturn Then
            ViewShoeData()
         End If

      End If

End Sub

Private Sub dgvFix_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvFix.RowsAdded
   dgvFix.Rows(e.RowIndex).Height = 27
End Sub

Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
   SearchDT()
   strOperation = "0"
End Sub

Private Sub FindDocID()     'ค้นหาเอกสาร

  Dim strSearch As String = txtFilter.Text.ToUpper.Trim

     If strSearch <> "" Then
        FilterData()

     Else
         MsgBox("โปรดกรอกข้อมูลเพื่อค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
         txtSeek.Focus()
    End If

End Sub

Private Sub SearchDT()     'ค้นหาเอกสาร

 Dim strSearch As String = txtSeek.Text.ToUpper.Trim

     If strSearch <> "" Then

            Select Case cmbType.SelectedIndex()

                Case Is = 0 'รหัสส่งซ่อม
                    SearchData(0, strSearch)     'ส่งตำเงื่อนไข ,Text ให้ ซับรูทีน SearchData

                Case Is = 1 'รหัสอุปกรณ์
                    SearchData(1, strSearch)

                Case Is = 2 'สถานะส่งซ่อม
                    SearchData(2, strSearch)

            End Select

        Else
           MsgBox("โปรดกรอกข้อมูลเพื่อค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
          txtSeek.Focus()

    End If

End Sub

Private Sub txtSeek_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSeek.GotFocus
  txtSeek.SelectAll()
End Sub

Private Sub txtSeek_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSeek.KeyPress

        ' แปลงเป็นตัวพิมพ์ใหญ่ทันที
        If Char.IsLower(e.KeyChar) Then
            txtSeek.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If

        If e.KeyChar = Chr(13) And txtSeek.Text.Length > 0 Then
            SearchDT()
            strOperation = "0"
        End If

    End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
  StateLockFind(True)
  gpbSearch.Visible = False
End Sub

Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
   FindDocID()
   strOperation = "1"  'บ่งบอกว่ากรองข้อมูล
End Sub

Private Sub txtFilter_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFilter.KeyPress

        ' แปลงเป็นตัวพิมพ์ใหญ่ทันที
        If Char.IsLower(e.KeyChar) Then
            txtFilter.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If

        If e.KeyChar = Chr(13) And txtFilter.Text.Length > 0 Then
            FindDocID()
        End If
    End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click
  StateLockFind(True)
  gpbFilter.Visible = False
End Sub

Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
  InputData()
End Sub

Private Sub lblGpbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblGpbClose.Click
  With gpbFx
       .Visible = False
        IsShowSeek = False
  End With
  StateLockFindDept(True)
End Sub

Private Sub InputGpbRecv()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmdSelc As String = ""

 Dim imgStaFix As Image
 Dim strEqpSta As String = ""

       With Conn

            If .State Then .Close()

               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

              strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                        & " WHERE fix_sta = '1'" _
                                        & " OR fix_sta = '3'" _
                                        & " ORDER BY fix_id"

              With Rsd

                   .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                   .LockType = ADODB.LockTypeEnum.adLockOptimistic
                   .Open(strSqlCmdSelc, Conn, , , )

                   If .RecordCount <> 0 Then

                       dgvShow.Rows.Clear()      'เคลียร์ datagrid
                       Do While Not .EOF

                                    '--------------------------- สถานะส่งซ่อม ------------------------------

                                            Select Case .Fields("fix_sta").Value.ToString.Trim

                                                   Case Is = "1"     'ส่งซ่อม
                                                        imgStaFix = My.Resources._16x16_ledred
                                                        strEqpSta = "ส่งซ่อม"

                                                   Case Is = "2"     'รับคืนส่งซ่อม
                                                        imgStaFix = My.Resources._16x16_ledgreen
                                                        strEqpSta = "รับคืนส่งซ่อม"

                                                   Case Is = "3"     'รับคืนบางส่วน
                                                        imgStaFix = My.Resources._16x16ledyellow
                                                        strEqpSta = "ค้างส่งซ่อม"

                                                   Case Else         'ปกติ
                                                        imgStaFix = My.Resources.blank

                                            End Select

                                            dgvShow.Rows.Add( _
                                                                 imgStaFix, _
                                                                 strEqpSta, _
                                                                 .Fields("fix_id").Value.ToString.Trim, _
                                                                 .Fields("eqp_id").Value.ToString.Trim, _
                                                                 .Fields("size_id").Value.ToString.Trim, _
                                                                 .Fields("eqp_name").Value.ToString.Trim, _
                                                                 "เลือก" _
                                                              )

                                         .MoveNext()            'ข้ามไปที่ระเบียนใหม่
                                     Loop

                            Else
                                MsgBox("Not found data")
                            End If

                   .Close()

              End With

              Rsd = Nothing

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub dgvShow_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShow.CellClick

 Dim btnReturn As Boolean

       With dgvShow

            Select Case .CurrentCell.ColumnIndex

                   Case Is = 6

                         btnReturn = CheckUserEntry(strDocCode, "act_edit")

                         If btnReturn Then

                            ClearTmpTableUser("tmp_fixeqptrn")
                            lblCmd.Text = "3"

                                 With frmAeFixEqpmnt
                                      .ShowDialog()
                                      .Text = "รับเข้าส่งซ่อม"
                                 End With

                            IsShowSeek = False
                            StateLockFindDept(True)
                            gpbFx.Visible = False

                         Else
                             MsnAdmin()
                         End If

            End Select

       End With

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
   With gpbPrint
        .Visible = False

        StateLockBtnprint(True)
        StateLockFind(True)
   End With
End Sub

Private Sub StateLockBtnprint(ByVal sta As Boolean)   'groupbox PRINT
   btnPrnOk.Enabled = sta
   btnExit.Enabled = sta
End Sub

Private Sub ChkAllEqp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAllEqp.CheckedChanged
  If ChkAllEqp.Checked Then
     cboEqpid.Enabled = False

  Else
     cboEqpid.Enabled = True

  End If

End Sub

Private Sub btnPrnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrnOk.Click
 Dim Conn As New ADODB.Connection

 Dim Rsd1 As New ADODB.Recordset
 Dim Rsd2 As New ADODB.Recordset
 Dim strSqlCmdSelc As String = ""
 Dim strSqlSelc As String = ""

 Dim da As New System.Data.OleDb.OleDbDataAdapter
 Dim ds As New DataSet

     With Conn

            If .State Then .Close()

                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .CommandTimeout = 30
                .Open()
     End With

                    '------------------------- คิวรี่ข้อมูล(กรณีส่งซ่อม,เลือกทุกอุปกรณ์ )------------------------

                    If rdo1.Checked = True And ChkAllEqp.Checked = True Then
                       frmMainPro.lblRptCentral.Text = "F"

                       strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                                             & " WHERE fix_sta = '1'" _
                                                             & " ORDER BY fix_id"

                    '------------------------- คิวรี่ข้อมูล(กรณีส่งซ่อม,ระบุอุปกรณ์)----------------------------

                         ElseIf rdo1.Checked = True And ChkAllEqp.Checked = False Then
                                frmMainPro.lblRptCentral.Text = "F"

                                strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                                & " WHERE eqp_id = '" & cboEqpid.Text.ToString & "'" _
                                                & " AND fix_sta = '1'" _
                                                & " ORDER BY fix_id"

                    '------------------------- คิวรี่ข้อมูล(กรณีรับกลับส่งซ่อม,ทุกรหัสอุปกรณ์)-----------------

                         ElseIf rdo2.Checked = True And ChkAllEqp.Checked = True Then
                                frmMainPro.lblRptCentral.Text = "G"

                                strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                                               & " WHERE fix_sta = '2'" _
                                                               & " OR fix_sta = '3'"

                    '------------------------- คิวรี่ข้อมูล(กรณีรับกลับส่งซ่อม ,ระบุอุปกรณ์)-----------------------------

                         ElseIf rdo2.Checked = True And ChkAllEqp.Checked = False Then
                                frmMainPro.lblRptCentral.Text = "G"

                                strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)" _
                                                        & " WHERE eqp_id = '" & cboEqpid.Text.ToString & "'" _
                                                        & " AND fix_sta = '2'" _
                                                        & " Or fix_sta = '3'"

                    '------------------------- คิวรี่ข้อมูล(กรณีทั้ง 2 กรณี)ok -----------------------------

                         ElseIf rdo3.Checked = True Then
                                frmMainPro.lblRptCentral.Text = "H"

                                gpbRpt2.Enabled = False
                                strSqlCmdSelc = "SELECT * FROM v_fixeqptrn (NOLOCK)"

                    End If

       Rsd1 = New ADODB.Recordset

       With Rsd1
                .CursorType = CursorTypeEnum.adOpenKeyset
                .LockType = LockTypeEnum.adLockOptimistic
                .Open(strSqlCmdSelc, Conn, , )


                If .RecordCount <> 0 Then

                   For i As Integer = 1 To .RecordCount

                    strSqlSelc = "SELECT * " _
                                               & " FROM tmp_v_fixeqptrn (NOLOCK)"

                    Rsd2 = New ADODB.Recordset
                    Rsd2.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                    Rsd2.LockType = ADODB.LockTypeEnum.adLockOptimistic
                    Rsd2.Open(strSqlSelc, Conn, , , )

                              Rsd2.AddNew()
                              Rsd2.Fields("group").Value = .Fields("group").Value
                              Rsd2.Fields("fix_sta").Value = .Fields("fix_sta").Value
                              Rsd2.Fields("fix_desc").Value = .Fields("fix_desc").Value
                              Rsd2.Fields("fix_id").Value = .Fields("fix_id").Value
                              Rsd2.Fields("eqp_id").Value = .Fields("eqp_id").Value
                              Rsd2.Fields("eqp_name").Value = .Fields("eqp_name").Value
                              Rsd2.Fields("size_id").Value = .Fields("size_id").Value
                              Rsd2.Fields("amt_out").Value = .Fields("amt_out").Value
                              Rsd2.Fields("amt_in").Value = .Fields("amt_out").Value
                              Rsd2.Fields("price").Value = .Fields("price").Value
                              Rsd2.Fields("fix_date").Value = .Fields("fix_date").Value
                              Rsd2.Fields("fix_by").Value = .Fields("fix_by").Value
                              Rsd2.Fields("pr_doc").Value = .Fields("pr_doc").Value
                              Rsd2.Fields("issue").Value = .Fields("issue").Value
                              Rsd2.Fields("desc_thai").Value = .Fields("desc_thai").Value
                              Rsd2.Fields("fix_issue").Value = .Fields("fix_issue").Value
                              Rsd2.Fields("sup_name").Value = .Fields("sup_name").Value
                              Rsd2.Fields("due_date").Value = .Fields("due_date").Value
                              Rsd2.Fields("recv_date").Value = .Fields("recv_date").Value
                              Rsd2.Fields("recv_by").Value = .Fields("recv_by").Value
                              Rsd2.Fields("fix_rmk").Value = .Fields("fix_rmk").Value
                              Rsd2.Fields("user_id").Value = frmMainPro.lblLogin.Text.ToString.Trim

                    Rsd2.Update()
                    Rsd2.ActiveConnection = Nothing
                    Rsd2.Close()
                    Rsd2 = Nothing
                   .MoveNext()     'เลื่อนไปที่ Record ถัดไป

                   Next i

                      InputEqpDataPrint()
                Else
                     MsgBox("ไมมีข้อมูล กรุณาเลือกรายการอื่น")
                     StateLockBtnprint(True)
                End If

       .ActiveConnection = Nothing
       .Close()

       End With

    Rsd1 = Nothing

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub dgvShow_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvShow.RowsAdded
 dgvShow.Rows(e.RowIndex).Height = 30
End Sub

Private Sub rdo1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo1.CheckedChanged
 gpbRpt2.Enabled = True
End Sub

Private Sub rdo3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo3.CheckedChanged
 gpbRpt2.Enabled = False
End Sub

Private Sub rdo2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo2.CheckedChanged
 gpbRpt2.Enabled = True
End Sub

End Class
