Imports ADODB
Imports System.IO
Imports System.Drawing.Drawing2D
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.IO.MemoryStream

Public Class frmCutting

Dim intBkPageCount As Integer   'ตัวเเปรนับจำนวน page ทั้งหมด
Dim blnHaveFilter As Boolean    'ตัวเเปรเก็บค่า กรณีกรองข้อมูล

Dim dubNumberStart As Double   'ถูกกำหนด = 1
Dim dubNumberEnd As Double     'ถูกกำหนด = 2100

Dim strSqlFindData As String
Dim strDocCode As String = "F2"

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet


Private Sub frmCutting_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

  Dim strSearch As String

    If FormCount("frmAeCutting") > 0 Then       'เช็คว่ามีฟอร์มเปิดอยู่หรือไม่

        With frmAeCutting

                strSearch = .lblComplete.Text           'ให้ strSearch รับค่า Label ที่ส่งมา

                If strSearch <> "" Then               'ตรวจสอบข้อมูล

                   SearchData(0, strSearch)           'เรียก SearchData

                End If

              .Close()   'ถ้าฟอร์มถูกเปิดอยู่ก่อนแล้วให้ close

        End With

    Timer1.Enabled = True  'สั่งรีเฟรชข้อมูลทุก 1 นาที    
    End If

    Me.Height = Int(lblHeight.Text)    'ความสูงฟอร์ม = lblHeight.text
    Me.Width = Int(lblWidth.Text)      'ความกว้างฟอร์ม = lblWidth.text

    Me.Top = Int(lblTop.Text)          'ขอบบน = lblTop.text
    Me.Left = Int(lblLeft.Text)        'ขอบล่าง = lblLeft.text
End Sub

Private Function FormCount(ByVal frmName As String) As Long

  Dim frm As Form

    For Each frm In My.Application.OpenForms

            If frm Is My.Forms.frmAeCutting Then
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

Private Sub frmCutting_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   Me.Dispose()
End Sub


Private Sub frmCutting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

   Me.WindowState = FormWindowState.Maximized
   StdDateTimeThai()   'ตั้งค่าวันที่เป็นภาษาไทย
   tlsBarFmr.Cursor = Cursors.Hand

   dubNumberStart = 1
   dubNumberEnd = 2100

   PreGroupType()
   InputData()
   tabCmd.Focus()

End Sub

Private Sub frmCutting_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
  lblHeight.Text = Me.Height.ToString.Trim
  lblWidth.Text = Me.Width.ToString.Trim
  lblTop.Text = Me.Top.ToString.Trim
  lblLeft.Text = Me.Left.ToString.Trim
End Sub

Private Sub PreGroupType()
  Dim strGpTopic(6) As String
  Dim i As Byte

      strGpTopic(0) = "รหัสอุปกรณ์"
      strGpTopic(1) = "รายละเอียดอุปกณ์"
      strGpTopic(2) = "ชิ้นส่วนที่ผลิต"
      strGpTopic(3) = "กลุ่มอุปกรณ์"
      strGpTopic(4) = "รายละเอียดกลุ่ม"
      strGpTopic(5) = "สถานะส่งมอบ"
      strGpTopic(6) = "สถานะส่งซ่อม"

  With cmbType
       For i = 0 To 6
           .Items.Add(strGpTopic(i))
       Next i
  End With

  With cmbFilter
       For i = 0 To 6
           .Items.Add(strGpTopic(i))
       Next i
  End With

End Sub

Private Sub InputData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset

  Dim strSqlCmdSelc As String = ""
  Dim DataDate As Date = Now()

  Dim intPageCount As Integer   'ตัวแปรนับจำนวน Record ใน DB
  Dim intPageSize As Integer    'จำนวน Record ที่แสดงใน Grid = 30
  Dim intCounter As Integer

  Dim strSearch As String = txtFilter.Text.ToString.Trim    'เก็บค่าจาก Text Search
  Dim strFieldFilter As String = ""

  Dim dteCom As Date = Now    'เก็บวันที่ ณ ปัจจุบัน
  Dim imgStaPrd As Image
  Dim imgStaFix As Image

         With Conn
              If .State Then Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()
         End With

         If blnHaveFilter Then   'กรณีกรองข้อมูล

               Select Case cmbFilter.SelectedIndex

                   Case Is = 0  'กรณีกรองจากรหัสอุปกรณ์
                        strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                   Case Is = 1  'กรณีกรองจากชื่อ
                        strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                   Case Is = 2
                        strFieldFilter = "part_nw like '%" & ReplaceQuote(strSearch) & "%'"

                   Case Is = 3
                        strFieldFilter = "desc_eng like '%" & ReplaceQuote(strSearch) & "%'"

                   Case Is = 4
                        strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                   Case Is = 5
                        strFieldFilter = "sta_pd like '%" & ReplaceQuote(strSearch) & "%'"

                   Case Is = 6
                        strFieldFilter = "sta_fx like '%" & ReplaceQuote(strSearch) & "%'"

               End Select

                   strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                       & " WHERE " & strFieldFilter _
                                       & " AND [group] ='E'" _
                                       & " ORDER BY eqp_id"

         Else

                strSqlCmdSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                                  & " WHERE RowNumber >=" & dubNumberStart.ToString.Trim _
                                                  & " AND RowNumber <=" & dubNumberEnd.ToString.Trim _
                                                  & " AND [group] ='E'" _
                                                  & " ORDER BY eqp_id"


         End If
         intPageSize = 30  ' กำหนดให้ 1 หน้าแสดง 30 รายการ
         Rsd = New ADODB.Recordset

         With Rsd

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , , )

               If .RecordCount <> 0 Then

                                    If intPageSize > .RecordCount Then   'ถ้าจำนวนRecordทีจะแสดง > จำนวน Record ใน DB
                                       intPageSize = .RecordCount
                                    Else : intPageSize = 0
                                       intPageSize = 30
                                    End If

                                    .PageSize = intPageSize             '.PageSize ใช้กำหนดว่าแต่ละหน้าจะให้มีกี่รายการ ในการแสดงผล
                                     intPageCount = .PageCount          '.PageCount นับจำนวนหน้าทั้งหมด ที่ได้จากการกำหนดขนาดของหน้า


                                  '--------------------------ถ้ามีการค้นหา----------------------------------------

                                    If strSqlFindData <> "" Then      'strSqlFindData ส่งมาจาก ซับรูทีน Searchdata

                                            .MoveFirst()          'เป็นออบเจ็กต์สำหรับ การเลื่อน Record ไปยัง Record แรกสุด
                                            .Find(strSqlFindData)

                                             If Not .EOF Then
                                                lblPage.Text = Str(.AbsolutePage)        '.AbsolutePage ใช้อ้างอิงไปยังหน้าที่ต้องการ
                                             End If

                                            strSqlFindData = ""

                                    End If

                                   '------------------------------------------------------------------------------------

                                    If Int(lblPage.Text.ToString) > intPageCount Then
                                       lblPage.Text = intPageCount.ToString
                                    End If

                                    txtPage.Text = lblPage.Text.ToString
                                    intBkPageCount = .PageCount       '.PageCount นับจำนวนหน้าทั้งหมด ที่ได้จากการกำหนดขนาดของหน้า
                                    lblPageAll.Text = "/ " & .PageCount.ToString
                                    .AbsolutePage = Int(lblPage.Text.ToString)

                                    dgvShoe.Rows.Clear()

                                    intCounter = 0


                                    Do While Not .EOF      '.EOF เป็นออบเจ็กต์ตรวจสอบ Pointer ในตำแหน่งสุดท้าย, .BOF เป็นออบเจ็กต์ตรวจสอบ Pointer ในตำแหน่งเริ๋มต้น
                                   '-------------------------------------------สถานะส่งมอบฝ่ายผลิต----------------------------------------------------------------

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
                              '-------------------------------------------สถานะส่งซ่อม----------------------------------------------------------------

                                            Select Case .Fields("fix_sta").Value.ToString.Trim

                                                          Case Is = "1"
                                                                    imgStaFix = My.Resources.sign_deny
                                                         Case Is = "2"
                                                                    imgStaFix = My.Resources.Chk
                                                          Case Else
                                                                    imgStaFix = My.Resources.blank

                                            End Select

                                            dgvShoe.Rows.Add( _
                                                                          .Fields("eqp_id").Value.ToString.Trim, _
                                                                          .Fields("exp_id").Value.ToString.Trim, _
                                                                          .Fields("eqp_name").Value.ToString.Trim, _
                                                                          .Fields("part_nw").Value.ToString.Trim, _
                                                                          .Fields("desc_eng").Value.ToString.Trim, _
                                                                          .Fields("desc_thai").Value.ToString.Trim, _
                                                                           imgStaPrd, .Fields("sta_pd").Value.ToString.Trim, _
                                                                          .Fields("eqptype").Value.ToString.Trim, _
                                                                           Mid(.Fields("pre_date").Value.ToString.Trim, 1, 10), _
                                                                           .Fields("pre_by").Value.ToString.Trim, _
                                                                           Mid(.Fields("last_date").Value.ToString.Trim, 1, 10), _
                                                                          .Fields("last_by").Value.ToString.Trim _
                                                                          )
                                            intCounter = intCounter + 1

                                              If intCounter = intPageSize Then
                                                    Exit Do
                                              End If

                                         .MoveNext()
                                     Loop

                Else
                     intBkPageCount = 1
                     txtPage.Text = "1"

                End If
                .Close()
                Rsd = Nothing
              End With

Conn.Close()
Conn = Nothing
End Sub

    Private Sub tabCmd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCmd.Click

        Dim btnReturn As Boolean

        With tabCmd

            Select Case .SelectedIndex

                Case Is = 0 'เพิ่มข้อมูล

                    btnReturn = CheckUserEntry(strDocCode, "act_add")   'ฟังก์ชั่นตรวจสอบสิทธิ์ strDocCode = 'F0', act_add คือสิทธิ์การ เพิ่มข้อมูล
                    If btnReturn Then

                        ClearTmpTableUser("tmp_eqptrn")     'สั่ง Clear ค่าใน tmp_eqptrn
                        lblCmd.Text = "0"

                        With frmAeCutting
                            .ShowDialog()
                            .Text = "เพิ่มข้อมูล"

                        End With

                    Else
                        MsnAdmin()     'message คุณไม่มีสิทธิใช้ข้อมูลส่วนนี้
                    End If

                Case Is = 1  'แก้ไขข้อมูล

                    If dgvShoe.Rows.Count <> 0 Then

                        btnReturn = CheckUserEntry(strDocCode, "act_edit")
                        If btnReturn Then

                            ClearTmpTableUser("tmp_eqptrn")
                            lblCmd.Text = "1"

                            With frmAeCutting
                                .ShowDialog()
                                .Text = "แก้ไขข้อมูล"
                            End With

                        Else
                            MsnAdmin()
                        End If

                    End If

                Case Is = 2  'มุมมอง

                    If dgvShoe.Rows.Count <> 0 Then

                        btnReturn = CheckUserEntry(strDocCode, "act_view")
                        If btnReturn Then
                            ViewShoeData()
                        Else
                            MsnAdmin()
                        End If

                    End If

                Case Is = 3   'กรองข้อมูล

                    If dgvShoe.RowCount > 0 Then

                        With gpbFilter

                            .Top = 230
                            .Left = 210
                            .Width = 348
                            .Height = 125

                            .Visible = True

                            cmbFilter.SelectedItem = cmbFilter.Items(0)
                            txtFilter.Text =
                                      dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim

                            StateLockFind(False)
                            txtFilter.Focus()

                        End With

                    End If

                Case Is = 4    'ค้นหาข้อมูล

                    If dgvShoe.RowCount > 0 Then
                        With gpbSearch

                            .Top = 230
                            .Left = 210
                            .Width = 348
                            .Height = 125

                            .BringToFront()
                            .Visible = True

                            cmbType.SelectedItem = cmbType.Items(0)
                            txtSeek.Text =
                                      dgvShoe.Rows(dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim

                            StateLockFind(False)
                            txtSeek.Focus()

                        End With
                    End If

                      'Case Is = 5  'พิมพ์เอกสาร

                Case Is = 5  'ฟื้นฟูข้อมูล
                    blnHaveFilter = False
                    InputData()

                Case Is = 6 'ลบข้อมูล
                    btnReturn = CheckUserEntry(strDocCode, "act_delete")
                    If btnReturn Then
                        DeleteData()
                    Else
                        MsnAdmin()
                    End If

                Case Is = 7 'ออก
                    Me.Close()

            End Select
        End With
    End Sub

    Private Sub DeleteData()
  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

  Dim strEqpid As String
  Dim strDepExp As String
  Dim strDetail As String

  Dim btyConsider As Byte

      With Conn
           If .State Then Close()
           .ConnectionString = strConnAdodb
           .CursorLocation = ADODB.CursorLocationEnum.adUseClient
           .ConnectionTimeout = 90
           .Open()
      End With

      With dgvShoe

           If .RowCount > 0 Then
                strEqpid = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim
                strDepExp = .Rows(.CurrentRow.Index).Cells(1).Value.ToString.Trim
                strDetail = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim

                btyConsider = MsgBox("รหัสอุปกรณ์ : " & strEqpid & vbNewLine _
                                                     & "รายละเอียดอุปกรณ์ : " & strDetail & vbNewLine _
                                                     & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                      + MsgBoxStyle.Critical, "Confirm Delete Data")

                If btyConsider = 6 Then

                    Conn.BeginTrans()
                    '---------------------------- ลบข้อมูลในตาราง eqpmst --------------------------------------------------

                    strSqlCmd = "DELETE FROM eqpmst " _
                                       & " WHERE eqp_id = '" & strEqpid & "'"

                    Conn.Execute(strSqlCmd)

                    '---------------------------- ลบข้อมูลในตาราง eqp_trn -------------------------------------------------
                    strSqlCmd = "DELETE FROM eqptrn" _
                                        & " WHERE  eqp_id = '" & strEqpid & "'"


                    Conn.Execute(strSqlCmd)
                    Conn.CommitTrans()

                  .Rows.RemoveAt(.CurrentRow.Index)
                  InputData()
               End If

         End If
         .Focus()
      End With

    Conn.Close()
    Conn = Nothing
End Sub

Private Sub StateLockFind(ByVal sta As Boolean)
  tabCmd.Enabled = sta
  dgvShoe.Enabled = sta
  tlsBarFmr.Enabled = sta
End Sub

Private Sub ViewShoeData()
  If dgvShoe.RowCount > 0 Then

     ClearTmpTableUser("tmp_eqptrn")
     lblCmd.Text = "2"

     With frmAeCutting
          .ShowDialog()
          .Text = "มุมมองข้อมูล"
     End With

     'Me.Hide()
     'frmMainPro.Hide()
  End If

End Sub

'--------------------------------------- เมื่อคลิกปุ่ม First ---------------------------------------------
Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
  lblPage.Text = "1"
  InputData()
End Sub

'--------------------------------------- เมื่อคลิกปุ่ม Previus -------------------------------------------
Private Sub btnPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPre.Click
  If Int(lblPage.Text) > 1 Then
     lblPage.Text = Str(Int(lblPage.Text) - 1)
     InputData()
  End If
End Sub

'---------------------------------------- เมื่อคลิกปุ่ม Next ---------------------------------------------
Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
  lblPage.Text = Str(Int(lblPage.Text) + 1)
  InputData()
End Sub

Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
  lblPage.Text = Str(intBkPageCount)
  InputData()
End Sub

Private Sub txtPage_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPage.KeyPress
  If e.KeyChar = Chr(13) Then
     dgvShoe.Focus()
  End If
End Sub

Private Sub txtPage_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPage.LostFocus
  ChangePage()
End Sub

Private Sub ChangePage()
  Dim i, x As Integer
  Dim strTmp As String = ""
  Dim strMerg As String = ""
  Dim IntMovePage As Integer     'เก็บเลขหน้าที่จะไป

      x = Len(txtPage.Text.ToString.Trim)
        For i = 1 To x
               strTmp = Mid(txtPage.Text.ToString.Trim, i, 1)

               Select Case strTmp
                      Case Is = ","
                      Case Is = "+"
                      Case Is = "-"
                      Case Is = "_"
                      Case Else
                           strMerg = strTmp = strTmp & Trim(strTmp)
               End Select
               strTmp = ""
        Next i

  Try

     IntMovePage = Int(strMerg)     'IntMovePage = จำนวนเลขหน้าใน TextBox
     If IntMovePage >= Int(lblPage.Text) Then
            If IntMovePage <= intBkPageCount Then

               lblPage.Text = IntMovePage.ToString.Trim
               txtPage.Text = lblPage.Text
               InputData()

            Else

               lblPage.Text = IntMovePage.ToString.Trim
               txtPage.Text = lblPage.Text
               InputData()

            End If

     Else

        If IntMovePage > 0 Then
           lblPage.Text = IntMovePage.ToString.Trim
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

Private Sub SearchData(ByVal bytColNumber As Byte, ByVal strSearchText As String)     'รับค่า Index,สตริง Search

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlFind As String = ""
 Dim strSqlSelc As String

 Dim intPageSize As Integer
 Dim intPageCount As Integer
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
                 strSqlFind = "eqp_id"
                 strSqlFind = strSqlFind & " LIKE '%" & ReplaceQuote(strSearchText) & "%'"

                 '--------------------- เว้นวรรณ ตรง LIKE เพื่อไม่ให้คำสั่ง  Command ติดกับ eqp_id ----------------------

            Case Is = 2
                 strSqlFind = "eqp_name"
                 strSqlFind = strSqlFind & " LIKE '%" & ReplaceQuote(strSearchText) & "%'"

            Case Is = 3
                 strSqlFind = "part_nw "
                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchText) & "%'"

            Case Is = 4
                 strSqlFind = "desc_eng "
                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchText) & "%'"

            Case Is = 5
                 strSqlFind = "desc_thai "
                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchText) & "%'"

            Case Is = 7
                 strSqlFind = "sta_pd "
                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchText) & "%'"

            Case Is = 9
                 strSqlFind = "sta_fx "
                 strSqlFind = strSqlFind & " Like '%" & ReplaceQuote(strSearchText) & "%'"

     End Select

        strSqlSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                             & " WHERE " & strSqlFind _
                             & " AND [group]= 'E'" _
                             & " ORDER BY eqp_id"


     intPageSize = 30  'จำนวน Record ต่อหน้า

     Rsd = New ADODB.Recordset
     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .CursorLocation = ADODB.CursorLocationEnum.adUseClient
          .Open(strSqlSelc, Conn, , , )

               If .RecordCount <> 0 Then

                      ' -------------------------------- โค้ดเซ็ตจำนวน Record ที่แสดงในฟอร์ม ------------------------------------------
                      If intPageSize > .RecordCount Then
                         intPageSize = .RecordCount

                      End If

                      If intPageSize = 0 Then
                         intPageSize = 30

                      End If

                      .PageSize = intPageSize         '.PageSize ใช้กำหนดว่าแต่ละหน้าจะให้มีกี่รายการ ในการแสดงผล
                      intPageCount = .PageCount       '.PageCount นับจำนวนหน้าทั้งหมด ที่ได้จากการกำหนดขนาดของหน้า

                    ' ---------------------------------------ค้นหาข้อมูล-------------------------------------------------------------
                      .MoveFirst()
                      .Find(strSqlFind)
                      lblPage.Text = Str(.AbsolutePage)

                    '-------------------------------------------------------------------------------------------------------------

                     If .Fields("RowNumber").Value > 2100 Then

                        'IIF()อ่านว่า If and only If
                        'จะมีการ return ค่าหลังจากตรวจสอบเงื่อนไข ซึ่งเขียนได้ในบรรทัดเดียวด้วย รูปแบบเป็นดังนี้ 
                        'รูปแบบเป็นดังนี้ IIF(เงื่อนไข, ค่าเมื่อเงื่อนไขเป็นจริง, ค่าเมื่อเงื่อนไขเป็นเท็จ) 

                         dubNumberStart = IIf(.Fields("RowNumber").Value <= 500, .Fields("RowNumber").Value, .Fields("RowNumber").Value - 500)
                         dubNumberEnd = .Fields("RowNumber").Value + 1000

                     Else
                         dubNumberStart = 1
                         dubNumberEnd = 2100
                     End If

                         strSqlFindData = strSqlFind

                         InputData()


                                                For i = 0 To dgvShoe.Rows.Count - 1

                                                        'ฟังก์ช่ัน InStr()คืนค้า int=ใช้ในการค้นหาว่า สตริงนี้ปรากฏอยู่ในสตริงที่ค้นหาอยู่หรือไม่โดยมีพารามิเตอร์ (ข้อความที่ถูกค้น,ข้อความที่จะนำไปค้น)
                                                        '---- UCase ใช้เเปลง String เป็นตัวอักษรพิมพ์ใหญ่ -------------------

                                                        If InStr(UCase(dgvShoe.Rows(i).Cells(bytColNumber).Value), strSearchText.Trim.ToUpper) <> 0 Then
                                                                dgvShoe.CurrentCell = dgvShoe.Item(bytColNumber, i)
                                                                dgvShoe.Focus()
                                                                Exit For
                                                        End If

                                                 Next i

               Else

                     MsgBox("ไม่มีข้อมูล : " & strSearchText & " ในระบบ" & vbNewLine _
                                         & "โปรดระบุการค้นหาข้อมูลใหม่!", vbExclamation, "Not Found Data")

               End If
               .ActiveConnection = Nothing      'Rsd
               .Close()                         'Rsd

     End With
     Rsd = Nothing

  Conn.Close()
  Conn = Nothing

StateLockFind(True)
gpbSearch.Visible = False
End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click

        If txtSeek.Text.Length > 0 Then
            FindDocID()
        End If

    End Sub

    Private Sub FindDocID()
   Dim strSearchTxt As String = txtSeek.Text.ToUpper.Trim
       If strSearchTxt <> "" Then
          Select Case cmbType.SelectedIndex
                 Case Is = 0 'รหัสอุปกรณ์
                         SearchData(0, strSearchTxt)   'เรียกฟังก์ชั่น SearchData(Index,สตริง Search)

                 Case Is = 1 'ชื่ออุปกรณ์
                         SearchData(1, strSearchTxt)

                 Case Is = 2
                         SearchData(2, strSearchTxt)

                 Case Is = 3
                         SearchData(3, strSearchTxt)

                 Case Is = 4
                         SearchData(4, strSearchTxt)

                 Case Is = 5 'สถานะส่งฝ่ายผลิต
                         SearchData(7, strSearchTxt)

                 Case Is = 6 'สถานะส่งซ่อม
                          SearchData(9, strSearchTxt)

          End Select
       Else
             MsgBox("ไม่มีข้อมูลที่ต้องการค้นหา!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
             txtSeek.Focus()


       End If

End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
   StateLockFind(True) 'ล็อค Tabcmd, DgvSize, tlsBarFmr
   gpbSearch.Visible = False
End Sub

Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
   FilterData()
End Sub

Private Sub FilterData()
   Dim Conn As New ADODB.Connection
   Dim Rsd As New ADODB.Recordset
   Dim strSqlSelc As String = ""
   Dim blnHaveData As Boolean

   Dim strFieldFilter As String = ""
   Dim strSearch As String = txtFilter.Text.ToUpper.Trim

   If strSearch <> "" Then
          With Conn
                  If .State Then Close()
                     .ConnectionString = strConnAdodb
                     .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                     .ConnectionTimeout = 90
                     .Open()
          End With

                    Select Case cmbFilter.SelectedIndex

                           Case Is = 0
                                     strFieldFilter = "eqp_id like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 1
                                     strFieldFilter = "eqp_name like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 2
                                     strFieldFilter = "part_nw like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 3
                                     strFieldFilter = "desc_eng like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 4
                                     strFieldFilter = "desc_thai like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 5
                                     strFieldFilter = "sta_pd like '%" & ReplaceQuote(strSearch) & "%'"

                           Case Is = 6
                                     strFieldFilter = "sta_fx like '%" & ReplaceQuote(strSearch) & "%'"
                    End Select

                        strSqlSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                        & "WHERE " & strFieldFilter _
                                        & "AND [group]='E'" _
                                        & "ORDER BY eqp_id"

                       Rsd = New ADODB.Recordset
                       With Rsd
                            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                            .Open(strSqlSelc, Conn, , , )

                       If .RecordCount <> 0 Then
                          blnHaveData = True
                       Else
                          blnHaveData = False
                       End If

                      End With
                      Rsd = Nothing

            Conn.Close()
            Conn = Nothing

            If blnHaveData = True Then
               blnHaveFilter = True        'ประกาศไว้ในตัวเเปรคลาส --> ส่งค่าไปให้ InputData
               InputData()

               StateLockFind(True)
               gpbFilter.Visible = False

            Else
               blnHaveFilter = False
               MsgBox("ไม่มีข้อมูลที่ต้องการกรองข้อมูล!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
               txtFilter.Focus()

            End If
   Else
        MsgBox("โปรดระบุข้อมูลที่ต้องการกรอง!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please input DocID!!")
        txtFilter.Focus()
   End If

End Sub

'------------------------------- ฟังก์ชั่นปรับขนาด size --------------------------------------------------------
Private Function SizeImage(ByVal img As Bitmap, ByVal width As Integer, ByVal height As Integer) As Bitmap

        Dim newBit As New Bitmap(width, height) 'new blank bitmap
        Dim g As Graphics = Graphics.FromImage(newBit)
        'change interpolation for reduction quality
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(img, 0, 0, width, height)
        Return newBit

End Function

'----------------- event DoubleClick dgv------------------------------------------------
Private Sub dgvShoe_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvShoe.CellDoubleClick
  Dim blnReturn As Boolean

      blnReturn = CheckUserEntry(strDocCode, "act_view")
     If blnReturn Then
       ViewShoeData()
     End If
End Sub

Private Sub dgvShoe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvShoe.KeyDown
  If e.KeyCode = Keys.Enter Then
    e.Handled = True
End If
End Sub

'---------------------------------- ตั้งค่า Row Height-----------------------------------
Private Sub dgvShoe_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles dgvShoe.RowsAdded
        dgvShoe.Rows(e.RowIndex).Height = 30
    End Sub

Private Sub btnFilterCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterCancel.Click

  If blnHaveFilter Then
     blnHaveFilter = False

     InputData()       'ให้ Input ข้อมูล

  Else
      StateLockFind(True)
      gpbFilter.Visible = False

  End If
End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        InputData()
    End Sub

    Private Sub txtSeek_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSeek.KeyPress

        ' แปลงเป็นตัวพิมพ์ใหญ่ทันที
        If Char.IsLower(e.KeyChar) Then
            txtSeek.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If

        If e.KeyChar = Chr(13) And txtSeek.Text.Length > 0 Then
            FindDocID()
        End If

    End Sub

    Private Sub txtFilter_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFilter.KeyPress

        ' แปลงเป็นตัวพิมพ์ใหญ่ทันที
        If Char.IsLower(e.KeyChar) Then
            txtFilter.SelectedText = Char.ToUpper(e.KeyChar)
            e.Handled = True
        End If

        If e.KeyChar = Chr(13) And txtFilter.Text.Length > 0 Then
            FilterData()
        End If

    End Sub

End Class
