Imports System.Data
Imports System.IO
Imports System.Drawing

Public Class frmSendPdf

 Dim strDateDefault As String     'ตัวแปรสำหรับวันที่ทั่วไป
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
 Dim dtComputer As Date = Now       'ตัวแปรเก็บค่าวันที่ปัจจุบัน
 Dim strCurrentDate As String       'เก็บค่าสตริงวันที่ปัจจุบัน

     Me.WindowState = FormWindowState.Maximized  'ให้ฟอร์มขยายเต็มหน้าจอ
     StdDateTimeThai()        'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module
     strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

          With txtBegin
               .Text = strCurrentDate
               strDateDefault = strCurrentDate

          End With

     PreDeptSend()  'โหลดแผนกรับเอกสาร
     LoadData()
     Importdata()
     ClearAlldata()  'เคลียร์ข้อมูล
     Timer1.Enabled = True

End Sub

Private Sub PreDeptSend()
 Dim strDept(5) As String
 Dim i As Integer
     strDept(0) = "- แผนกตัดชิ้นส่วน (คุณประดิษฐ์ สังข์ทอง)"
     strDept(1) = "- แผนกฉีด EVA(คุณอิทธิศักดิ์ ปานอุดมกิตติ์)"
     strDept(2) = "- แผนกฉีด PU(คุณทศพร ภูมิสุนทรธรรม)"
     strDept(3) = "- แผนกเย็บ(คุณสถิตถ์ แสนรัก)"
     strDept(4) = "- แผนกผลิตโฟม(คุณเตชินม์ บัวหลง)"
     strDept(5) = "- แผนกฉีด PVC(คุณพีระ แสงอรุณบริสุทธิ์)"

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
             dgvDoc.ScrollBars = ScrollBars.None      'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

             dgvDoc.Rows.Clear() 'เคลี่ยร์ Data grid

             Do While Not .EOF()

                  Select Case .Fields("doc_sta").Value.ToString.Trim

                         Case Is = "0"
                                strSta = "รอดำเนินการ"

                         Case Is = "1"
                                strSta = "รับเอกสารแล้ว"

                         Case Else

                  End Select

                              Select Case .Fields("recv_dept").Value.ToString.Trim

                                     Case Is = "D1"
                                        strDepnm = "- แผนกตัดชิ้นส่วน (คุณประดิษฐ์ สังข์ทอง)"

                                     Case Is = "D2"
                                        strDepnm = "- แผนกฉีด EVA(คุณอิทธิศักดิ์ ปานอุดมกิตติ์)"

                                     Case Is = "D3"
                                        strDepnm = "- แผนกฉีด PU(คุณทศพร ภูมิสุนทรธรรม)"

                                     Case Is = "D4"
                                        strDepnm = "- แผนกเย็บ(คุณสถิตถ์ แสนรัก)"

                                     Case Is = "D5"
                                        strDepnm = "- แผนกผลิตโฟม(คุณเตชินม์ บัวหลง)"

                                     Case Is = "D6"
                                        strDepnm = "- แผนกฉีด PVC(คุณพีระ แสงอรุณบริสุทธิ์)"

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

              dgvDoc.Rows.Clear() 'เคลี่ยร์ Data grid

              StateLockbtn(False)
              StateLockEdit(False)

              btnAdd.Enabled = True

          End If

            dgvDoc.ScrollBars = ScrollBars.Both 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

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
       Loadfiles() 'Upload ไฟล์เอกสาร
       StatelockUpload(True)

    Else
       MessageBox.Show("กรุณาเลือกแผนกที่รับเอกสาร ก่อนแนบไฟล์ข้อมูล", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
       cboDepRecv.DroppedDown = True

    End If

 Else
       MessageBox.Show("กรุณาระบุหัวข้อเอกสาร ก่อนแนบไฟล์อมูล", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

Private Sub Loadfiles()  'ไดอะล็อค Import files
 Dim OpenFileDialog As New OpenFileDialog
 Dim strFullpart As String
 Dim strFilename As String

     With OpenFileDialog

          .CheckFileExists = True
          .ShowReadOnly = False
          .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png;*.pdf;*.xls;*.doc"
          .FilterIndex = 2

          Try

             If .ShowDialog = Windows.Forms.DialogResult.OK Then

                strFullpart = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์
                strFilename = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์

                lblPath.Text = strFullpart
                lblFilename.Text = strFilename

                strStatus = "นำเข้าข้อมูล"

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
     Thayear = Mid(SaveChangeThaYear(strCurrentDate), 3, 2) 'ต้ดปีไทย 5X

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

             .MoveLast()              'เลื่อนไปยัง Record สุดท้าย

             LastNumber = CInt(Mid(.Fields("doc_id").Value.ToString.Trim, 3))  'ตัดสตริง เอา 4 ต้วท้าย  000x
             LastYear = Mid(.Fields("doc_id").Value.ToString.Trim, 1, 2)  'ตัดเอาปี  5x เฉพาะ 2ตัวแรก

               If String.Compare(LastYear, Thayear) = 0 Then       'เปรียบเทียบ สตริงปี 5x
                  LastYear = LastYear
                  LastNumber += 1


               Else
                  LastYear += 1  ' เพิ่มค่า LestRec อีก 1.
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

       If chkDocidExist(strDocid) Then                                               'เช็ครหัส docid ซ้ำ
          strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                   & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                   & Mid(txtBegin.Text.Trim, 1, 2)

             strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"                  'ใส่ "" เพื่อกัน error


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
             LoadData()              'แสดงข้อมูล


                     '------------------------------ค้นหารหัสที่เพิ่มเข้าไปใหม่------------------------------------------

                     For i = 1 To dgvDoc.Rows.Count - 1

                             If dgvDoc.Rows(i).Cells(2).Value.ToString = lblDocid.Text.ToString.Trim Then    'ถ้าคอลัมน์ Size ใน dgvSize มีค่าเท่ากับ txtSize
                                dgvDoc.CurrentCell = dgvDoc.Item(5, i)
                                dgvDoc.Focus()

                             End If
                     Next i
                             StateLockbtn(True)
                             StateLockEdit(False)

                     strStatus = ""  'เคลี่ยร์ตัวแปร
                     ClearAlldata()

       Else
          MsgBox("รหัสเอกสารซ้ำ โปรดเลือกรหัสใหม่ ")
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
  Dim fname As String = String.Empty  'เท่ากับ ""
  Dim dFile As String = String.Empty
  Dim dFilePath As String = String.Empty

  Dim fServer As String = String.Empty
  Dim intResult As Integer   'คืนค่าเป็นจำนวนเต็ม

  On Error GoTo Err70

        fname = strPicPath & "\" & strPicName  'พาร์ท \\10.32.0.15\data1\EquipDocument\"ชื่อรูปภาพ"
        fServer = PthName & "\" & strPicName    'partServer \\10.32.0.15\data1\EquipDocument\"ชื่อรูปภาพ"
        If File.Exists(fServer) Then    'ถ้าไฟล์มีอยู่จริง
           CallCopyPicture = True      'ให้คืนค่า true

        Else

            If File.Exists(fname) Then
               dFile = Path.GetFileName(fname)
               dFilePath = DrvName + dFile


               intResult = String.Compare(fname.ToString.Trim, dFilePath.ToString.Trim)

                    '-------------------------- ถ้าค่าเป็น 0 แสดงว่าโหลดไฟล์ใช้อยู่ ไม่สามารถ Copy ไฟล์ได้ ------------------------------

                    If intResult = 1 Then 'ค่าที่ได้ = 1 ถึง copy รูปมาไว้ที่เครื่อง 10.32.0.14
                       File.Copy(fname, dFilePath, True)
                    End If
                    CallCopyPicture = True

            Else
                CallCopyPicture = True

            End If

        End If

Err70:

        If Err.Number <> 0 Then

           MsgBox("UserName ของคุณไม่มีสิทธิแก้ไขรูปภาพได้!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Permission Can't Edit Picture")
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

                   If ChkStaRecv() Then     'เช็คว่าเอกสารถูกรับเข้าแล้วหรือยัง
                      StateLockEdit(True)
                      strStatus = "แก้ไขข้อมูล"
                      LockEditData()

                   Else

                      MessageBox.Show("เอกสารถูกรับเข้าแล้ว ไม่สามารถแก้ไขได้", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                   End If

            Else
                MsgBox("กรุณาเลือกรายการ ก่อนแก้ไขข้อมูล")
            End If

        StateLockbtn(True)

     Else
         MsgBox("ไม่มีข้อมูลสำหรับแก้ไข")

     End If
End Sub

Private Function ChkStaRecv() As Boolean           'เช็คว่าเอกสารถูกรับแล้วหรือยัง
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

               Return False           'เอกสารถูกรับเข้าแล้ว

            Else

               Return True            'เอกสารยังไม่ได้รับเข้า

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

              strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"    'ใส่ "" เพื่อกัน error


             '--------------------------- Copy เอกสาร ----------------------------------------------

             blnReturnCopyPic = CallCopyPicture(lblPath.Text.ToString.Trim, lblFilename.Text.ToString.Trim)

             If blnReturnCopyPic Then
                lblPath.Text = PthName

             Else
                txtName.Text = ""
                lblFilename.Text = ""

             End If

                     '---------------------------- UPDATE ข้อมูลในตาราง eqpmst ------------------------------

                      strSqlCmd = "UPDATE docsend SET filename ='" & ReplaceQuote(lblFilename.Text.ToString.Trim) & "'" _
                                           & "," & "content_type ='" & ReplaceQuote(txtName.Text.ToString.Trim) & "'" _
                                           & "," & "remark = '" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                           & " WHERE doc_id ='" & lblDocid.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)

                      LoadData()
                      strStatus = "" 'เคลี่ยร์ค่าตัวแปร

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
                               strDepnm = "- แผนกตัดชิ้นส่วน (คุณประดิษฐ์ สังข์ทอง)"

                          Case Is = "D2"
                               strDepnm = "- แผนกฉีด EVA(คุณอิทธิศักดิ์ ปานอุดมกิตติ์)"

                          Case Is = "D3"
                               strDepnm = "- แผนกฉีด PU(คุณทศพร ภูมิสุนทรธรรม)"

                          Case Is = "D4"
                               strDepnm = "- แผนกเย็บ(คุณสถิตถ์ แสนรัก)"

                          Case Is = "D5"
                               strDepnm = "- แผนกผลิตโฟม(คุณเตชินม์ บัวหลง)"

                          Case Is = "D6"
                               strDepnm = "- แผนกฉีด PVC(คุณพีระ แสงอรุณบริสุทธิ์)"

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

      btyConsider = MsgBox("รหัสเอกสาร : DC" & strDocid.ToString.Trim & vbNewLine _
                                            & "หัวข้อเอกสาร : " & strContent.ToString.Trim & vbNewLine _
                                            & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                            + MsgBoxStyle.Exclamation, "Confirm Delete Data")

            If btyConsider = 6 Then
               DeleteFileImport()      'ลบเอกสารใน  \\10.32.0.15\data1
               DeleteData(strDocid)
               ClearAlldata()

            End If

   Else
       MsgBox("ไม่มีข้อมูลให้ดำเนินการ")

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

               LoadData() 'รีเฟรชข้อมูล
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                Case Is = 38 'ปุ่มลูกศรขึ้น    

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtRemark.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtRemark.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    btnUpload.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                        If .SelectionStart = 0 Then
                           txtName.Focus()
                        End If
                Case Is = 38 'ปุ่มลูกศรขึ้น   
                        If .SelectionStart = 0 Then
                           txtName.Focus()
                        End If

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        btnUpload.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            btnUpload.Focus()
                        End If
                    End If

                Case Is = 40  'ปุ่มลง  
                    btnUpload.Focus()

                Case Is = 113 'ปุ่ม F2
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
  LoadData() 'รีเฟรชข้อมูล
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

  'เคลียร์ข้อมูลใน table tmp_eqptrn
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
                Case Is = "0"  'ลบข้อมูลหลังปิดฟอร์ม
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

             Case Is = "นำเข้าข้อมูล"
                   SaveData() 'บันทึกข้อมูลใหม่

             Case Is = "แก้ไขข้อมูล"
                   SaveEditdata() 'บันทึกแก้ไขข้อมูล

      End Select

     ClearAlldata()
     btnUpload.Enabled = False

  Else
     MsgBox("กรุณาระบุ หัวข้อเอกสาร")
     txtName.Focus()
  End If

End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
  ClearAlldata()
  GenDocid()
  strStatus = "เพิ่มข้อมูล"

  stateLockAdd(True)
  StateLockbtn(False)
  txtName.Focus()


End Sub

'----------------------- เปิดไฟล์เอกสารที่แนบ -----------------------------------------------------------
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

Private Sub DeleteFileImport()   'ลบเอกสารใน 10.32.0.15
 Dim Path As String = DrvName
 Dim strDocid As String
 Dim strFilename As String
      strDocid = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(2).Value.ToString.Trim
      strFilename = dgvDoc.Rows(dgvDoc.CurrentRow.Index).Cells(4).Value.ToString.Trim

       Path = System.IO.Path.GetDirectoryName(Path) & "\" & strFilename

       '--------------- ลบเอกสารใน  \\10.32.0.15\data1\Eqpdocument\

       If System.IO.File.Exists(Path) = True Then
          System.IO.File.Delete(Path)

       End If

End Sub
End Class