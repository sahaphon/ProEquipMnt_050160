Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmPastMold
    Dim IsShowSeek As Boolean        'ตัวเเปรแสดงสถานะ gpbSeek
    Dim strDateDefault As String     'ตัวแปรสำหรับวันที่ทั่วไป

    Public Const DrvName As String = "\\10.32.0.15\data1\EquipPicture\"
    Public Const PthName As String = "\\10.32.0.15\data1\EquipPicture"         'ตัวแปรสำหรับเก็บ part รูปภาพ

   Private tt As ToolTip = New ToolTip 'แสดงทุูลทิป ในรูปภาพเวลาเลื่อนเคอร์เซอร์

Protected Overrides ReadOnly Property CreateParams() As CreateParams       'ป้องกันการปิดโดยใช้ปุ่ม Close Button(ปุ่มกากบาท)
  Get
       Dim cp As CreateParams = MyBase.CreateParams
           Const CS_DBLCLKS As Int32 = &H8
           Const CS_NOCLOSE As Int32 = &H200
           cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
           Return cp
       End Get
End Property

Private Sub frmPastMold_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTable(0, "")  'ลบข้อมูล Table tmp_eqptrn where user_id..
   ClearTmpTable(2, "") 'ล้างข้อมูลตาราง tmp_fixeqptrn
   ClearTmpTable(3, "") 'ล้างข้อมูลตาราง tmp_eqptrn_newsize
   frmEqpSheet.lblCmd.Text = "0"  'เคลียร์สถานะ
   Me.Dispose()     'ทำลายฟอร์ม คืนหน่วยความจำ
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

                Case Is = "0"  'ลบข้อมูลหลังปิดฟอร์ม
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

 Dim dtComputer As Date = Now       'ตัวแปรเก็บค่าวันที่ปัจจุบัน
 Dim strCurrentDate As String       'เก็บค่าสตริงวันที่ปัจจุบัน

     StdDateTimeThai()        'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module
     strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

     ClearDataGpbHead()
     PrePartSeek()            'โหลดรายละเอียดใส่ใน cbo ชิ้้นส่วนที่ผลิต
     PreMoldStatus()            'สถานะโมล์ด

     Select Case frmEqpSheet.lblCmd.Text.ToString

            Case Is = "0" 'เพิ่มข้อมูล

                With txtBegin
                     .Text = strCurrentDate
                     strDateDefault = strCurrentDate
                End With

                With Me
                     .Text = "เพิ่มข้อมูล"
                     txtEqp_id.Focus()
                End With

                '--------------------- เวลาเพิ่มข้อมูลไม่ต้องแสดงสถานะ(ซ่อนคอลัมน์ใน Gridview) ----------------------------

                dgvSize.Columns(0).Visible = False  'ซ่อนคอลัมน์ที่ 1
                dgvSize.Columns(1).Visible = False  'ซ่อนคอลัมน์ที่ 2
                dgvSize.Columns(2).Visible = False  'ซ่อนคอลัมน์ที่ 3
                dgvSize.Columns(3).Visible = False  'ซ่อนคอลัมน์ที่ 4

            Case Is = "1" 'แก้ไขข้อมูล

                With Me
                     .Text = "เเก้ไขข้อมูล"
                End With

                LockEditData()
                txtEqp_id.ReadOnly = True   'ให้อ่านอย่างเดียว
                txtEqpnm.ReadOnly = True
                txtShoe.ReadOnly = True
                txtOrder.ReadOnly = True
                txtRemark.ReadOnly = True
                cboPart.Enabled = False

            Case Is = "2"   'มุมมองข้อมูล

                With Me
                     .Text = "มุมมองข้อมูล"
                End With

                LockEditData()
                txtEqp_id.ReadOnly = True  'ให้อ่านอย่างเดียว
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

 Dim strCmd As String              'เก็บสตริง Command

 Dim strLoadFilePicture As String   'เก็บค่าสตริงโหลด Picture
 Dim strPathPicture As String = "\\10.32.0.15\data1\EquipPicture\"   'เก็บ part

 Dim blnHavedata As Boolean       'เก็บค่าตัวเเปร สำหรับเช็คว่ามีข้อมูลหรือไม่
 Dim strSqlSelc As String = ""    'เก็บสตริง sql select
 Dim strPart As String = ""

 Dim strCode As String = frmEqpSheet.dgvShoe.Rows(frmEqpSheet.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim

     With Conn
         If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

    End With
        ' NOLOCK หมายถึงสามารถทำงานกับ Table ได้แม้ถูกใช้จาก user อื่น
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

                '-----------------------------------ใส่ชิ้นงาน------------------------------

                Select Case .Fields("part").Value.ToString.Trim
                    Case Is = "SOLE2"
                        cboPart.Text = "พื้นล่าง"
                    Case Is = "SOLE1"
                        cboPart.Text = "พื้นบน"
                    Case Is = "FILL"
                       cboPart.Text = "ใส้พื้นบน"
                    Case Is = "LOGOHEEL"
                        cboPart.Text = "โลโก้ส้น"
                    Case Is = "LOGO"
                        cboPart.Text = "โลโก้พื้น"
                    Case Is = "EVAH"
                        cboPart.Text = "EVA ติดส้น"
                    Case Is = "EVAF"
                        cboPart.Text = "EVA บนหนังหน้า"
                    Case Is = "UPPER"
                       cboPart.Text = "หนังหน้า"
                    Case Is = "ONUPPER"
                       cboPart.Text = "ONUPPER"
                    Case Is = "ONUPPER"
                       cboPart.Text = "ONUPPER"
                End Select

                '-------------------------------Load รูปภาพ(รูปแบบแผง)-----------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_ctain").Value.ToString.Trim
                ' เช็ดพาร์ทไฟล์มีอยู่จริง
                If File.Exists(strLoadFilePicture) Then

                    Dim img1 As Image      'ประกาศตัวแปร img1 เพื่อเก็บภาพ
                    img1 = Image.FromFile(strLoadFilePicture)  'img1 เท่ากับpicture ที่โหลดมาจาก db
                    picEqp1.Image = ScaleImage(img1, picEqp1.Height, picEqp1.Width)

                Else
                    picEqp1.Image = Nothing  'ถ้าไม่มีรูปภาพให้ picEqp1 ว่างเปล่า
                End If

                '-------------------------------Load รูปภาพเริ่มชิ้นงาน ----------------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_io").Value.ToString.Trim
                If File.Exists(strLoadFilePicture) Then
                    Dim img2 As Image
                    img2 = Image.FromFile(strLoadFilePicture)
                     picEqp2.Image = ScaleImage(img2, picEqp2.Height, picEqp2.Width)

                Else
                    picEqp2.Image = Nothing
                End If


                 '-------------------------------Load รูปภาพรองเท้าสำเร็จรูป  ----------------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_part").Value.ToString.Trim
                If File.Exists(strLoadFilePicture) Then
                    Dim img3 As Image
                    img3 = Image.FromFile(strLoadFilePicture)
                    picEqp3.Image = ScaleImage(img3, picEqp3.Height, picEqp3.Width)
                Else
                    picEqp3.Image = Nothing
                End If


                strCmd = frmEqpSheet.lblCmd.Text.ToString.Trim    'ให้ strCmd เท่ากับค่าใน lblcmd ในฟอร์ม frmEqpSheet

                Select Case strCmd
                       Case Is = "1"   'ให้ล็อคตอนแก้ไข
                       Case Is = "2"   'ให้ล็อคตอนมุมมอง
                            btnSaveData.Enabled = False  'ปิดปุ่ม "บันทึกข้อมูล"
                End Select

                '------------------------------- อ่านข้อมูลการส่งซ่อมไว้ในตาราง tmp_eqptrn --------------------

                 strSqlSelc = "INSERT INTO tmp_fixeqptrn " _
                                  & "SELECT user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM fixeqptrn " _
                                  & " WHERE eqp_id = '" & strCode & "'" _
                                  & " AND fix_sta= '" & "1" & "'"

                Conn.Execute(strSqlSelc)

                '------------------------------- บันทึกข้อมูลงในตาราง tmp_eqptrn----------------------------

                strSqlSelc = "INSERT INTO tmp_eqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM eqptrn " _
                                  & " WHERE eqp_id = '" & strCode & "' "

                Conn.Execute(strSqlSelc)

                '------------------------------ เรียง SIZE ใหม่ -------------------------------------------

                ReSizeSort(strCode)   'จัดเรียง size เสียใหม่

                blnHavedata = True     'มีข้อมูล

        Else
                blnHavedata = False    'ไม่มีข้อมูล
        End If

            .ActiveConnection = Nothing   'สั่งปิดการเชื่อมต่อ
            .Close()
        End With

        Rsd = Nothing   'เคลียร์ค่า RecordSet
        Conn.Close()    'สั่งตัดการเชื่อมต่อ
        Conn = Nothing  'เคลี่ย์ Connection

             If blnHavedata Then          'ถ้า blnHavedata = true
                ShowScrapItem()
             End If
End Sub

Private Sub ShowScrapItem()  'วนลูปเอาเพื่อแสดงค่า จำนวน SET, รวมคู่ลงผลิต, รวมราคาอุปกรณ์

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmdSelc As String    'เก็บค่า string command

 Dim dubQty As Double
 Dim dubAmt As Double
 Dim sngSetQty As Single     'เก็บจำนวน SET
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
            dgvSize.ScrollBars = ScrollBars.None 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

            If .RecordCount <> 0 Then

                Do While Not .EOF()

                   mold_id = .Fields("eqp_id").Value.ToString.Trim
                   strArr = Split(.Fields("size_desc").Value.ToString.Trim, "-")  'ตัด array ออกมา
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

                txtSet.Text = sngSetQty.ToString.Trim        'จำนวน SET
                txtAmount.Text = Format(dubQty, "#,##0")     'รวมคู่ลงผลิต
                lblAmt.Text = Format(dubAmt, "#,##0.00")     'รวมราคาอุปกรณ์

            Else
                txtSet.Text = "0.0"
                txtAmount.Text = "0"
                lblAmt.Text = "0.00"

            End If

            .ActiveConnection = Nothing
            .Close()
            Rsd = Nothing

            dgvSize.ScrollBars = ScrollBars.Both 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

        End With

        Conn.Close()
        Conn = Nothing

    End Sub

Private Sub PrePartSeek()    'โหลดรายละเอียดใส่ใน cbo ชิ้้นส่วนที่ผลิต
   Dim strGpTopic(8) As String
   Dim i As Byte

     strGpTopic(0) = "พื้นบน"
     strGpTopic(1) = "ใส้พื้นบน"
     strGpTopic(2) = "โลโก้ส้น"
     strGpTopic(3) = "พื้นล่าง"
     strGpTopic(4) = "EVA ติดส้น"
     strGpTopic(5) = "โลโก้พื้น"
     strGpTopic(6) = "หนังหน้า"
     strGpTopic(7) = "EVA บนหนังหน้า"
     strGpTopic(8) = "ONUPPER"

        With cboPart

            For i = 0 To 8
                .Items.Add(strGpTopic(i))
            Next i

        End With

End Sub

Private Sub CheckDataBfSave()

  Dim intListwc As Integer = dgvSize.Rows.Count  'ตัวแปร inListwc เก็บจำนวนเรคคอร์ดของ dgvSize
  Dim strProd As String = ""
  Dim strProdnm As String = ""

  Dim bytConSave As Byte  'เก็บค่า megbox 

        If txtEqp_id.Text.ToString.Trim <> "" Then

            If txtEqpnm.Text.ToString.Trim <> "" Then

                If cboPart.Text.ToString.Trim <> "" Then

                    If intListwc > 0 Then  'ถ้า dgvSize มีข้อมูล

                        'ถ้า MsgBox = ปุ่ม Yes 
                           bytConSave = MsgBox("คุณต้องการบันทึกข้อมูลใช่หรือไม่!" _
                                  , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")


                               If bytConSave = 6 Then

                                        Select Case Me.Text
                                               Case Is = "เพิ่มข้อมูล"

                                                    If CheckCodeDuplicate() Then
                                                       SaveNewRecord()

                                                    Else
                                                       MessageBox.Show("กรุณากรอกรหัสอุปกรณ์ใหม่!....", "***รหัสอุปกรณ์ซ้ำ***", _
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
                             ShowResvrd()       'แสดงฟอร์มย่อย gpbSeek 
                             gpbSeek.Text = "เพิ่มข้อมูล"
                             txtSize.ReadOnly = False

                        Else

                             MessageBox.Show("กรุณากรอกรหัสอุปกรณ์ใหม่!....", "รหัสอุปกรณ์ซ้ำ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                             txtEqp_id.Text = ""
                             txtEqp_id.ReadOnly = False
                             txtEqp_id.Enabled = True
                             txtEqp_id.Focus()

                        End If


                    End If

                Else
                    MsgBox("โปรดระบุข้อมูลรายละเอียดอุปกรณ์  " _
                                & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                    cboPart.DroppedDown = True
                    cboPart.Focus()

                End If

            Else
                MsgBox("โปรดระบุข้อมูลรายละเอียดอุปกรณ์  " _
                              & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                txtEqpnm.Focus()
            End If

        Else
            MsgBox("โปรดระบุข้อมูลรหัสอุปกรณ์  " _
                          & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
            txtEqp_id.Focus()
        End If

End Sub

Private Sub SaveEditRecord()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

  Dim dateSave As Date = Now()
  Dim strDate As String = ""

  Dim strCredate As String
  Dim strDocdate As String           'เก็บสตริงวันที่เอกสาร
  Dim strGpType As String = ""       'เก็บประเภทอุปกรณ์
  Dim strPartType As String = ""     'เก็บชิ้นส่วนที่ผลิต

  Dim blnReturnCopyPic As Boolean

        With Conn
            If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()

        End With

               Conn.BeginTrans()      'จุดเริ่มต้น Transection

               strDate = dateSave.Date.ToString("yyyy-MM-dd")
               strDate = SaveChangeEngYear(strDate)

              '------------------------- วันที่เอกสาร ----------------------------------------------------

               strDocdate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocdate = SaveChangeEngYear(strDocdate)


        '---------------------------------------- วดป.ที่ผลิด --------------------------------------------

        If txtCdate.Text <> "__/__/____" Then

            strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtCdate.Text.ToString, 1, 2)
            strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

        Else
            strCredate = "NULL"

        End If

        '------------------------------------ กำหนดกลุ่มของชิ้นงาน -----------------------------------------

             Select Case cboPart.Text.ToString
                    Case Is = "พื้นล่าง"
                        strPartType = "SOLE2"
                    Case Is = "พื้นบน"
                        strPartType = "SOLE1"
                    Case Is = "ใส้พื้นบน"
                        strPartType = "FILL"
                    Case Is = "โลโก้ส้น"
                        strPartType = "LOGOHEEL"
                    Case Is = "โลโก้พื้น"
                        strPartType = "LOGO"
                    Case Is = "EVA ติดส้น"
                        strPartType = "EVAH"
                    Case Is = "EVA บนหนังหน้า"
                        strPartType = "EVAF"
                    Case Is = "หนังหน้า"
                        strPartType = "UPPER"
                    Case Is = "ONUPPER"
                        strPartType = "ONUPPER"

            End Select


                      '------------------------------------ บันทึกรูปแผง ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath1.Text.ToString.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath1.Text = PthName

                       Else
                          lblPicName1.Text = ""
                          lblPicPath1.Text = ""
                          picEqp1.Image = Nothing

                       End If


                      '------------------------------------ บันทึกรูปโมล์ด ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath2.Text.ToString.Trim, ReturnImageName(lblPicName2.Text.ToString.Trim), lblPicName2.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath2.Text = PthName

                       Else
                          lblPicName2.Text = ""
                          lblPicPath2.Text = ""
                          picEqp2.Image = Nothing

                       End If

                      '------------------------------------ บันทึกรูป PIC_PART------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath3.Text.ToString.Trim, ReturnImageName(lblPicName3.Text.ToString.Trim), lblPicName3.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath3.Text = PthName

                       Else
                          lblPicName3.Text = ""
                          lblPicPath3.Text = ""
                          picEqp3.Image = Nothing

                       End If

                      '---------------------------- UPDATE ข้อมูลในตาราง eqpmst ------------------------------

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


        '------------------------------------------------ลบข้อมูลในตาราง eqptrn--------------------------------------------------------------------

                     strSqlCmd = "Delete FROM eqptrn" _
                                          & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlCmd)

        '-----------------------------------------บันทึกข้อมูลในตาราง eqptrn โดย Select จาก tmp_eqptrn-------------------------------------------------

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
        Conn.CommitTrans()  'สั่ง Commit transection

        frmEqpSheet.lblCmd.Text = txtEqp_id.Text.ToString.Trim
        frmEqpSheet.Activating()
        Me.Close()

        Conn.Close()
        Conn = Nothing

    End Sub

    Private Sub SaveNewRecord()

        Dim Conn As New ADODB.Connection
        Dim strSqlCmd As String
        Dim dateSave As Date = Now()    'เก็บค่าวันที่ปัจจุบัน
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
                    strDate = SaveChangeEngYear(strDate)               'เเปลงเป็นปี ค.ศ.

                    strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                   & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                   & Mid(txtBegin.Text.Trim, 1, 2)
                    strDateDoc = SaveChangeEngYear(strDateDoc)

                    '---------------------------------------- บันทึกรูปแบบแผง ----------------------------------------------------

                    blnRetuneCopyPic = CallCopyPicture(lblPicPath1.Text.Trim, ReturnImageName(lblPicName1.Text.Trim), lblPicName1.Text.Trim)

                    If blnRetuneCopyPic Then       'ถ้า CallCopyPicture = true
                       lblPicPath1.Text = PthName
                    Else
                       lblPicPath1.Text = ""
                       lblPicName1.Text = ""
                       picEqp1 = Nothing

                    End If

                    '---------------------------------------- บันทึกรูปเริ่มชิ้นงาน  ----------------------------------------------------
                     blnRetuneCopyPic = CallCopyPicture(lblPicPath2.Text.Trim, ReturnImageName(lblPicName2.Text.Trim), lblPicName2.Text.Trim)

                    If blnRetuneCopyPic Then
                       lblPicPath2.Text = PthName

                    Else
                       lblPicPath2.Text = ""
                       lblPicName2.Text = ""
                       picEqp2 = Nothing
                    End If


                    '---------------------------------------- บันทึกรูปรองเท้าสำเร็จรูป ----------------------------------------------------
                     blnRetuneCopyPic = CallCopyPicture(lblPicPath3.Text.Trim, ReturnImageName(lblPicName3.Text.Trim), lblPicName3.Text.Trim)

                    If blnRetuneCopyPic Then
                       lblPicPath3.Text = PthName

                    Else
                       lblPicPath3.Text = ""
                       lblPicName3.Text = ""
                       picEqp3 = Nothing
                    End If


                    '---------------------------------------- วดป.ที่ผลิด ----------------------------------------------------
                    If txtCdate.Text <> "__/__/____" Then

                       strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtCdate.Text.ToString, 1, 2)
                       strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

                   Else
                       strCredate = "NULL"
                   End If


                    '---------------------------------------- วดป.เปิดใบสั่งซื้อ -------------------------------------------------
                   If txtPrdate.Text <> "__/__/____" Then

                      strPRdate = Mid(txtPrdate.Text.ToString, 7, 4) & "-" _
                                    & Mid(txtPrdate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtPrdate.Text.ToString, 1, 2)
                      strPRdate = "'" & SaveChangeEngYear(strPRdate) & "'"

                   Else
                      strPRdate = "NULL"
                   End If


                    '---------------------------------------- วันทีนัดเข้า -------------------------------------------------
                   If txtFCdate.Text <> "__/__/____" Then

                       strFCdate = Mid(txtFCdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtFCdate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtFCdate.Text.ToString, 1, 2)
                       strFCdate = "'" & SaveChangeEngYear(strFCdate) & "'"

                   Else
                       strFCdate = "NULL"
                   End If


                    '---------------------------------------- วดป.ที่รับเข้า -------------------------------------------------
                  If txtIndate.Text <> "__/__/____" Then

                     strINdate = Mid(txtIndate.Text.ToString, 7, 4) & "-" _
                                    & Mid(txtIndate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtIndate.Text.ToString, 1, 2)
                     strINdate = "'" & SaveChangeEngYear(strINdate) & "'"

                 Else
                     strINdate = "NULL"
                 End If

                    '------------------------------------กำหนดกลุ่มของชิ้นงาน--------------------------------------------------

                    Select Case cboPart.Text.ToString.Trim

                        Case Is = "พื้นล่าง"
                            strPartType = "SOLE2"
                        Case Is = "พื้นบน"
                            strPartType = "SOLE1"
                        Case Is = "ใส้พื้นบน"
                            strPartType = "FILL"
                        Case Is = "โลโก้ส้น"
                            strPartType = "LOGOHEEL"
                        Case Is = "โลโก้พื้น"
                            strPartType = "LOGO"
                        Case Is = "EVA ติดส้น"
                            strPartType = "EVAH"
                        Case Is = "EVA บนหนังหน้า"
                            strPartType = "EVAF"
                        Case Is = "หนังหน้า"
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


                   '------------------------------------------------บันทึกข้อมูลในตาราง eqptrn----------------------------------------------------------

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

    Private Function RetrnAmount() As String   'ฟังก์ชั่นหายอดรวม ราคาอุปกรณ์ตาม UserLogin

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

        '-------------- คำสั่ง SELCT SUM()AS คอลัมน์ใหม่ ---------------------------------------

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

    'ฟังก์ชั่น CopyPicture
Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String, ByVal strNewPicName As String) As Boolean

  Dim fname As String = String.Empty  'เท่ากับ ""
  Dim dFile As String = String.Empty
  Dim dFilePath As String = String.Empty

  Dim fServer As String = String.Empty
  Dim intResult As Integer   'คืนค่าเป็นจำนวนเต็ม

  On Error GoTo Err70

        fname = strPicPath & "\" & strPicName  'พาร์ท \\10.32.0.15\data1\EquipPicture\"ชื่อรูปภาพ"
        fServer = PthName & "\" & strNewPicName    'partServer \\10.32.0.15\data1\EquipPicture\"ชื่อรูปภาพ"

        If File.Exists(fServer) Then    'ถ้าไฟล์มีอยู่จริง
           CallCopyPicture = True      'ให้คืนค่า true

        Else

            If File.Exists(fname) Then
               dFile = Path.GetFileName(fname)
               dFilePath = DrvName + dFile

               intResult = String.Compare(fname.ToString.Trim, dFilePath.ToString.Trim)

                '------------------------------------ถ้าค่าเป็น 0 แสดงว่าโหลดไฟล์ใช้อยู่ ไม่สามารถ Copy ไฟล์ได้------------------------------

                    If intResult = 1 Then 'ค่าที่ได้ = 1 ถึง copy รูปมาไว้ที่เครื่อง 10.32.0.14
                       File.Copy(fname, dFilePath, True)
                    End If

                    My.Computer.FileSystem.RenameFile(dFilePath, strNewPicName)  'เปลี่ยนชื่อไฟล์รูปใหม่
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

    Private Function CheckCodeDuplicate() As Boolean     'ฟังก์ชั่นเช็ครหัสซ้ำ

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
            .ActiveConnection = Nothing   'เคลียร์ Connection
            .Close()

        End With
        Rsd = Nothing   'เคลียร์ RecordSet
        Conn.Close()    'ปิดการเชื่อมต่อ
        Conn = Nothing   'เคลียร์ RecordSet

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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                Case Is = 38 'ปุ่มลูกศรขึ้น

                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtEqpnm.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtEqpnm.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                     txtOrder.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
    End Sub

Private Sub txtEqp_id_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqp_id.KeyPress

  Select Case Asc(e.KeyChar)

         Case 45 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
              e.Handled = False

         Case 8, 46 ' Backspace = 8,  Delete = 46
              e.Handled = False

         Case 13   'Enter = 13
              e.Handled = False
              txtEqpnm.Focus()

         Case Else
              e.Handled = True
              MsgBox("กรุณาระบุข้อมูลเป็นภาษาอังกฤษ", MsgBoxStyle.Critical, "ผิดพลาด")
  End Select

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBfSave()
End Sub

Private Sub ShowResvrd()  'ให้แสดง GroupBox gpbSeek ขึ้นมา

  tabMain.SelectedTab = tabSize
  IsShowSeek = Not IsShowSeek  'ถ้าสถานะแถบ seek ไม่แสดง ให้แสดง

  If IsShowSeek Then

     With gpbSeek
          .Visible = True
          .Left = 10    'แกน X
          .Top = 215   'แกน Y 252
          .Height = 504
          .Width = 1006
    End With

    StateLockFindDept(False) ' ล็อค FindDept โดยส่งค่าเป็น False ไป
    txtTdesc.Focus()

 Else
      StateLockFindDept(True)
 End If

End Sub

    Private Sub StateLockFindDept(ByVal Sta As Boolean)  'ล็อค gpbHead ให้ปุ่ม enabled = false
        Dim strMode As String = frmEqpSheet.lblCmd.Text.ToString   'ตัวแปร lblCmd ใน frmEqpSheet ส่งค่ามาให้ strMode 

        btnAdd.Enabled = Sta    'ปุ่มเพิ่มข้อมูล
        gpbHead.Enabled = Sta

        tabMain.Enabled = Sta

        btnSaveData.Enabled = Sta  'ปุ่มบันทึกข้อมูล

        Select Case strMode

            Case Is = "1" 'แก้ไขข้อมูล                        
            Case Is = "2" 'มุมมองข้อมูล
                btnSaveData.Enabled = False
        End Select

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        ShowResvrd()
        ClearSubData2()
        CallEditData()    'ซับรูทีนแสดง Size เพื่อแก้ไขข้อมูล
        'CallEditData2()
        gpbSeek.Text = "เพิ่มข้อมูล"

        txtPId.ReadOnly = False
        txtSize.ReadOnly = False
        txtSizeDesc.ReadOnly = False
        txtPId.Focus()
    End Sub

    Private Sub btnSeekExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekExit.Click
        StateLockFindDept(True)
        gpbSeek.Text = ""
        gpbSeek.Visible = False  'ทำให้ gpbSeek ซ่อน
        IsShowSeek = False
    End Sub

Private Sub CheckSubDataBfSave()  'เช็คข้อมูล gpbSeek ก่อน save ข้อมูล
Dim i As Integer

   If txtPId.Text.Trim <> "" Then

        If txtSize.Text.Trim <> "" Then

                 If gpbSeek.Text = "เพิ่มข้อมูล" Then
                    SaveSubRecord()
                    ReSizeSort(txtEqp_id.Text.Trim)

                 Else
                    EditSubRecord()
                    ReSizeSort(txtEqp_id.Text.Trim)
                 End If

                     ShowScrapItem()  'แสดข้อมูลที่่บันทึใน dgvSize โดย Select จาก v_tmp_eqptrn

                     '------------------------------ค้นหารหัสที่เพิ่มเข้าไปใหม่------------------------------------------
                     For i = 1 To dgvSize.Rows.Count - 1

                             If dgvSize.Rows(i).Cells(4).Value.ToString = txtSize.Text.ToString.Trim Then    'ถ้าคอลัมน์ Size ใน dgvSize มีค่าเท่ากับ txtSize
                                dgvSize.CurrentCell = dgvSize.Item(5, i)
                                dgvSize.Focus()
                             End If

                     Next i
                             StateLockFindDept(True)
                             gpbSeek.Visible = False
                             IsShowSeek = False


         Else
               MsgBox("โปรดระบุ Size" _
                                   & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
               txtSize.Focus()
         End If

   Else
        MsgBox("โปรดระบุรหัสแผง" _
                       & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
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

             '----------------------------------------วันที่เปิดซื้อ-----------------------------------------------

             If txtPrdate.Text <> "__/__/____" Then

                strPrDate = Mid(txtPrdate.Text.ToString, 7, 4) & "-" _
                                & Mid(txtPrdate.Text.ToString, 4, 2) & "-" _
                                & Mid(txtPrdate.Text.ToString, 1, 2)
                                strPrDate = "'" & SaveChangeEngYear(strPrDate) & "'"

            Else
                strPrDate = "NULL"
            End If

           '----------------------------------------วันที่รับเข้า--------------------------------------------------

            If txtIndate.Text <> "__/__/____" Then

               strRecvDate = Mid(txtIndate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtIndate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtIndate.Text.ToString, 1, 2)
                                 strRecvDate = "'" & SaveChangeEngYear(strRecvDate) & "'"

           Else
               strRecvDate = "NULL"
           End If


          '----------------------------------------วันที่นัดเข้า---------------------------------------------------

           If txtFCdate.Text <> "__/__/____" Then

              strFcDate = Mid(txtFCdate.Text.ToString, 7, 4) & "-" _
                              & Mid(txtFCdate.Text.ToString, 4, 2) & "-" _
                              & Mid(txtFCdate.Text.ToString, 1, 2)
                              strFcDate = "'" & SaveChangeEngYear(strFcDate) & "'"
           Else
               strFcDate = "NULL"
           End If


            '------------------------------------กำหนดกลุ่มประเภท------------------------------------------------


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
          MsgBox("พบข้อผิดพลาดขณะทำการบันทึก โปรดดำเนินการใหม่อีกครั้ง", MsgBoxStyle.Critical, "ผิดพลาด")
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

 Dim dateSave As Date = Now()            'เก็บสตริงวันที่ปัจจุบัน
 Dim strDate As String = ""              'เก็บสตริงวันที่
 Dim strDateDoc As String = ""
 Dim strCreDate As String = ""           'วันที่ผลิต
 Dim strPrdate As String = ""            'วันที่เปิดใบ PR
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

        '------------------------------------เช็คข้อมูล่ก่อนว่ามีอยู่หรือเปล่า-------------------------------------------------

        strSqlSelec = "SELECT * FROM tmp_eqptrn " _
                          & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                          & "AND size_id = '" & txtSize.Text.ToString.Trim & "'" _
                          & "AND size_desc = '" & txtPId.Text.ToString.Trim & "'" _
                          & "AND size_group = '" & txtSizeDesc.Text & "'"

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelec, Conn, , , )

            If .RecordCount <> 0 Then        'ถ้า RecordSet มีข้อมูล
                MessageBox.Show("Size :" & txtSize.Text.ToString & " มีในระบบแล้ว กรุณาระบุ Size ใหม่", "ข้อมูลซ้ำ!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                SaveSubRecord = False

            Else

                strDate = dateSave.Date.ToString("yyyy-MM-dd")
                strDate = SaveChangeEngYear(strDate)

                strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                      & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                      & Mid(txtBegin.Text.ToString, 1, 2)
                strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"


                '---------------------------------------- วดป.ที่ผลิด ----------------------------------------------------
                If txtCdate.Text <> "__/__/____" Then

                    strCreDate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                         & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                         & Mid(txtCdate.Text.ToString, 1, 2)
                    strCreDate = "'" & SaveChangeEngYear(strCreDate) & "'"

                Else
                    strCreDate = "NULL"
                End If

                '--------------------------วันที่เปิดซื้อ ------------------------------------------------

                If txtCdate.Text <> "__/__/____" Then   'ตัด ปี เดือน วัน
                    strCreDate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtCdate.Text.ToString, 1, 2)
                    strCreDate = "'" & SaveChangeEngYear(strCreDate) & "'"     'เเปลงค่าเป็นปี คศ.(ใน module)
                Else
                    strCreDate = "NULL"
                End If

                '-------------------------- วันที่เปิดสั่งซื้้อ -----------------------------------------------

                If txtPrdate.Text <> "__/__/____" Then
                    strPrdate = Mid(txtPrdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtPrdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtPrdate.Text.ToString, 1, 2)
                    strPrdate = "'" & SaveChangeEngYear(strPrdate) & "'"

                Else
                    strPrdate = "NULL"
                End If

                '-------------------------- วันที่นัดเข้า -----------------------------------------------

                If txtFCdate.Text <> "__/__/____" Then
                    strFcDate = Mid(txtFCdate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtFCdate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtFCdate.Text.ToString, 1, 2)
                    strFcDate = "'" & SaveChangeEngYear(strFcDate) & "'"

                Else
                    strFcDate = "NULL"
                End If


                '-------------------------- วันที่รับเข้า -----------------------------------------------

                If txtIndate.Text <> "__/__/____" Then
                    strIndate = Mid(txtIndate.Text.ToString, 7, 4) & "-" _
                               & Mid(txtIndate.Text.ToString, 4, 2) & "-" _
                               & Mid(txtIndate.Text.ToString, 1, 2)
                    strIndate = "'" & SaveChangeEngYear(strIndate) & "'"

                Else
                    strIndate = "NULL"
                End If

                '----------------------- กำหนดกลุ่มของชิ้นงาน  -------------------------------

                Select Case cboPart.Text.ToString.Trim

                    Case Is = "พื้นล่าง"
                        strPartType = "SOLE2"
                    Case Is = "พื้นบน"
                        strPartType = "SOLE1"
                    Case Is = "ใส้พื้นบน"
                        strPartType = "FILL"
                    Case Is = "โลโก้ส้น"
                        strPartType = "LOGOHEEL"
                    Case Is = "โลโก้พื้น"
                        strPartType = "LOGO"
                    Case Is = "EVA ติดส้น"
                        strPartType = "EVAH"
                    Case Is = "EVA บนหนังหน้า"
                        strPartType = "EVAF"
                    Case Is = "หนังหน้า"
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
           MsgBox("พบข้อผิดพลาดขณะทำการบันทึก โปรดดำเนินการใหม่อีกครั้ง", MsgBoxStyle.Critical, "ผิดพลาด")
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
                x = Len(txtThk.Text.ToString)   'นับตัวอักษรใน txtThk

                For i = 1 To x
                    strTmp = Mid(txtThk.Text, i, 1)  'ต้ดสตริงเริ่มจากตัวแรกไปสุดข้อความ

                    Select Case strTmp     'เช็คค่า strTmp
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
                Case Is = 35   'ปุ่ม End
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTdesc.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionStart = 0 Then
                        txtTtrait.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length   'ตำแน่ง ChkPoint = ความยาวตัวอักษร
                        If .SelectionStart = intChkPoint Then
                            txtTtrait.Focus()
                        End If

                    End If

                Case Is = 40  'ปุ่มลง
                    txtPId.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Select Case strMerg.ToString.Length    ' Check ความยาวสตริง
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTtime2.Focus()
                    End If
                Case Is = 38 'ลูกศรขึ้น
                    txtTeva.Focus()
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                        txtPId.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtPId.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtPId.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtPr.Focus()
                    End If
                Case Is = 38   'ปุ่มลูกศรขึ้น
                    txtPrice.Focus()
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtFCdate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtFCdate.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง
                    txtPId.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtFCdate.Focus()
                    End If
                Case Is = 38   'ปุ่มลูกศรขึ้น
                    txtPrdate.Focus()
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtSupplier.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtSupplier.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง
                    txtRmk.Focus()
                Case Is = 113 'ปุ่ม F2
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

                'If Year(z) < 2500 Then  'กรณีกรอกเป็น ค.ศ. จะเเปลงเป็น พ.ศ. ทันที
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtPr.Focus()
                    End If
                Case Is = 38   'ปุ่มลูกศรขึ้น
                    txtPId.Focus()
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtLg.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtLg.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง
                    txtRmk.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtw.Focus()
                    End If
                Case Is = 38   'ปุ่มลูกศรขึ้น
                    txtFCdate.Focus()
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtPrice.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtPrice.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง
                    txtRmk.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtSizeDesc.Focus()
                    End If
                Case Is = 38   'ปุ่มลูกศรขึ้น

                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtSizeQty.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     'ให้ InChkPoint = ความยาวของ  mskLdate

                        If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtSizeQty.Focus()
                        End If

                    End If
                Case Is = 40 'ปุ่มลง
                    txtPrice.Focus()

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

         End Select
        End With
    End Sub

    Private Sub mskSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSetQty.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)  'ถ้า Enter
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtw.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น
                    txtPr.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                        txtIndate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtIndate.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง           
                    txtSupplier.Focus()
                Case Is = 113 'ปุ่ม F2
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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtCdate.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น  
                    txtTwarm.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtSize.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtSize.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtw.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
    End Sub

Private Sub txtPId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPId.KeyPress

  Select Case Asc(e.KeyChar)

         Case 48 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
             e.Handled = False

         Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

         Case 13   'Enter = 13
              e.Handled = False
              txtSize.Focus()

         Case Else
             e.Handled = True
             MsgBox("กรุณาระบุข้อมูลเป็นภาษาอังกฤษหรือตัวเลข", MsgBoxStyle.Critical, "ผิดพลาด")
  End Select

End Sub

Private Sub btnEditEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp1.Click

  Dim OpenFileDialog1 As New OpenFileDialog
  Dim strFileFullPath As String   'เก็บพาร์ทไฟล์
  Dim strFileName As String       'เก็บ filenaem
  Dim img As Image = Nothing

  Dim dateNow As Date = Now
  Dim typePic As String
  Dim strNamePic As String
  Dim lengPic, lengTypePic As Integer

      With OpenFileDialog1
           .CheckFileExists = True
           .ShowReadOnly = False
           .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png"
           .FilterIndex = 2

           Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    ' Load ไฟล์ใส่ picturebox
                    strFileName = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์

                    img = ScaleImage(Image.FromFile(.FileName), picEqp1.Height, picEqp1.Width)
                    picEqp1.Image = img

                    '----------- เพิ่มวันที่เวลาต่อชื่อไฟล์ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4) ' ตัดเอา .jpg .png .gif 
                    lengPic = strFileName.Length - 4   'เอาจำนวน charactor ทั้งหมดมาลบออกด้วยสกุล picture
                    strNamePic = Mid(strFileName, 1, lengPic)     'ตัดเอาเฉพาะชื่อรูป
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           'เพิ่มวันที่ต่อท้ายชื่อไฟล์รูป

                    lblPicPath1.Text = strFileFullPath
                    lblPicName1.Text = strFileName

                End If

            Catch ex As Exception
                  ClearBlankPicture1()
            End Try

        End With
        CallEditData()   'สั่งเรียกข้อมูลทางเทคนิคขึ้นมารอ 
    End Sub

Private Sub btnEditEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp2.Click
 Dim OpenFileDialog2 As New OpenFileDialog
 Dim strFileFullPath As String   'เก็บพาร์ทไฟล์
 Dim strFileName As String       'เก็บ filenaem
 Dim img As Image = Nothing

 Dim dateNow As Date = Now
 Dim typePic As String
 Dim strNamePic As String
 Dim lengPic, lengTypePic As Integer

        With OpenFileDialog2
            .CheckFileExists = True
            .ShowReadOnly = False
            .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png"
            .FilterIndex = 2

            Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    ' Load the specified file into a PictureBox control.
                    strFileName = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์

                    img = ScaleImage(Image.FromFile(.FileName), picEqp2.Height, picEqp2.Width)
                    picEqp2.Image = img

                    '----------- เพิ่มวันที่เวลาต่อชื่อไฟล์ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4) ' ตัดเอา .jpg .png .gif 
                    lengPic = strFileName.Length - 4   'เอาจำนวน charactor ทั้งหมดมาลบออกด้วยสกุล picture
                    strNamePic = Mid(strFileName, 1, lengPic)     'ตัดเอาเฉพาะชื่อรูป
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           'เพิ่มวันที่ต่อท้ายชื่อไฟล์รูป

                    lblPicPath2.Text = strFileFullPath
                    lblPicName2.Text = strFileName 'นี้ก็ใช้ได้แต่ประกาศตัวแปรเยอะกว่า

                End If

            Catch ex As Exception
                ClearBlankPicture2()
            End Try

        End With
         CallEditData()   'สั่งเรียกข้อมูลทางเทคนิคขึ้นมารอ 
    End Sub

 Private Sub btnEditEqp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp3.Click

  Dim OpenFileDialog3 As New OpenFileDialog
  Dim strFileFullPath As String   'เก็บพาร์ทไฟล์
  Dim strFileName As String       'เก็บ filenaem
  Dim img As Image = Nothing

  Dim dateNow As Date = Now
  Dim typePic As String
  Dim strNamePic As String
  Dim lengPic, lengTypePic As Integer

      With OpenFileDialog3
           .CheckFileExists = True
           .ShowReadOnly = False
           .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png"
           .FilterIndex = 2

            Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    ' Load the specified file into a PictureBox control.
                    strFileName = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์

                    img = ScaleImage(Image.FromFile(.FileName), picEqp3.Height, picEqp3.Width)
                    picEqp3.Image = img

                    '----------- เพิ่มวันที่เวลาต่อชื่อไฟล์ ------------
                    strFileName = Trim(strFileName)
                    lengTypePic = strFileName.Length - 4
                    typePic = Mid(strFileName, lengTypePic + 1, 4)        ' ตัดเอา .jpg .png .gif 
                    lengPic = strFileName.Length - 4                      'เอาจำนวน charactor ทั้งหมดมาลบออกด้วยสกุล picture
                    strNamePic = Mid(strFileName, 1, lengPic)             'ตัดเอาเฉพาะชื่อรูป
                    strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           'เพิ่มวันที่ต่อท้ายชื่อไฟล์รูป

                    lblPicPath3.Text = strFileFullPath
                    lblPicName3.Text = strFileName                    'นี้ก็ใช้ได้แต่ประกาศตัวแปรเยอะกว่า

                End If

            Catch ex As Exception
                ClearBlankPicture3()
            End Try

        End With
         CallEditData()   'สั่งเรียกข้อมูลทางเทคนิคขึ้นมารอ 
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
                            If InStr("0123456789.", strTmp) > 0 Then    'ค้นหาสตริง
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtSetQty.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น  

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtw.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkpoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtw.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtPr.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

    Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress
        Select Case e.KeyChar
            Case Is = Chr(13)
                txtw.Focus()
            Case Is = Chr(46)   'เครื่องหมายจุลภาค(.) 
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
                mskSizeQty.Text = ""     'เคลียร์ mskSizeQty
                z = CDbl(strMerg)        'แปลง Type dbl
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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTeva.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น     
                    txtTsht.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtTtime1.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtTtime1.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtPId.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTwarm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTwarm.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 13
                e.Handled = False
                txtTtime1.Focus()
            Case 8                  'ปุ่ม Backspace
                e.Handled = False
            Case 32                 'เคาะ spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Select
    End Sub

    Private Sub txtTtime1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTtime1.KeyDown
        Dim intChkPoint As Integer
        With txtTtime1
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTwarm.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น  
                    txtTsht.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtTtime2.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtTtime2.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtPId.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTtime1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTtime1.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 13
                e.Handled = False
                txtTtime2.Focus()
            Case 8                  'ปุ่ม Backspace
                e.Handled = False
            Case 32                   'เคาะ spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Select
    End Sub

    Private Sub txtTtime2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTtime2.KeyDown
        Dim intChkPoint As Integer
        With txtTtime2
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTtime1.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น  
                    txtTeva.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtCdate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtCdate.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtPId.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTtime2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTtime2.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 13
                e.Handled = False
                txtCdate.Focus()
            Case 8                  'ปุ่ม Backspace
                e.Handled = False
            Case 32                  ' เคาะ spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Select
    End Sub

    Private Sub txtEqpnm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqpnm.KeyDown
        Dim intChkPoint As Integer
        With txtEqpnm
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtEqp_id.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น        
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtShoe.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtShoe.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                     cboPart.DroppedDown = True
                     cboPart.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

    End Sub

    Private Sub txtShoe_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtShoe.KeyDown
        Dim intChkPoint As Integer
        With txtShoe
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtEqpnm.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น        
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtOrder.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtOrder.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                     txtOrder.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

    End Sub

    Private Sub txtOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrder.KeyDown
        Dim intChkPoint As Integer
        With txtOrder
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtShoe.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น     
                    txtEqp_id.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        cboPart.DroppedDown = True
                        cboPart.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            cboPart.DroppedDown = True
                            cboPart.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                     txtRemark.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
        Dim intChkPoint As Integer
        With txtRemark
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        cboPart.DroppedDown = True
                        cboPart.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น     
                    txtShoe.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        btnSaveData.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            btnSaveData.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง              
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
    End Sub

    Private Sub txtTdesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTdesc.KeyDown
        Dim intChkPoint As Integer
        With txtTdesc
            Select Case e.KeyCode
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น                       
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtThk.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtThk.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง      
                    txtTsht.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
    End Sub

    Private Sub txtTtrait_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTtrait.KeyDown
        Dim intChkPoint As Integer
        With txtTtrait
            Select Case e.KeyCode
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtThk.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น                       
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtTsht.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtTsht.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง      
                    txtTeva.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        cboPart.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtShoe.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtSet.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtSet.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง      
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtAmount.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtOrder.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtRemark.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtRemark.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง      
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTdesc.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtTtrait.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtTtrait.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtTeva.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTtrait.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtTdesc.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtTeva.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtTeva.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtTwarm.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtTsht.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtTtrait.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtTwarm.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtTwarm.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtCdate.Focus()
                Case Is = 113 'ปุ่ม F2
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
             Case Is = 35 'ปุ่ม End 
             Case Is = 36 'ปุ่ม Home
             Case Is = 37 'ลูกศรซ้าย
                  If .SelectionStart = 0 Then
                     txtPId.Focus()
                  End If
             Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtCdate.Focus()
             Case Is = 39 'ปุ่มลูกศรขวา
                   If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                       txtSizeDesc.Focus()   ''
                   Else
                       intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtSizeDesc.Focus()
                        End If
                    End If
             Case Is = 40  'ปุ่มลง  
                    txtPrice.Focus()
             Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress

   Select Case Asc(e.KeyChar)
          Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข 0 57คือเลข 9ตามลำดับ
                e.Handled = False

          Case 8, 46, 45  ' ปุ่ม Backspace = 8, ปุ่มDelete = 46 , ปุ่ม 45 = ขีดกลาง
                e.Handled = False

          Case 13 'ปุ่ม Enter = 13
               e.Handled = False
               txtSizeDesc.Focus()

          Case Else
                e.Handled = True
                MsgBox("กรุณาระบุข้อมูลเป็นภาษาอังกฤษหรือตัวเลข", MsgBoxStyle.Critical, "ผิดพลาด")

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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtSizeQty.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtSize.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtPrdate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtPrdate.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtFCdate.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

 End Sub

Private Sub txtPr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPr.KeyPress

Select Case Asc(e.KeyChar)

       Case 45 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
             e.Handled = False

       Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

       Case 13   'Enter = 13
              e.Handled = False
              txtPrdate.Focus()

      Case Else
             e.Handled = True
             MessageBox.Show("กรุณาระบุข้อมูลเป็นภาษาอังกฤษ")
  End Select

End Sub

Private Sub txtPrdate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPrdate.KeyDown
 Dim intChkPoint As Integer
     With txtPrdate

         Select Case e.KeyCode
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtPr.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtSizeQty.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtFCdate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtFCdate.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtFCdate.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtPrdate.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtPr.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtIndate.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtIndate.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtSupplier.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtFCdate.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtPrdate.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtSupplier.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtSupplier.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtSupplier.Focus()
                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtIndate.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtFCdate.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtMouth_mold.Focus()   ''
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtMouth_mold.Focus()
                        End If
                    End If
                Case Is = 40  'ปุ่มลง  
                    txtRmk.Focus()

                Case Is = 113 'ปุ่ม F2
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

                Select Case strMerge.IndexOf(".") 'หาตำแหน่งที่พบเป็นครั้งแรก

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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                    If .SelectionStart = 0 Then
                        txtSizeQty.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtPId.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                        txtPr.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtPr.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                    txtFCdate.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
    End Sub

    Private Sub mskPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrice.KeyPress
        Select Case e.KeyChar
               Case Is = Chr(13)  'ถ้า Enter
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
            Case Is = Chr(13)  'ถ้า Enter
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
     ShowResvrd()       'แสดง gpbSeek 
     CallEditData()    'ซับรูทีนแสดง Size เพื่อแก้ไขข้อมูล
     CallEditData2()   'ซับรูทีนแสดงข้อมูลทางเทคนิค

     gpbSeek.Text = "แก้ไขข้อมูล"
     txtPId.ReadOnly = True
     txtSize.ReadOnly = True
     txtSizeDesc.ReadOnly = True
  Else
      MsgBox("ไม่มีรายการ SIZE ที่ต้องการแก้ไข!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
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

Dim strWd As String = ""   'เก็บค่ากว้าง
Dim strLg As String = ""   'เก็บความยาว
Dim strHg As String = ""   'เก็บความสูง

If dgvSize.Rows.Count <> 0 Then

   Dim strCode As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(2).Value.ToString.Trim     'เก็บ Size
   Dim strLot As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(4).Value.ToString.Trim      'รหัสแผง
   Dim strGpsize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(5).Value.ToString.Trim    'กรุ๊ฟ size
 
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
            .ActiveConnection = Nothing    'เคลียร์การเชื่อมต่อ
            .Close()

    End With
    Rsd = Nothing

Conn.Close()
Conn = Nothing
End If

End Sub

'----------------------- ฟังก์ชั่นคืนค่า Dimension -------------------------------------------------------------
Private Function RetrnDiams(ByVal strDia As String, ByRef strW As String, ByRef strL As String) As Boolean
Dim i, x As Integer
Dim strTmp As String = ""
Dim strMerg As String = ""

Dim strDiamns(1) As String  'เก็บสตริงเป็นอะเรย์
Dim y As Integer = 0

                 x = Len(strDia)
                 For i = 1 To x
                         strTmp = Mid(strDia, i, 1)
                         Select Case strTmp

                                Case Is = "x"       'ถ้าเป็นเครื่องหมายคูณ
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
             strGpsize = .Rows(.CurrentRow.Index).Cells(5).Value.ToString       'กรุ๊ฟ size


             If strSizeAct <> "" Then

                    btyConsider = MsgBox("SIZE : " & strSize.ToString.Trim & vbNewLine _
                                                   & "รหัสเเผง : " & strSizeDesc.ToString.Trim & vbNewLine _
                                                   & "กรุ๊ปไซต์ : " & strGpsize.ToString.Trim & vbNewLine _
                                                   & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
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

                             '------------------ ลบข้อมูลในตาราง tmp_eqptrn_newsize -------------------

                              strCmd = "DELETE FROM tmp_eqptrn_newsize" _
                                                 & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                 & " AND size_id ='" & strSize.ToString.Trim & "'" _
                                                 & " AND size_desc = '" & strSizeDesc.ToString.Trim & "'"
                              Conn.Execute(strCmd)

                              Conn.Close()
                              Conn = Nothing

                              .Rows.RemoveAt(.CurrentRow.Index)  'ลบข้อมูล Cell ปัจจุบัน
                              ShowScrapItem()

                           End If
                   Else
                      .Focus()
                   End If

     Else
               MsgBox("ไม่มีรายการ SIZE ที่ต้องการลบข้อมูล!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
               dgvSize.Focus()

     End If
    End With
End Sub

Private Sub btnDelEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp1.Click
  ClearBlankPicture1()
  CallEditData()   'สั่งเรียกข้อมูลทางเทคนิคขึ้นมารอ 
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
  CallEditData()   'สั่งเรียกข้อมูลทางเทคนิคขึ้นมารอ 
End Sub

Private Sub btnDelEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp2.Click
  ClearBlankPicture2()
  CallEditData()   'สั่งเรียกข้อมูลทางเทคนิคขึ้นมารอ 
End Sub

Private Sub txtPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrice.KeyPress
  Select Case e.KeyChar
         Case Is = Chr(13)  'ถ้า Enter
                   txtPr.Focus()
  End Select
End Sub

Private Sub txtRmk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRmk.KeyDown
    Dim intChkPoint As Integer
        With txtRmk
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                    If .SelectionStart = 0 Then
                        txtMouth_mold.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น    
                    txtIndate.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                        btnSeekSave.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            btnSeekSave.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSizeDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSizeDesc.KeyDown
 Dim intChkPoint As Integer
        With txtSizeDesc
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                    If .SelectionStart = 0 Then
                        txtSize.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น    

                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                        txtSetQty.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtSetQty.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                          txtPrice.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSizeDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSizeDesc.KeyPress

  Select Case Asc(e.KeyChar)

         Case 48 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
             e.Handled = False

         Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

         Case 13   'Enter = 13
              e.Handled = False
             txtSetQty.Focus()

         Case Else
             e.Handled = True
             MessageBox.Show("กรุณาระบุข้อมูลเป็นภาษาอังกฤษ")
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

'-------------------------------- ซับรูทีนเพิ่มสตริง Eqpid -----------------------------------------------
Private Sub AddtextEqpid()

  Select Case cboPart.Text.ToString.Trim
        Case Is = "พื้นบน"
              txtEqp_id.Text = txtEqp_id.Text & "-A"
        Case Is = "ใส้พื้นบน"
              txtEqp_id.Text = txtEqp_id.Text & "-B"
        Case Is = "โลโก้ส้น"
              txtEqp_id.Text = txtEqp_id.Text & "-C"
        Case Is = "พื้นล่าง"
              txtEqp_id.Text = txtEqp_id.Text & "-D"
        Case Is = "EVA ติดส้น"
              txtEqp_id.Text = txtEqp_id.Text & "-E"
        Case Is = "โลโก้พื้น"
              txtEqp_id.Text = txtEqp_id.Text & "-F"
        Case Is = "หนังหน้า"
              txtEqp_id.Text = txtEqp_id.Text & "-G"
        Case Is = "EVA บนหนังหน้า"
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
         Case Is = Chr(13)  'ถ้า Enter
                   txtRemark.Focus()
  End Select

End Sub

Private Sub cboPart_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPart.LostFocus
  Dim strEqpid As String
      strEqpid = txtEqp_id.Text.Trim

         '--------------- ค้นหาสตริง "-" ว่ามีหรือไม่ ----------------------

      If InStr(1, strEqpid, "-") > 0 Then
         txtEqp_id.Text = strEqpid
      Else
             AddtextEqpid()  'เพิ่มสตริงต่อท้าย eqp_id
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
       .Show("คลิกเพื่อดูรูปใหญ่", picEqp1)
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
       .Show("คลิกเพื่อดูรูปใหญ่", picEqp2)
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
       .Show("คลิกเพื่อดูรูปใหญ่", picEqp3)
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

               '----------------------- ล้างข้อมูลใน tmp_eqptrn_newsize ------------------------------

                 sqlCmd = "DELETE FROM tmp_eqptrn_newsize " _
                              & "WHERE user_id= '" & frmMainPro.lblLogin.Text.ToString.Trim & "'"

                 Conn.Execute(sqlCmd)

               ' ---------- วนลูปจัดเรียง size ใหม่ --------------

               Do While Not .EOF

                  SearchWithinThis = .Fields("size_id").Value.ToString.Trim
                  If SearchWithinThis.IndexOf("-") <> -1 Then          'หาก size ต้นฉบับไม่มี size รว่ม (x-xx)
                     strArr = SearchWithinThis.Split("-")              'อ่านค่า size เก็บไว้ในตัวเเปร
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

                  '----------------------- Insert ข้อมูลลงในตารางใหม่หลังเรียง size ใหม่ ----------------------

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

     sta(0) = "------ โปรดเลือกสถานะ ------"
     sta(1) = "ปกติ / มีการเคลื่อนไหว"
     sta(2) = "รอใช้งาน / ไม่มีการเคลื่อนไหว"
     sta(3) = "ยกเลิกการใช้งาน"

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
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                  txtSupplier.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtRmk.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtRmk.Focus()
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง               
                                txtRmk.Focus()

                      Case Is = 113 'ปุ่ม F2
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