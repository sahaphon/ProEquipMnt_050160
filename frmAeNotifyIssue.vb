Imports ADODB
Imports System.IO
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.IO.MemoryStream

Public Class frmAeNotifyIssue
 Dim IsShowSeek As Boolean        'ตัวเเปรแสดงสถานะ gpbSeek
 Dim strDateDefault As String     'ตัวแปรสำหรับวันที่ทั่วไป

 Public Const DrvName As String = "\\10.32.0.15\data1\EquipPicture\"
 Public Const PthName As String = "\\10.32.0.15\data1\EquipPicture"         'ตัวแปรสำหรับเก็บ part รูปภาพ

Protected Overrides ReadOnly Property CreateParams() As CreateParams       'ป้องกันการปิดโดยใช้ปุ่ม Close Button(ปุ่มกากบาท)

    Get
            Dim cp As CreateParams = MyBase.CreateParams
            Const CS_DBLCLKS As Int32 = &H8
            Const CS_NOCLOSE As Int32 = &H200
            cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
            Return cp
    End Get

End Property

Private Sub frmAeNotifyIssue_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
 ClearTmpTable(0, "")  'ลบข้อมูล Table tmp_eqptrn where user_id..
 frmNotifyIssue.lblCmd.Text = "0"  'เคลียร์สถานะ
 Me.Dispose()     'ทำลายฟอร์ม คืนหน่วยความจำ

End Sub

Private Sub frmAeReqFxeqp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Dim dtComputer As Date = Now       'ตัวแปรเก็บค่าวันที่ปัจจุบัน
  Dim strCurrentDate As String       'เก็บค่าสตริงวันที่ปัจจุบัน

  Me.WindowState = FormWindowState.Maximized  'ให้ฟอร์มขยายเต็มหน้าจอ
  StdDateTimeThai()        'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module
  strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

  ClearAlldata()
  PreDeptSeek()            'โหลดรายชื่อการแผนก
  PreGroupSeek()           'โหลดกลุ่มอุปกรณ์


     Select Case frmNotifyIssue.lblCmd.Text.ToString

            Case Is = "0" 'เพิ่มข้อมูล

                With txtBegin
                     .Text = strCurrentDate
                     strDateDefault = strCurrentDate
                End With

                GenDocid()  'รันเลขที่เอกสาร
                gpbReceive.Enabled = False
                txtDocid.ReadOnly = True


             Case Is = "1" 'แก้ไขข้อมูล

               ' InputTmpData()  'Copy ข้อมูลลง Tmpdata

                         If frmMainPro.lblLogin.Text.Trim = "SUTID" Then

                                      If chkFirstApprove() Then     'ตรวจสอบว่า ผจก.แผนกที่แจ้งเซ็นอนุมัติแล้ว
                                         LockEditData()

                                         gpbNotify.Enabled = False   'ปิดgroupbox ส่วนผู้แจ้ง
                                         gpbReceive.Enabled = True

                                      Else

                                         LockEditData()

                                         gpbNotify.Enabled = False   'ปิดgroupbox ส่วนผู้แจ้ง
                                         gpbReceive.Enabled = False
                                         btnSaveData.Enabled = False

                                         MessageBox.Show("เอกสารยังไม่ได้อนุมัติจากแผนกที่แจ้งปัญหา โปรดติดต่อแผนกที่แจ้ง!...", _
                                                                     "ไม่สามารถดำเนินการได้ ", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                                      End If

                          Else
                                LockEditData()

                                gpbNotify.Enabled = True
                                gpbReceive.Enabled = False
                                txtDocid.ReadOnly = True

                          End If


             Case Is = "2"   'มุมมองข้อมูล

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

            .ActiveConnection = Nothing   'เคลียร์ Connection
            .Close()

        End With
        Rsd = Nothing   'เคลียร์ RecordSet
  Conn.Close()    'ปิดการเชื่อมต่อ
  Conn = Nothing   'เคลียร์ RecordSet

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

                   Case Is = "0"  'ลบข้อมูลหลังปิดฟอร์ม

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

  Dim strCmd As String  ' เก็บสตริง Command

  Dim strLoadFilePicture As String   'เก็บค่าสตริงโหลด Picture
  Dim strPathPicture As String = "H:\EquipPicture\"   'เก็บ part

  Dim strSqlSelc As String = ""   'เก็บสตริง sql select
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
                         strGpType = "โมล์ดฉีด EVA INJECTION"
                    Case Is = "B"
                         strGpType = "โมล์ดฉีด PVC INJECTION"
                    Case Is = "C"
                         strGpType = "โมล์ดหยอด PU"
                    Case Is = "D"
                         strGpType = "โมล์ดแผงอัดลายหนังหน้า,พื้น"
                    Case Is = "E"
                         strGpType = "มีดตัด"
                    Case Is = "F"
                         strGpType = "บล็อกสกรีน"
                    Case Is = "G"
                         strGpType = "บล็อกอาร์ค"

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

                '------------------------------- Load รูปภาพอุปกรณ์ -----------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_issue").Value.ToString.Trim
                ' เช็ดพาร์ทไฟล์มีอยู่จริง
                If File.Exists(strLoadFilePicture) Then

                   Dim img1 As Image      'ประกาศตัวแปร img1 เพื่อเก็บภาพ
                   img1 = Image.FromFile(strLoadFilePicture)  'img1 เท่ากับpicture ที่โหลดมาจาก db
                   Dim s1 As String = ImageToBase64(img1, System.Drawing.Imaging.ImageFormat.Jpeg)  'ประกาศตัวแปร s1 เก็บค่าสตริงที่แปลงแล้ว
                   img1.Dispose()     'ทำลายตัวแปร img1
                   Piceqp.Image = Base64ToImage(s1)

                Else
                    Piceqp.Image = Nothing  'ถ้าไม่มีรูปภาพให้ picEqp1 ว่างเปล่า

                End If


                                  strCmd = frmNotifyIssue.lblCmd.Text.ToString.Trim    'ให้ strCmd เท่ากับค่าใน lblcmd ในฟอร์ม frmEqpSheet

                                  Select Case strCmd

                                         Case Is = "1"   'ให้ล็อคตอนแก้ไข

                                         Case Is = "2"   'ให้ล็อคตอนมุมมอง
                                                btnSaveData.Enabled = False  'ปิดปุ่ม "บันทึกข้อมูล"

                                  End Select
   
        End If

            .ActiveConnection = Nothing   'สั่งปิดการเชื่อมต่อ
            .Close()

        End With

   Rsd = Nothing   'เคลียร์ค่า RecordSet
   Conn.Close()    'สั่งตัดการเชื่อมต่อ
   Conn = Nothing  'เคลี่ย์ Connection

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

    strDept(0) = "121000 แผนกผลิตโฟม"
    strDept(1) = "122000 แผนกตัดชิ้นส่วน"
    strDept(2) = "123000 แผนกเย็บ"
    strDept(3) = "124000 แผนกฉีด PVC"
    strDept(4) = "125000 แผนกฉีด EVA INJECTION"
    strDept(5) = "126000 แผนกฉีด PU"


    strDeptTo(0) = "แผนกขั้นตอนและอุปกรณ์การผลิต"

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

    strGroup(0) = "โมล์ดฉีด EVA INJECTION"
    strGroup(1) = "โมล์ดฉีด PVC"
    strGroup(2) = "โมล์ดหยอด PU"
    strGroup(3) = "โมล์ดแผงอัดลายหนังหน้า,พื้น"
    strGroup(4) = "มีดตัด"
    strGroup(5) = "บล็อกสกรีน"
    strGroup(6) = "บล็อกอาร์ค"

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
     Thayear = Mid(SaveChangeThaYear(strCurrentDate), 3, 2) 'ต้ดปีไทย 5X

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

             .MoveLast()              'เลื่อนไปยัง Record สุดท้าย

             LastNumber = CInt(Mid(.Fields("req_id").Value.ToString.Trim, 5))  'ตัดสตริง เอา 4 ต้วท้าย  000x
             LastYear = Mid(.Fields("req_id").Value.ToString.Trim, 3, 2)  'ตัดเอาปี  5x เฉพาะ 2ตัวแรก

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
   MsgBox("ทำ checkDatabfsave")
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

                                                                Case Is = "0"      'กรณีเพิ่มข้อมูล

                                                                     If CheckCodeDuplicate() Then
                                                                        MsgBox("ทำ SaveNewdata")
                                                                        SaveNewData()

                                                                     Else
                                                                         MessageBox.Show("กรุณาออกจากฟอร์ม แล้วดำเนินการใหม่!....", "***รหัสเอกสารซ้ำ***" _
                                                                            , MessageBoxButtons.OK, MessageBoxIcon.Error)

                                                                         ClearAlldata()    'ล้างหน้าจอ                                                            
                                                                         cboDepto.Focus()
                                                                      End If


                                                                 Case Is = "1"         'กรณีแก้ไขข้อมูล
                                                                       SaveEditData()


                                                            End Select

                                                         Else

                                                             MsgBox("โปรดระบุปัญหาที่พบ  " _
                                                                   & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                                             txtIssue.Focus()
                                                         End If

                                                Else
                                                    MsgBox("โปรดระบุ Size อุปกรณ์  " _
                                                      & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                                    txtSize.Focus()

                                                 End If

                                        Else
                                            MsgBox("โปรดเลือกกลุ่มอุปกรณ์  " _
                                                 & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                            cboGroup.DroppedDown = True
                                            cboGroup.Focus()

                                         End If

                                  Else
                                       MsgBox("โปรดระบุรุ่นอุปกรณ์  " _
                                           & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                       txtShoe.Focus()

                                  End If


                            Else

                                 MsgBox("โปรดเลือกแผนก / ฝ่ายผู้แจ้งปัญหา  " _
                                            & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                 cboDepfrom.DroppedDown = True
                                 cboDepfrom.Focus()
                           End If

                  Else
                       MsgBox("โปรดระบุชื่อผู้แจ้งปัญหา  " _
                                            & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                       txtName.Focus()

                  End If

            Else

                 MsgBox("โปรดระบุส่วนงาน  " _
                                    & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                 txtFrom.Focus()

            End If

    Else

        MsgBox("โปรดระบุแผนกรับเรื่องเเจ้งปัญหา  " _
                               & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
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
  Dim strDocdate As String           'เก็บสตริงวันที่เอกสาร
  Dim strGpType As String = ""       'เก็บประเภทอุปกรณ์
  Dim strPartType As String = ""     'เก็บชิ้นส่วนที่ผลิต
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

               'Conn.BeginTrans()      'จุดเริ่มต้น Transection

               strDate = dateSave.Date.ToString("yyyy-MM-dd")
               strDate = "'" & SaveChangeEngYear(strDate) & "'"

              '------------------------- วันที่เอกสาร ----------------------------------------------------

               strDocdate = Mid(txtBegin.Text.ToString.Trim, 7, 4) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 4, 2) & "-" _
                           & Mid(txtBegin.Text.ToString.Trim, 1, 2)
               strDocdate = "'" & SaveChangeEngYear(strDocdate) & "'"


              '---------------------------------------- วดป.ที่ผลิด --------------------------------------------

               If txtNeedDate.Text <> "__/__/____" Then

                  strNeedDate = Mid(txtNeedDate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtNeedDate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtNeedDate.Text.ToString, 1, 2)
                  strNeedDate = "'" & SaveChangeEngYear(strNeedDate) & "'"

               Else
                  strNeedDate = "NULL"

               End If


               '---------------------------------------- วดป.ที่ผลิด --------------------------------------------

               If txtWantDate.Text <> "__/__/____" Then

                  strWantDate = Mid(txtWantDate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtWantDate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtWantDate.Text.ToString, 1, 2)
                  strWantDate = "'" & SaveChangeEngYear(strWantDate) & "'"

               Else
                  strWantDate = "NULL"

               End If

                      '------------------------------------ บันทึกรูปอุปกรณ์  ------------------------------------

                       blnReturnCopyPic = CallCopyPicture(lblPicPath.Text.ToString.Trim, lblPicName.Text.ToString.Trim)

                       If blnReturnCopyPic Then
                          lblPicPath.Text = PthName

                       Else
                          lblPicName.Text = ""
                          lblPicPath.Text = ""
                          Piceqp.Image = Nothing

                       End If


                           Select Case cboGroup.Text.ToString.Trim

                                  Case Is = "โมล์ดฉีด EVA INJECTION"
                                       strGpType = "A"
                                  Case Is = "โมล์ดฉีด PVC INJECTION"
                                       strGpType = "B"
                                  Case Is = "โมล์ดหยอด PU"
                                       strGpType = "C"
                                  Case Is = "โมล์ดแผงอัดลายหนังหน้า,พื้น"
                                       strGpType = "D"
                                  Case Is = "มีดตัด"
                                       strGpType = "E"
                                  Case Is = "บล็อกสกรีน"
                                       strGpType = "F"
                                  Case Is = "บล็อกอาร์ค"
                                       strGpType = "G"

                          End Select

                            If frmMainPro.lblLogin.Text = "MALIWAN" Or frmMainPro.lblLogin.Text = "SUTID" Then       'ตรวจสอบ user login

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
                              'Conn.CommitTrans()  'สั่ง Commit transection

                            End If

        lblComplete.Text = txtDocid.Text.ToString.Trim  'บ่งบอกว่าบันทึกข้อมูลสำเร็จ

        Me.Hide()
        frmMainPro.Show()
        frmNotifyIssue.Show()

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub SaveNewData()
 MsgBox("ทำ Sub SaveNewData")
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim dateSave As Date = Now()    'เก็บค่าวันที่ปัจจุบัน
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
                strDate = SaveChangeEngYear(strDate)               'เเปลงเป็นปี ค.ศ.

                strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                 & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                 & Mid(txtBegin.Text.Trim, 1, 2)
                strDateDoc = "'" & SaveChangeEngYear(strDateDoc) & "'"

                '---------------------------------------- บันทึกรูปอุปกรณ์ ----------------------------------------------

                blnRetuneCopyPic = CallCopyPicture(lblPicPath.Text.Trim, lblPicName.Text.Trim)

                    If blnRetuneCopyPic Then          'ถ้า CallCopyPicture = true
                       lblPicPath.Text = PthName
                    Else
                       lblPicPath.Text = ""
                       lblPicName.Text = ""
                       picEqp = Nothing

                    End If


                    '---------------------------------------- วดป.ที่ต้องการ ----------------------------------------------------
                    If txtNeedDate.Text <> "__/__/____" Then

                       strNeedDate = Mid(txtNeedDate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtNeedDate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtNeedDate.Text.ToString, 1, 2)
                       strNeedDate = "'" & SaveChangeEngYear(strNeedDate) & "'"

                   Else
                       strNeedDate = "NULL"
                   End If


                    '---------------------------------------- วดป.กำหนดเสร็จ -------------------------------------------------
                   If txtWantDate.Text <> "__/__/____" Then

                      strWantDate = Mid(txtWantDate.Text.ToString, 7, 4) & "-" _
                                    & Mid(txtWantDate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtWantDate.Text.ToString, 1, 2)
                      strWantDate = "'" & SaveChangeEngYear(strWantDate) & "'"

                   Else
                      strWantDate = "NULL"
                   End If

                   '------------------------------------กำหนดกลุ่มของชิ้นงาน--------------------------------------------------

                   Select Case cboGroup.Text.ToString.Trim

                          Case Is = "โมล์ดฉีด EVA INJECTION"
                               strGpType = "A"
                          Case Is = "โมล์ดฉีด PVC INJECTION"
                               strGpType = "B"
                          Case Is = "โมล์ดหยอด PU"
                               strGpType = "C"
                          Case Is = "โมล์ดแผงอัดลายหนังหน้า,พื้น"
                               strGpType = "D"
                          Case Is = "มีดตัด"
                               strGpType = "E"
                          Case Is = "บล็อกสกรีน"
                               strGpType = "F"
                          Case Is = "บล็อกอาร์ค"
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


           lblComplete.Text = txtDocid.Text.ToString.Trim     'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
           Me.Hide()    'ซ่อมฟอร์มปัจจุบัน

           frmMainPro.Show()
           frmNotifyIssue.Show()

   Conn.Close()
   Conn = Nothing

End Sub

Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String) As Boolean

  Dim fname As String = String.Empty  'เท่ากับ ""
  Dim dFile As String = String.Empty
  Dim dFilePath As String = String.Empty

  Dim fServer As String = String.Empty
  Dim intResult As Integer   'คืนค่าเป็นจำนวนเต็ม

  On Error GoTo Err70

        fname = strPicPath & "\" & strPicName  'พาร์ท \\10.32.0.15\data1\EquipPicture\"ชื่อรูปภาพ"
        fServer = PthName & "\" & strPicName    'partServer \\10.32.0.15\data1\EquipPicture\"ชื่อรูปภาพ"

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
            .ActiveConnection = Nothing   'เคลียร์ Connection
            .Close()

        End With
        Rsd = Nothing   'เคลียร์ RecordSet
  Conn.Close()    'ปิดการเชื่อมต่อ
  Conn = Nothing   'เคลียร์ RecordSet

End Function

Private Sub btnEditEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp1.Click

 Dim OpenFileDialog1 As New OpenFileDialog
 Dim strFileFullPath As String   'เก็บพาร์ทไฟล์
 Dim strFileName As String       'เก็บ filenaem

     With OpenFileDialog1

            .CheckFileExists = True
            .ShowReadOnly = False
            .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png"
            .FilterIndex = 2

            Try

                If .ShowDialog = Windows.Forms.DialogResult.OK Then

                    ' Load ไฟล์ใส่ picturebox
                    Piceqp.Image = Image.FromFile(.FileName)

                    strFileName = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์
                    strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์

                    lblPicPath.Text = strFileFullPath
                    lblPicName.Text = strFileName             'นี้ก็ใช้ได้แต่ประกาศตัวแปรเยอะกว่า
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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                        txtFrom.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtFrom.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                      txtFrom.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSendTo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

  Select Case e.KeyChar
         Case Is = Chr(13)  'ถ้า Enter
               txtFrom.Focus()
  End Select

End Sub

Private Sub txtFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFrom.KeyDown

  Dim intChkPoint As Integer

      With txtFrom

         Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                      If .SelectionLength = .Text.Trim.Length Then
                           cboDepto.Focus()

                      End If

                Case Is = 38 'ปุ่มลูกศรขึ้น    
                         cboDepto.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                        txtName.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtName.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     cboDepfrom.DroppedDown = True
                     cboDepfrom.Focus()

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtFrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFrom.KeyPress

 Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
              txtName.Focus()
 End Select

End Sub

Private Sub cboDepfrom_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboDepfrom.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
               txtOrder.Focus()
  End Select

End Sub

Private Sub txtOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrder.KeyDown
 Dim intChkPoint As Integer

      With txtOrder

         Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                        If .SelectionLength = .Text.Trim.Length Then
                          cboDepfrom.DroppedDown = True

                        End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                        cboDepfrom.DroppedDown = True
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                       txtShoe.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                           txtShoe.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtShoe.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrder.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
             txtShoe.Focus()
  End Select

End Sub

Private Sub cboGroup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboGroup.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
              txtSize.Focus()
  End Select

End Sub

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown

  Dim intChkPoint As Integer

      With txtSize

         Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                        If .SelectionLength = .Text.Trim.Length Then
                          cboGroup.DroppedDown = True

                        End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                       txtShoe.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                       txtEqpnm.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtEqpnm.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtEqpnm.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtSize.Focus()

                       End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                      cboGroup.DroppedDown = True

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                       txtIssue.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtIssue.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtIssue.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSizeQty.KeyPress

     Select Case Asc(e.KeyChar)

            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 13
                e.Handled = False
                txtEqpnm.Focus()

            Case 8                  'ปุ่ม Backspace
                e.Handled = False
            Case 32                   'เคาะ spacebar
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
                            If InStr("0123456789.", strTmp) > 0 Then    'ค้นหาสตริง
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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtSize.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น  
                     cboGroup.DroppedDown = True
                     cboGroup.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtEqpnm.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkpoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtEqpnm.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtEqpnm.Focus()

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress
  Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
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
                mskSizeQty.Text = ""     'เคลียร์ mskSizeQty
                z = CDbl(strMerg)        'แปลง Type dbl
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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtSizeQty.Focus()

                       End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                      txtSize.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                       txtIssue.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtIssue.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtIssue.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtEqpnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqpnm.KeyPress

  Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
             txtIssue.Focus()

   End Select

End Sub

Private Sub txtIssue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtIssue.KeyDown

  Dim intChkPoint As Integer

      With txtIssue

         Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtEqpnm.Focus()

                       End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                      txtEqpnm.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                       txtCause.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtCause.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtCause.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select

        End With
End Sub

Private Sub txtIssue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIssue.KeyPress

  Select Case e.KeyChar
        Case Is = Chr(13)  'ถ้า Enter
             txtCause.Focus()

   End Select
End Sub

Private Sub txtCause_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCause.KeyDown

  Dim intChkPoint As Integer

      With txtCause

         Select Case e.KeyCode
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtIssue.Focus()

                       End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                      txtIssue.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                       txtNeedDate.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtNeedDate.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtNeedDate.Focus()

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

        End Select

     End With

End Sub

Private Sub txtCause_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCause.KeyPress
   Select Case e.KeyChar
          Case Is = Chr(13)  'ถ้า Enter
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

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtCause.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น
                    txtCause.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then
                        txtNeedtime.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtNeedtime.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง           
                    txtNeedtime.Focus()
                Case Is = 113 'ปุ่ม F2
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

                If Year(z) < 2500 Then 'ปีคริสต์ < 2100                        
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtNeedDate.Focus()

                       End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                      txtNeedDate.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                       txtRemark.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtRemark.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtRemark.Focus()

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

        End Select

     End With

End Sub

Private Sub txtNeedtime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtNeedtime.KeyPress

   Select Case Asc(e.KeyChar)

            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False

            Case 13
                e.Handled = False
                 txtRemark.Focus()

            Case 8                  'ปุ่ม Backspace
                e.Handled = False

            Case 32                 'เคาะ spacebar
                e.Handled = False

            Case 58
                e.Handled = False   'คือ :

            Case 46
                e.Handled = False   ' คือ .

            Case 44
                e.Handled = False   ' คือ ,

            Case Else
                e.Handled = True

                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)

     End Select

End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  On Error Resume Next    'หากเกิด error โปรแกรมยังจะทำงานต่อไปโดยไม่สนใจ error ที่เกิดขึ้น
  Dim strCode As String

     If MessageBox.Show("ต้องการออกจากฟอร์ม หรือไม่", "กรุณายืนยันออกจากฟอร์ม", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

           With frmNotifyIssue.dgvIssue
                If .Rows.Count > 0 Then   'ถ้ามีข้อมูลใน Grid
                    strCode = .Rows(.CurrentRow.Index).Cells(2).Value.ToString.Trim          'ให้strCode = ข้อมูลในแถวปัจจุบัน Cell แรก
                    lblComplete.Text = strCode  'ให้ label แสดงข้อมูลใน Cell ปัจจุบัน   
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtOrder.Focus()

                       End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                      txtOrder.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                       txtSizeQty.Focus()

                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtSizeQty.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtSizeQty.Focus()

                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย 
                       If .SelectionLength = .Text.Trim.Length Then
                          txtFrom.Focus()

                       End If
                Case Is = 38 'ปุ่มลูกศรขึ้น    
                      txtFrom.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา

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

                Case Is = 40 'ปุ่มลง    
                     cboDepfrom.DroppedDown = True
                     cboDepfrom.Focus()

                Case Is = 113 'ปุ่ม F2
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtFxIssue.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น  
                     txtFxIssue.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtWanttime.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkpoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtWanttime.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtWanttime.Focus()

                Case Is = 113 'ปุ่ม F2
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

                If Year(z) < 2500 Then 'ปีคริสต์ < 2100                        
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
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtWantDate.Focus()
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น  
                     txtFxIssue.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        btnSaveData.Focus()
                    Else
                        intChkpoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkpoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            btnSaveData.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    btnSaveData.Focus()

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length

            End Select
        End With
End Sub

Private Sub txtWanttime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWanttime.KeyPress
    Select Case Asc(e.KeyChar)

            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False

            Case 13
                e.Handled = False
                 btnSaveData.Focus()

            Case 8                  'ปุ่ม Backspace
                e.Handled = False

            Case 32                 'เคาะ spacebar
                e.Handled = False

            Case 58
                e.Handled = False   'คือ :

            Case 46
                e.Handled = False   ' คือ .

            Case 44
                e.Handled = False   ' คือ ,

            Case Else
                e.Handled = True

                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)

     End Select
End Sub

End Class