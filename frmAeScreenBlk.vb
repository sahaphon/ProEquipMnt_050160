Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeScreenBlk
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

Private Sub frmAeScreenBlk_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
   ClearTmpTable(0, "")  'ลบข้อมูล Table tmp_eqptrn where user_id..
   frmScreenBlk.lblCmd.Text = "0"  'เคลียร์สถานะ
   Me.Dispose()     'ทำลายฟอร์ม คืนหน่วยความจำ
End Sub

Private Sub frmAeScreenBlk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Dim dtComputer As Date = Now()
  Dim strCurrentDate As String

      StdDateTimeThai()        'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module
      strCurrentDate = dtComputer.Date.ToString("dd/MM/yyyy")

      ClearDataGpbHead()
      PrePartSeek()            'โหลดรายละเอียดใส่ใน cbo ชิ้้นส่วนที่ผลิต
      'txtEqp_id.Focus()

        Select Case frmScreenBlk.lblCmd.Text.ToString

               Case Is = "0" 'เพิ่มข้อมูล

                    With txtBegin                 'โหลดวันที่ปัจจุบันใส่ใน txtBegin
                         .Text = strCurrentDate
                         strDateDefault = strCurrentDate
                    End With

                    With Me
                         .Text = "เพิ่มข้อมูล"
                    End With

                '---------------------เวลาเพิ่มข้อมูลไม่ต้องแสดงสถานะ(ซ่อนคอลัมน์ใน Gridview)----------------------------

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

              Case Is = "2"   'มุมมองข้อมูล

                With Me
                     .Text = "มุมมองข้อมูล"
                End With

                LockEditData()
                txtEqp_id.ReadOnly = True  'ให้อ่านอย่างเดียว
                btnSaveData.Enabled = False

        End Select

    txtEqp_id.Focus()
End Sub

Private Sub LockEditData()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim Rsdwc As New ADODB.Recordset

  Dim strCmd As String  ' เก็บสตริง Command

  Dim blnHavedata As Boolean   'เก็บค่าตัวเเปร สำหรับเช็คว่ามีข้อมูลหรือไม่
  Dim strSqlSelc As String = ""   'เก็บสตริง sql select
  Dim strPart As String = ""
        'เก้บค่า Row ปัจจุบันในฟอร์ม frmScreenBlk
  Dim strCod As String = frmScreenBlk.dgvScreenBlk.Rows(frmScreenBlk.dgvScreenBlk.CurrentRow.Index).Cells(0).Value.ToString.Trim

        With Conn

            If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

        End With

        strSqlSelc = "SELECT * " _
                                    & "FROM v_moldinj_hd (NOLOCK)" _
                                    & " WHERE eqp_id = '" & strCod & "'"

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


                strCmd = frmScreenBlk.lblCmd.Text.ToString.Trim    'ให้ strCmd เท่ากับค่าใน lblcmd ในฟอร์ม frmEqpSheet

                Select Case strCmd
                    Case Is = "1"   'ให้ล็อคตอนแก้ไข
                    Case Is = "2"   'ให้ล็อคตอนมุมมอง
                        btnSaveData.Enabled = False  'ปิดปุ่ม "บันทึกข้อมูล"
                End Select

                '----------------- Insert ข้อมูลลงตาราง tmp_eqptrn โดยเอาข้อมูลจากตาราง  tmp_eqptrn มาด้วย----------------------

                strSqlSelc = "INSERT INTO tmp_eqptrn" _
                                  & " SELECT user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                  & " FROM eqptrn " _
                                  & " WHERE eqp_id = '" & strCod & "' "

                Conn.Execute(strSqlSelc)
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

Private Sub ClearDataGpbHead()
  txtEqp_id.Text = ""
  txtEqpnm.Text = ""
  txtShoe.Text = ""
  txtOrder.Text = ""
  txtAmount.Text = ""
  txtSet.Text = ""
End Sub

Private Sub PrePartSeek()
 Dim strGpTopic(8) As String
 Dim i As Integer

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

           Next

     End With

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
            End Select

     End With
 Conn.Close()
 Conn = Nothing

End Sub

Private Sub ShowScrapItem()                     'แสดงข้อมูลใน DataGridview 
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String             'เก็บค่า string command

 Dim dubQty As Double
 Dim dubAmt As Double

 Dim sngSetQty As Single                 'เก็บจำนวน SET

        With Conn
            If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

        End With

        strSqlCmdSelc = "SELECT * FROM v_tmp_eqptrn (NOLOCK)" _
                                 & "WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim.ToString & "' " _
                                 & "ORDER BY size_desc, size_id "

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlCmdSelc, Conn, , , )

            dgvSize.Rows.Clear()
            dgvSize.ScrollBars = ScrollBars.None                 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

            If .RecordCount <> 0 Then

                Do While Not .EOF()

                    dgvSize.Rows.Add( _
                                        IIf(.Fields("delvr_sta").Value.ToString.Trim = "1", My.Resources.accept, My.Resources._16x16_ledred), _
                                        "", _
                                        My.Resources.blank, _
                                        "", _
                                        .Fields("size_id").Value.ToString.Trim, _
                                        .Fields("size_act").Value.ToString.Trim, _
                                        .Fields("size_desc").Value.ToString.Trim, _
                                        .Fields("size_group").Value.ToString.Trim, _
                                        .Fields("backgup").Value.ToString.Trim, _
                                        Format(.Fields("set_qty").Value, "##0.0"), _
                                        Format(.Fields("size_qty").Value, "##0.0"), _
                                        .Fields("dimns").Value.ToString.Trim, _
                                        .Fields("price").Value, _
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

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
  CheckDataBfSave()
End Sub

Private Sub CheckDataBfSave()
 Dim IntListwc As Integer = dgvSize.Rows.Count
 Dim strProd As String = ""
 Dim strProdnm As String = ""

 Dim bytConSave As Byte  'เก็บค่า megbox 

     If txtEqp_id.Text <> "" Then

           If txtEqpnm.Text <> "" Then

                          If IntListwc > 0 Then

                             bytConSave = MsgBox("คุณต้องการบันทึกข้อมูลใช่หรือไม่!" _
                                  , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                    If bytConSave = 6 Then

                                             Select Case Me.Text

                                                    Case Is = "เพิ่มข้อมูล"

                                                         If CheckCodeDuplicate(txtEqp_id.Text) Then   'เช็ครหัสซ้ำ
                                                            SaveNewRecord()

                                                         Else
                                                             MessageBox.Show("รหัสอุปกรณ์ซ้ำ กรุณากรอกรหัสอุปกรณ์ใหม่!....", _
                                                                                  "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                             txtEqp_id.Text = ""
                                                             txtEqp_id.Focus()

                                                         End If

                                                     Case Else
                                                            SaveEditRecord()



                                             End Select

                                    Else
                                          dgvSize.Focus()
                                    End If

                          Else

                                If CheckCodeDuplicate(txtEqp_id.Text) Then           'ตรวจสอบรหัสอุปกรณ์ซ้ำ
                                   ShowResvrd()       'แสดงฟอร์มย่อย gpbSeek 
                                   gpbSeek.Text = "เพิ่มข้อมูล"
                                   txtSize.ReadOnly = False

                                Else

                                   MessageBox.Show("กรุณากรอกรหัสอุปกรณ์ใหม่!....", "รหัสอุปกรณ์ซ้ำ!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)

                                   txtEqp_id.Text = ""
                                   txtEqp_id.Focus()
                                End If

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

Private Sub SaveNewRecord()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String
  Dim dateSave As Date = Now()    'เก็บค่าวันที่ปัจจุบัน
  Dim strDate As String

  Dim strCredate As String
  Dim strDateDoc As String
  Dim strType As String = ""
  Dim Rsd As New ADODB.Recordset

      With Conn

           If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      Conn.BeginTrans()

             strDate = Date.Now.ToString("yyyy-MM-dd")
             strDate = SaveChangeEngYear(strDate)            'เเปลงเป็นปี ค.ศ.

             strDateDoc = Mid(txtBegin.Text.Trim, 7, 4) & "-" _
                                   & Mid(txtBegin.Text.Trim, 4, 2) & "-" _
                                   & Mid(txtBegin.Text.Trim, 1, 2)
                    strDateDoc = SaveChangeEngYear(strDateDoc)

                '---------------------------------------- วดป.ที่ผลิด ----------------------------------------------

                   If txtCdate.Text <> "__/__/____" Then

                       strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                     & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                     & Mid(txtCdate.Text.ToString, 1, 2)
                       strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

                   Else
                       strCredate = "NULL"
                   End If

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
                      & ",'" & "F" & "'" _
                      & ",'" & ReplaceQuote(txtEqp_id.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtEqpnm.Text.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                      & ",'" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & ReplaceQuote(txtRef.Text.ToString.Trim) & "'" _
                      & ",'" & ChangFormat(txtSet.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & " '" _
                      & ",'" & "" & " '" _
                      & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & " '" _
                      & ",'" & "" & " '" _
                      & "," & strCredate _
                      & ",'" & strDate & "'" _
                      & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                      & ",'" & ChangFormat(txtAmount.Text.ToString.Trim) & "'" _
                      & ",'" & RetrnAmount() & "'" _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ")"

                Conn.Execute(strSqlCmd)

               '------------------------------------------------บันทึกข้อมูลในตาราง eqptrn----------------------------------------------------------

                strSqlCmd = "INSERT INTO eqptrn " _
                                     & " SELECT [group] ='F'" _
                                     & ",eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'" _
                                     & ",size_id,size_desc,size_qty,weight,dimns,backgup,price,men_rmk" _
                                     & ",delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                     & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & " FROM tmp_eqptrn" _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                Conn.Execute(strSqlCmd)
                Conn.CommitTrans()

                frmScreenBlk.lblCmd.Text = txtEqp_id.Text.ToString.Trim   'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
                frmScreenBlk.Activating()
                Me.Close()

   Conn.Close()
   Conn = Nothing

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

               If txtCdate.Text <> "__/__/____" And txtCdate.Text <> "" Then

                  strCredate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                 & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                 & Mid(txtCdate.Text.ToString, 1, 2)
                  strCredate = "'" & SaveChangeEngYear(strCredate) & "'"

               Else
                  strCredate = "NULL"

               End If

                      '---------------------------------- UPDATE ข้อมูลในตาราง eqpmst ------------------------

                      strSqlCmd = "UPDATE  eqpmst SET eqp_name ='" & ReplaceQuote(txtEqpnm.Text.ToString.Trim) & "'" _
                                                & "," & "pi ='" & ReplaceQuote(txtOrder.Text.ToString.Trim) & "'" _
                                                & "," & "shoe ='" & ReplaceQuote(txtShoe.Text.ToString.Trim) & "'" _
                                                & "," & "part ='" & "" & "'" _
                                                & "," & "eqp_type ='" & " " & "'" _
                                                & "," & "set_qty =" & ChangFormat(txtSet.Text.ToString.Trim) _
                                                & "," & "pic_ctain ='" & "" & "'" _
                                                & "," & "pic_io ='" & "" & "'" _
                                                & "," & "pic_part ='" & "" & " '" _
                                                & "," & "remark ='" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                                & "," & "tech_desc = '" & "" & "'" _
                                                & "," & "tech_thk = '" & "" & "'" _
                                                & "," & "tech_lg = '" & "" & " '" _
                                                & "," & "tech_sht = '" & "" & "'" _
                                                & "," & "tech_eva = '" & "" & "'" _
                                                & "," & "tech_warm = '" & "" & "'" _
                                                & "," & "tech_time1 = '" & "" & "'" _
                                                & "," & "tech_time2 = '" & "" & " '" _
                                                & "," & "creat_date = " & strCredate _
                                                & "," & "eqp_amt = " & RetrnAmount() _
                                                & "," & "last_date = '" & strDate & "'" _
                                                & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                                & "," & "exp_id ='" & "" & "'" _
                                                & "," & "tech_trait ='" & "" & "'" _
                                                & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)

                      '------------------------------- ลบข้อมูลในตาราง eqptrn ---------------------------------------

                     strSqlCmd = "Delete FROM eqptrn" _
                                          & " WHERE eqp_id ='" & txtEqp_id.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlCmd)

                      '-------------------------------- บันทึกข้อมูลในตาราง eqptrn โดย Select จาก tmp_eqptrn ---------------

                     strSqlCmd = "INSERT INTO eqptrn " _
                                      & "SELECT [group] = 'F'" _
                                      & ",eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'" _
                                      & ",size_id,size_desc,size_qty,weight,dimns,backgup" _
                                      & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date" _
                                      & ",pr_doc,recv_date,ord_rep,ord_qty,fc_date,impt_id" _
                                      & ",sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                      & " FROM tmp_eqptrn " _
                                      & " WHERE user_id= '" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                    Conn.Execute(strSqlCmd)
                    Conn.CommitTrans()  'สั่ง Commit transection

                    frmScreenBlk.lblCmd.Text = txtEqp_id.Text.ToString.Trim   'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
                    frmScreenBlk.Activating()
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

Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String) As Boolean
 Dim fname As String = String.Empty
 Dim dFile As String = String.Empty
 Dim dFilePath As String = String.Empty

 Dim fServer As String = String.Empty
 Dim intResult As Integer

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

                '--------------------------- ถ้าค่าเป็น 0 แสดงว่าโหลดไฟล์ใช้อยู่ ไม่สามารถ Copy ไฟล์ได้ ------------------------------

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

Private Function CheckCodeDuplicate(ByVal strCod As String) As Boolean
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
                              & " WHERE eqp_id = '" & strCod & "'"


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

Private Sub ShowResvrd()  'ให้แสดง GroupBox gpbSeek ขึ้นมา

 tabMain.SelectedTab = tabSize
 IsShowSeek = Not IsShowSeek  'ถ้าสถานะแถบ seek ไม่แสดง ให้แสดง

   If IsShowSeek Then

       With gpbSeek
            .Visible = True
            .Left = 8    'แกน X
            .Top = 230   'แกน Y 252
            .Height = 500
            .Width = 1014
       End With

            StateLockFindDept(False) ' ล็อค FindDept โดยส่งค่าเป็น False ไป
   Else
            StateLockFindDept(True)

        End If
   End Sub

Private Sub StateLockFindDept(ByVal sta As Boolean)

 Dim strMode As String = frmScreenBlk.lblCmd.Text.ToString   'ตัวแปร lblCmd ฟอร์มบันทึกข้อมูลบล็อคอาร์ค ส่งค่ามาให้ strMode 
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

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  On Error Resume Next    'หากเกิด error โปรแกรมยังจะทำงานต่อไปโดยไม่สนใจ error ที่เกิดขึ้น
  Dim strCode As String

        If MessageBox.Show("ต้องการออกจากฟอร์ม หรือไม่", "กรุณายืนยันออกจากฟอร์ม", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

            With frmScreenBlk.dgvScreenBlk
                If .Rows.Count > 0 Then   'ถ้ามีข้อมูลใน Grid
                    strCode = .Rows(.CurrentRow.Index).Cells(0).Value.ToString.Trim          'ให้strCode = ข้อมูลในแถวปัจจุบัน Cell แรก
                    lblComplete.Text = strCode  'ให้ label แสดงข้อมูลใน Cell ปัจจุบัน   

                End If
            End With
            Me.Close()

            frmMainPro.Show()
            frmScreenBlk.Show()

        End If
End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
  ShowResvrd()
  ClearSubData2()
  'CallEditData()    'ซับรูทีนแสดง Size เพื่อแก้ไขข้อมูล
  gpbSeek.Text = "เพิ่มข้อมูล"
  cboPart.Enabled = True
  txtSize.ReadOnly = False
  txtSizeDesc.ReadOnly = False
End Sub

Private Sub ClearSubData2()
   txtCdate.Text = "__/__/____"
   txtSize.Text = ""
   txtSizeDesc.Text = ""
   txtSizeQty.Text = "0"
   txtSetQty.Text = "0"
   txtPrice.Text = "0.00"
   txtRmk.Text = ""
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

   strSqlSelc = " SELECT creat_date FROM eqpmst (NOLOCK)" _
                            & " WHERE eqp_id = '" & txtEqp_id.Text.ToString.Trim & "'"

   With Rsd

         .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
         .LockType = ADODB.LockTypeEnum.adLockOptimistic
         .Open(strSqlSelc, Conn, , , )

         If .RecordCount <> 0 Then

            If .Fields("creat_date").Value.ToString <> "" Then
                 txtCdate.Text = Mid(.Fields("creat_date").Value.ToString.Trim, 1, 10)
            Else
                 txtCdate.Text = "__/__/____"
            End If

         Else
            txtCdate.Focus()

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

     If dgvSize.Rows.Count <> 0 Then

        Dim strSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(4).Value.ToString.Trim         'เก็บ Size
        Dim strGroupSize As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(7).Value.ToString.Trim      'เก็บรหัสบล็อค

        With Conn

               If .State Then Close()
                  .ConnectionString = strConnAdodb
                  .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                  .ConnectionTimeout = 90
                  .Open()

        End With

          strSqlSelc = " SELECT * " _
                          & " FROM v_tmp_eqptrn (NOLOCK)" _
                          & " WHERE user_id = '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                          & " AND size_id= '" & strSize & "'" _
                          & " AND size_group = '" & strGroupSize & "'"

          With Rsd

                 .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                 .LockType = ADODB.LockTypeEnum.adLockOptimistic
                 .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                Select Case .Fields("backgup").Value.ToString.Trim

                    Case Is = "พื้นล่าง"
                        cboPart.Text = "พื้นล่าง"
                    Case Is = "พื้นบน"
                        cboPart.Text = "พื้นบน"
                    Case Is = "ใส้พื้นบน"
                       cboPart.Text = "ใส้พื้นบน"
                    Case Is = "โลโก้ส้น"
                        cboPart.Text = "โลโก้ส้น"
                    Case Is = "โลโก้พื้น"
                        cboPart.Text = "โลโก้พื้น"
                    Case Is = "EVA ติดส้น"
                        cboPart.Text = "EVA ติดส้น"
                    Case Is = "EVA บนหนังหน้า"
                        cboPart.Text = "EVA บนหนังหน้า"
                    Case Is = "หนังหน้า"
                       cboPart.Text = "หนังหน้า"
                    Case Is = "ONUPPER"
                       cboPart.Text = "ONUPPER"

                End Select

                     txtSize.Text = .Fields("size_id").Value.ToString.Trim
                     txtSizeDesc.Text = .Fields("size_group").Value.ToString.Trim
                     txtSetQty.Text = Format(.Fields("set_qty").Value, "#,##0.0")
                     txtSizeQty.Text = Format(.Fields("size_qty").Value, "#,##0.0")
                     txtPrice.Text = Format(.Fields("price").Value, "#,##0.00")
                     txtRmk.Text = .Fields("men_rmk").Value.ToString.Trim

         End If
               .ActiveConnection = Nothing    'เคลียร์การเชื่อมต่อ
               .Close()

         End With
            Rsd = Nothing

    Conn.Close()
    Conn = Nothing
End If

End Sub

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

Private Sub btnSeekExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekExit.Click
  StateLockFindDept(True)
  gpbSeek.Text = ""
  gpbSeek.Visible = False  'ทำให้ gpbSeek ซ่อน
  IsShowSeek = False
End Sub

Private Sub btnSeekSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekSave.Click
   CheckSubDataBfSave()
End Sub

Private Sub CheckSubDataBfSave()
 Dim i As Integer

     If txtSize.Text.Trim <> "" Then

        If gpbSeek.Text = "เพิ่มข้อมูล" Then
           SaveSubRecord()
        Else
            EditSubRecord()
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
  Dim strDateNull As String = "NULL"       'วันที่ค่าว่าง(Null)  

     Try

      With Conn
           If .State Then Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

     End With

        '------------------------------------ เช็คข้อมูล่ก่อนว่ามีอยู่หรือเปล่า -------------------------------------------------

        strSqlSelec = "SELECT size_id FROM tmp_eqptrn" _
                           & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                           & " AND size_group = '" & txtSizeDesc.Text.ToString.Trim & "'" _
                           & " AND size_id = '" & txtSize.Text.ToString.Trim & "'"

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



                '---------------------------------------- วดป.ที่ผลิด --------------------------------

                If txtCdate.Text <> "__/__/____" Then

                    strCreDate = Mid(txtCdate.Text.ToString, 7, 4) & "-" _
                                         & Mid(txtCdate.Text.ToString, 4, 2) & "-" _
                                         & Mid(txtCdate.Text.ToString, 1, 2)
                    strCreDate = "'" & SaveChangeEngYear(strCreDate) & "'"

                Else
                    strCreDate = "NULL"
                End If

                strSqlCmd = "INSERT INTO tmp_eqptrn " _
                                     & "(user_id,[group],eqp_id,size_id,size_desc,size_qty,weight" _
                                     & ",dimns,backgup,price,men_rmk,delvr_sta,sent_sta,set_qty" _
                                     & ",pr_date,pr_doc,recv_date,ord_rep,ord_qty,fc_date,impt_id" _
                                     & ",sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                     & ")" _
                                     & " VALUES (" _
                                     & "'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                     & ",'" & "F" & "'" _
                                     & ",'" & ReplaceQuote(txtEqp_id.Text.ToString.Trim) & "'" _
                                     & ",'" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                     & ",'" & "" & "'" _
                                     & "," & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                     & "," & 0.0 _
                                     & ",'" & "" & "'" _
                                     & ",'" & cboPart.Text.ToString.Trim & "'" _
                                     & "," & ChangFormat(txtPrice.Text.ToString.Trim) _
                                     & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                     & ",'" & "0" & "'" _
                                     & ",'" & "0" & "'" _
                                     & "," & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                     & "," & strDateNull _
                                     & ",'" & "" & "'" _
                                     & "," & strDateNull _
                                     & ",'" & "0" & "'" _
                                     & ",'" & "0" & "'" _
                                     & "," & strCreDate _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & ",'" & "" & "'" _
                                     & "," & 0.0 _
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

Private Sub EditSubRecord()

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

  Dim dateSave As Date = Now()
  Dim strDate As String = ""

  Dim strCredate As String
  Dim strDocdate As String            'เก็บสตริงวันที่เอกสาร
  Dim strGpType As String = ""        'เก็บประเภทอุปกรณ์
  Dim strPartType As String = ""      'เก็บชิ้นส่วนที่ผลิต
  Dim strDateNull As String = "NULL"

      Try

        With Conn
            If .State Then Close()

            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()

        End With


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

                       strSqlCmd = "UPDATE  tmp_eqptrn SET size_qty =" & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                                        & "," & "dimns ='" & "" & "'" _
                                                        & "," & "price = " & ChangFormat(txtPrice.Text.ToString.Trim) _
                                                        & "," & "backgup = '" & cboPart.Text.ToString.Trim & "'" _
                                                        & "," & "pr_doc ='" & "" & "'" _
                                                        & "," & "men_rmk ='" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                                        & "," & "set_qty =" & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                                        & "," & "pr_date = " & strDateNull _
                                                        & "," & "recv_date = " & strDateNull _
                                                        & "," & "fc_date = " & strDateNull _
                                                        & "," & "sup_name = '" & "" & "'" _
                                                        & "," & "lp_type = '" & "" & "'" _
                                                        & "," & "size_group = '" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                                        & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                                        & " AND size_id ='" & txtSize.Text.ToString.Trim & "'" _
                                                        & " AND size_group = '" & txtSizeDesc.Text.ToString.Trim & "'"

                       Conn.Execute(strSqlCmd)

   Conn.Close()
   Conn = Nothing

   Catch ex As Exception
         MsgBox("พบข้อผิดพลาดขณะทำการบันทึก โปรดดำเนินการใหม่อีกครั้ง", MsgBoxStyle.Critical, "ผิดพลาด")
         MsgBox(ex.Message)
   End Try

End Sub

Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

 If dgvSize.Rows.Count > 0 Then
     ShowResvrd()       'แสดง gpbSeek 
     CallEditData()    'ซับรูทีนแสดง Size เพื่อแก้ไขข้อมูล
     CallEditData2()   'ซับรูทีนแสดงข้อมูลทางเทคนิค

     gpbSeek.Text = "แก้ไขข้อมูล"
     cboPart.Enabled = True
     txtSize.ReadOnly = True
     txtSizeDesc.ReadOnly = True
  Else
      MsgBox("ไม่มีรายการ SIZE ที่ต้องการแก้ไข!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
      dgvSize.Focus()
  End If

End Sub

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
  DeleteSubData()
End Sub

Private Sub DeleteSubData()
 Dim btyConsider As Byte
 Dim strSize As String = ""
 Dim strSizeAct As String = ""
 Dim strSizeBlock As String = ""
 Dim strGpsize As String

     With dgvSize

        If .Rows.Count > 0 Then

            strSize = .Rows(.CurrentRow.Index).Cells(4).Value.ToString
            strSizeAct = .Rows(.CurrentRow.Index).Cells(5).Value.ToString
            strSizeBlock = .Rows(.CurrentRow.Index).Cells(6).Value.ToString
            strGpsize = .Rows(.CurrentRow.Index).Cells(7).Value.ToString

             If strSizeAct <> "" Then

                    btyConsider = MsgBox("SIZE : " & strSize.ToString.Trim & vbNewLine _
                                                   & "รหัสบล็อค : " & strSizeBlock.ToString.Trim & vbNewLine _
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
                                                  & " AND size_desc = '" & strSizeBlock.ToString.Trim & "'" _
                                                  & " AND size_group = '" & strGpsize.ToString.Trim & "'"

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

Private Sub txtEqp_id_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqp_id.LostFocus
 txtEqp_id.Text = txtEqp_id.Text.ToUpper.Trim
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
                     txtOrder.Focus()
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

Private Sub txtEqpnm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpnm.LostFocus
  txtEqpnm.Text = txtEqpnm.Text.ToUpper.Trim
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
                     cboPart.DroppedDown = True
                     cboPart.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With
End Sub

Private Sub txtShoe_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtShoe.KeyPress
 If e.KeyChar = Chr(13) Then
    txtOrder.Focus()
 End If
End Sub

Private Sub txtShoe_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShoe.LostFocus
 txtShoe.Text = txtShoe.Text.ToUpper.Trim
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
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                       txtRemark.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
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

Private Sub txtOrder_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrder.KeyPress
 If e.KeyChar = Chr(13) Then
   txtRemark.Focus()
 End If
End Sub

Private Sub txtOrder_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrder.LostFocus
  txtOrder.Text = txtOrder.Text.ToUpper.Trim
End Sub

Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
  Dim intChkPoint As Integer
        With txtOrder
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                      txtOrder.Focus()
                    End If
                Case Is = 38 'ปุ่มลูกศรขึ้น      
                       txtOrder.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                       txtCdate.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                         txtCdate.Focus()
                        End If
                    End If

                Case Is = 40 'ปุ่มลง
                    txtCdate.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

End Sub

Private Sub txtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.LostFocus
  txtRemark.Text = txtRemark.Text.ToUpper.Trim
End Sub

Private Sub cboPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboPart.KeyPress
  If e.KeyChar = Chr(13) Then
     txtSize.Focus()
  End If
End Sub

Private Sub txtCdate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCdate.GotFocus
 With mskCdate
      .BringToFront()
      txtCdate.SendToBack()
      .Focus()
 End With
End Sub

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown
  Dim intChkPoint As Integer

        With txtSize
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                    If .SelectionStart = 0 Then
                        cboPart.DroppedDown = True
                    End If

                Case Is = 38 'ปุ่มลูกศรขึ้น    
                     txtCdate.Focus()

                Case Is = 39 'ปุ่มลูกศรขวา

                    If .SelectionLength = .Text.Trim.Length Then
                        txtSizeDesc.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                            txtSizeDesc.Focus()
                        End If

                    End If

                Case Is = 40 'ปุ่มลง    
                     txtSetQty.Focus()
                Case Is = 113 'ปุ่ม F2
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
 If e.KeyChar = Chr(13) Then
    txtSetQty.Focus()
 End If
End Sub

Private Sub txtSizeDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeDesc.LostFocus
 txtSizeDesc.Text = txtSizeDesc.Text.ToUpper.Trim
End Sub

Private Sub txtSetQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSetQty.GotFocus

  With mskSetQty
       .BringToFront()
       txtSetQty.SendToBack()
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

Private Sub mskSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSetQty.KeyPress

  Select Case Asc(e.KeyChar)

         Case 48 To 57            ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
         Case 13
                e.Handled = False
                txtSizeQty.Focus()
         Case 8                   ' ปุ่ม Backspace
                e.Handled = False
         Case 32                   'เคาะ spacebar
                e.Handled = False
         Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtSetQty.Focus()

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
                    txtSize.Focus()

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
                    txtRmk.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
        End With

End Sub

Private Sub txtPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrice.GotFocus
 With mskPrice
      .BringToFront()
      txtPrice.SendToBack()
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
                Case Is = 38   'ปุ่มลูกศรขึ้น
                    txtSize.Focus()
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        txtRmk.Focus()
                    Else
                        intChkPoint = .Text.ToString.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
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

Private Sub mskPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrice.KeyPress

 Select Case Asc(e.KeyChar)

            Case 48 To 57            ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 13                  'ปุ่ม Enter
                e.Handled = False
                txtRmk.Focus()
            Case 8                   ' ปุ่ม Backspace
                e.Handled = False
            Case 32                   ' ปุ่ม Tab
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPrice.Focus()
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

Private Sub txtRmk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRmk.KeyDown
  Dim intChkPoint As Integer

      With txtRmk
            Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                    If .SelectionStart = 0 Then
                        txtPrice.Focus()
                    End If
                Case Is = 38   'ปุ่มลูกศรขึ้น
                    txtSetQty.Focus()
                Case Is = 39   'ปุ่มลูกศรขวา
                    If .SelectionLength = .Text.Trim.Length Then  'ถ้าตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                        btnSeekSave.Focus()
                    Else
                        intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then   'ให้ Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                           btnSeekSave.Focus()
                        End If
                    End If
                Case Is = 40 'ปุ่มลง

                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select

        End With
End Sub

Private Sub txtRmk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRmk.KeyPress
  If e.KeyChar = Chr(13) Then
     btnSeekSave.Focus()
  End If
End Sub

Private Sub txtPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
 If e.KeyChar = Chr(13) Then
    txtSize.Focus()
 End If
End Sub

Private Sub btnDelEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  CallEditData()   'สั่งเรียกข้อมูลทางเทคนิคขึ้นมารอ 
End Sub

Private Sub mskSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.GotFocus
 Dim i, x As Integer

 Dim strTmp As String = ""
 Dim strMerge As String = ""

     With mskSizeQty

        If txtSizeQty.Text.ToString.Trim <> "0" Then
           x = Len(txtSizeQty.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtSizeQty.Text.ToString, i, 1)

                                Select Case strTmp

                                          Case Is = "_"
                                          Case Else

                                                    If InStr("0123456789.", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select

                         Next i


                 Select Case strMerge.IndexOf(".")

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

                        .SelectedText = strMerge

                End If

        .SelectAll()

End With

End Sub

Private Sub mskSizeQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSizeQty.KeyDown
 Dim intChkPoint As Integer

        With mskSizeQty

            Select Case e.KeyCode

                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย

                         If .SelectionStart = 0 Then
                             mskSetQty.Focus()
                         End If

                   Case Is = 38 'ปุ่มลูกศรขึ้น    
                          cboPart.DroppedDown = True

                   Case Is = 39 'ปุ่มลูกศรขวา

                        If .SelectionLength = .Text.Trim.Length Then
                           txtPrice.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length
                            If .SelectionStart = intChkPoint Then
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

Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress

      Select Case Asc(e.KeyChar)

            Case 48 To 57            ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 13                  'ปุ่ม Enter
                e.Handled = False
                txtPrice.Focus()
            Case 8                   ' ปุ่ม Backspace
                e.Handled = False
            Case 32                   ' ปุ่ม Tab
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นตัวเลข", "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                mskSizeQty.Focus()
      End Select

End Sub

Private Sub txtSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeQty.GotFocus
  With mskSizeQty
       .BringToFront()
       txtSizeQty.SendToBack()
       .Focus()
  End With
End Sub

Private Sub mskSizeQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.LostFocus

 Dim i, x As Integer
 Dim z As Double

 Dim strTmp As String = ""
 Dim strMerge As String = ""
 Dim intFull As Integer

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

              intFull = Int(z)

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
                   Case Is = 38 'ลูกศรขึ้น
                   Case Is = 39   'ปุ่มลูกศรขวา
                        If .SelectionLength = .Text.Trim.Length Then
                           cboPart.Focus()
                        Else
                            intChkPoint = .Text.Trim.Length
                            If .SelectionStart = intChkPoint Then
                               cboPart.Focus()
                            End If
                        End If

                  Case Is = 40 'ปุ่มลง
                       cboPart.Focus()
                  Case Is = 113 'ปุ่ม F2
                       .SelectionStart = .Text.Trim.Length
            End Select

        End With

End Sub

Private Sub mskCdate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskCdate.KeyPress
    If e.KeyChar = Chr(13) Then
       cboPart.Focus()
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

                'If Year(z) < 2500 Then  'กรณีกรอกเป็น ค.ศ. จะเเปลงเป็น พ.ศ. ทันที
                '    txtRecvDate.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))
                'Else
                '    txtRecvDate.Text = z.ToString("dd/MM/yyyy")
                'End If

            Catch ex As Exception
                  mskCdate.Text = "__/__/____"
                  txtCdate.Text = "__/__/____"
            End Try

          mskCdate.SendToBack()
          txtCdate.BringToFront()

        End With

End Sub

End Class