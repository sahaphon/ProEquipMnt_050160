Imports ADODB
Imports System.Data.OleDb.OleDbDataAdapter
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports System.Drawing

Public Class frmRptEqpTrnsf
Dim cryRpt As New ReportDocument
Dim strUser As String

Private Sub frmRptEqpTrnsf_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    Me.Dispose() 'เมื่อ close Form ให้ทำลายฟอร์มนั้นเสีย แล้วคืนค่าให้หน่วยความจำ
End Sub

Private Sub frmRptEqpTrnsf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlCmdSelc As String

  Dim da As New System.Data.OleDb.OleDbDataAdapter
  Dim ds As New DataSet


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

          strSqlCmdSelc = "SELECT DISTINCT eqp_id FROM v_rpt_delvr"

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

 btnCancle.Enabled = False
 btnExit.Enabled = False

End Sub

Private Function StartDate() As String    'ฟ้งก์ชั่นแปลงปี ค.ศ.(DatetimePicker_ เริ่ม)

      Dim Sdate, Sdate2, Sdate3, DateEn As String

      Sdate = Mid(dtpStart.Value.ToString("yyyy-MM-dd"), 1, 4)
      Sdate2 = Mid(dtpStart.Value.ToString("yyyy-MM-dd"), 6, 2)
      Sdate3 = Mid(dtpStart.Value.ToString("yyyy-MM-dd"), 9, 2)

      If CInt(Sdate) > 2500 Then
         Sdate = CInt(Sdate) - 543
         DateEn = Sdate & "-" & Sdate2 & "-" & Sdate3

      Else
         DateEn = dtpStart.Value.ToString("yyyy-MM-dd")
      End If

     Return DateEn

End Function

Private Function EndDate() As String   'ฟ้งก์ชั่นแปลงปี ค.ศ. (DatetimePicker_ถึง)
      Dim Edate, Edate2, Edate3, DateEn As String

      Edate = Mid(dtpEnd.Value.ToString("yyyy-MM-dd"), 1, 4)  'ตัดสตริงเอาเฉพาะปี
      Edate2 = Mid(dtpEnd.Value.ToString("yyyy-MM-dd"), 6, 2) 'ตัดสตริงเอาเฉพาะเดือน
      Edate3 = Mid(dtpEnd.Value.ToString("yyyy-MM-dd"), 9, 2) 'ตัดสตริงเอาเฉพาะวันที่

      If CInt(Edate) > 2500 Then       'ค่าที่ได้เป็นปี พ.ศ.
             Edate = CInt(Edate) - 543
            DateEn = Edate & "-" & Edate2 & "-" & Edate3

      Else                             ' ค่าที่ได้เป็นปี พ.ศ.
             DateEn = dtpEnd.Value.ToString("yyyy-MM-dd")
      End If

     Return DateEn

End Function

Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
Dim Conn As New ADODB.Connection

Dim Rsd1 As New ADODB.Recordset
Dim strSqlCmdSelc As String

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim SDThai, EDThai As String   'ตัวเเปรเก็บสตริงวันที่จากฟังก์ชั่น StartDate() และ EndDate()


        With Conn

            If .State Then .Close()

                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .CommandTimeout = 30
                .Open()

        End With
                    '-------------------------คิวรี่ข้อมูล(กรณีเลือกทุกอุปกรณ์ และ เลือกทุกวันที่)---------------------------

                    If ChkAllEqp.Checked = True And ChkTime.Checked = True Then

                                 strSqlCmdSelc = "SELECT * FROM v_rpt_delvr (NOLOCK)"

                    '-------------------------คิวรี่ข้อมูล(กรณีเลือกทุกอุปกรณ์ แต่ระบุช่วงเวลา)-----------------------------

                         ElseIf ChkAllEqp.Checked = True And ChkTime.Checked = False Then

                                 strSqlCmdSelc = "SELECT * FROM v_rpt_delvr (NOLOCK)" _
                                               & " WHERE  (doc_date >= '" & StartDate() _
                                               & "' AND doc_date <= '" & EndDate() & "')"

                    '-------------------------คิวรี่ข้อมูล(กรณีระบุรหัสอุปกรณ์ แต่เลือกทุกวันที่)-----------------------------

                         ElseIf ChkAllEqp.Checked = False And ChkTime.Checked = True Then

                                 strSqlCmdSelc = "SELECT * FROM v_rpt_delvr (NOLOCK)" _
                                               & " WHERE eqp_id = '" & cboEqpid.Text & "' "

                         Else

                    '-------------------------คิวรี่ข้อมูล(กรณีระบุรหัสอุปกรณ์ และระบุช่วงเวลา)-----------------------------

                                 strSqlCmdSelc = "SELECT * FROM v_rpt_delvr (NOLOCK)" _
                                               & " WHERE eqp_id = '" & cboEqpid.Text & "' " _
                                               & " AND (doc_date >= '" & StartDate() _
                                               & "' AND doc_date <= '" & EndDate() & "')"

                    End If

       Rsd1 = New ADODB.Recordset
       With Rsd1
                .CursorType = CursorTypeEnum.adOpenKeyset
                .LockType = LockTypeEnum.adLockOptimistic
                .Open(strSqlCmdSelc, Conn, , )


                If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, Rsd1, "eqp")

                     cryRpt.Load(Application.StartupPath & "\BillDelvr2.rpt")
                     cryRpt.SetDatabaseLogon("sa", "sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("eqp"))

                '-------------------------ส่งค่าตัวแปรไปให้ Crystal Report (กรณีเช็คเลือกทุกอุปกรณ์ และ เช็คเลือกทุกวันที่)------------------------  

                If ChkAllEqp.Checked = True And ChkTime.Checked = True Then

                    Dim cryeqp As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryeqp = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryeqp")
                    cryeqp.Text = "ทุกรหัสอุปกรณ์"

                    Dim cryDf As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDf = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDf")
                    cryDf.Text = "-"

                    Dim cryDe As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDe = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDe")
                    cryDe.Text = "-"

                '-------------------------ส่งค่าตัวแปรไปให้ Crystal Report (กรณีเช็คเลือกทุกรหัสอุปกรณ์  แต่ไม่เลือกทุกวันที่)----------------------

                ElseIf ChkAllEqp.Checked = True And ChkTime.Checked = False Then

                    Dim cryeqp As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryeqp = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryeqp")
                    cryeqp.Text = "ทุกรหัสอุปกรณ์"

                    Dim cryDf As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDf = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDf")
                    SDThai = StartDate()
                    cryDf.Text = DateThai(SDThai)

                    Dim cryDe As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDe = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDe")
                    EDThai = DateThai(EndDate)
                    cryDe.Text = EDThai

                '-------------------------ส่งค่าตัวแปรไปให้ Crystal Report (กรณีไม่เลือกทุกอุปกรณ์ แต่เช็คเลือกทุกวันที่)------------------------  

                ElseIf ChkAllEqp.Checked = False And ChkTime.Checked = True Then

                    Dim cryeqp As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryeqp = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryeqp")
                    cryeqp.Text = cboEqpid.Text

                    Dim cryDf As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDf = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDf")
                    cryDf.Text = "-"

                    Dim cryDe As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDe = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDe")
                    cryDe.Text = "-"

                '--------------------- ส่งค่าตัวแปรไปให้ Crystal Report (กรณีไม่เช็คเลือกทุกอุปกรณ์ และ ไม่เช็คเลือกทุกวันที่) -------------------

                ElseIf ChkAllEqp.Checked = False And ChkTime.Checked = False Then

                    Dim cryeqp As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryeqp = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryeqp")
                    cryeqp.Text = cboEqpid.Text

                    Dim cryDf As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDf = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDf")
                    SDThai = StartDate()
                    cryDf.Text = DateThai(SDThai)

                    Dim cryDe As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDe = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDe")
                    EDThai = EndDate()
                    cryDe.Text = DateThai(EDThai)

                End If

               '----------------------- ตัวแปรจาก Crystal Report รับค่าจากตัวเเปรภายใน---- ------------------------

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Landscape

                     Viewer1.ReportSource = cryRpt
                     Viewer1.DisplayStatusBar = True
                     Viewer1.Refresh()
                     Viewer1.Zoom(80)

                     LockOptions(False)    'ให้ GroupboxMain และ GroupboxSub  ใช้งานไม่ได้ และปิดปุ่ม OK,Cancle

            Else
                
                MsgBox("ไม่พบข้อมูลที่คุณต้องการค้นหา!!" & vbNewLine _
                       & "กรุณาระบุข้อมูลใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                LockOptions(True)
               
                Exit Sub

                End If
                .ActiveConnection = Nothing
       End With

        btnOK.Enabled = False
        btnCancle.Enabled = True
        btnExit.Enabled = True

    Rsd1 = Nothing

    ds.Clear()    'เคลี่ยร์ค่าใน Dataset
    ds.Dispose()  'ทำลายตัวเเปร DataSet

    da.Dispose() 'ทำลายตัวเเปร DataAdapter

    ds = Nothing  'เคลียร์ข้อมูลใน Dataset
    da = Nothing  'เคลี่ยร์ข้อมูลใน adapter

  Conn.Close()
  Conn = Nothing

End Sub

    Function DateThai(ByVal str As String)  'ฟังก์ชั่นเเปลงวันที่จาก ("2013-10-XX") เป็น dd/MM/yyyy

        Dim yyyy, MM, dd As String
        Dim d As String

        'รับมาเป็น(yyyy - MM - dd) มีทั้งหมด 14 ตำเเหน่ง

        yyyy = Mid(str, 1, 4)  'ตัดปี
        dd = Mid(str, 9, 2)  'ตัดวันที
        MM = Mid(str, 6, 2)  'ตัดเดือน

        d = dd & "/" & MM & "/" & yyyy            'แปลงวันที่เป็น dd/MM/yyyy
        Return d
    End Function


    '--------------------------- Event เมื่อ CheckBox เลือกทุกรหัสอุปกรณ์  ---------------------------------------

Private Sub chkAllEqp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAllEqp.CheckedChanged
  If ChkAllEqp.Checked Then
     cboEqpid.Enabled = False
  Else
     cboEqpid.Enabled = True
  End If
  End Sub

    '--------------------------- Event เมื่อCheckBox เลือกข้อมูลเวลาทั้งหมด  ---------------------------------------

Private Sub chkTime_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkTime.CheckedChanged
  If ChkTime.Checked Then
     dtpStart.Enabled = False
     dtpEnd.Enabled = False
  Else
     dtpStart.Enabled = True
     dtpEnd.Enabled = True
  End If
  End Sub

Private Sub btnCancle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancle.Click
    Viewer1.ReportSource = Nothing   'เคลียร์ Report ที่โหลดขึ้นมา
    LockOptions(True)   'กรณีกดปุ่ม btnCancle ให้ปุ่ม btnOK, GroupboxMain และ GroupboxSub  ใช้งานได้
 End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
   Me.Close()
End Sub

Private Sub LockOptions(ByVal blnSta As Boolean)   'ซับรูทีนล็อคปุ่มกด โดยรับค่า boolean มา
    gpbMain.Enabled = blnSta
    gpbSub.Enabled = blnSta
    If blnSta Then   'ถ้าเป็นจริง
        btnOK.Enabled = blnSta       'ปุ่ม btnOK สามารถใช้งานได้
        btnCancle.Enabled = False    'ปุ่ม btnCancle ถูกล็อค
    Else                                  ' กรณีเป็นเท็จ
        btnOK.Enabled = blnSta       'ให้ปุ่ม btnOK ถุกล็อค
        btnCancle.Enabled = blnSta   'ให้ปุ่ม btnCancle ล็อค
    End If
End Sub

End Class