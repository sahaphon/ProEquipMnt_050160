Imports ADODB
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Drawing
Imports System.IO
Imports System.Drawing.Drawing2D
Imports System.Data.OleDb.OleDbDataAdapter
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared

Public Class frmRptDelvApprove2
 Dim cryRpt As New ReportDocument
 Dim strUser As String
 Dim intWidth As Integer

Private Sub frmRptDelvApprove2_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  CheckApprove()       'ลบข้อมูลที่โหลดไว้รอ และไม่ได้ดำเนินการ
  ClearData(0)
  frmMainPro.lblRptDesc.Text = ""
  Me.Dispose()

  frmApproveDelv.Show()
  frmMainPro.Show()

End Sub

Private Sub frmRptDelvApprove_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

 Me.WindowState = FormWindowState.Maximized
 StdDateTimeThai()                              'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module
 InsertNewDocid()                               'เพิ่มรหัส doc_id ก่อน พิมพ์
 HidePanel()
 Viewdata()

End Sub

Private Function CheckApprove() As Boolean       'เช็คว่ามีการ approve เกิดขึ้นหรือไม่

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String
 Dim strSqlCmd As String

      With Conn

              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()
     End With

     strSqlSelc = "SELECT ps01_result FROM delv_approve (NOLOCK)" _
                                    & " WHERE ps01_result = 'False'" _
                                    & " AND doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

               '----------------- ลบข้อมูลที่โหลดไว้รอผู้บันทึกอนุมัติ (กรณีไม่ได้ดำเนินการใดๆ) -----------------------------

               strSqlCmd = "DELETE delv_approve " _
                                   & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

               Conn.Execute(strSqlCmd)

          End If

      .ActiveConnection = Nothing
      .Close()
     End With

 Conn.Close()
 Conn = Nothing

End Function

Private Sub Viewdata()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim da As New System.Data.OleDb.OleDbDataAdapter
 Dim ds As New DataSet

     strUser = frmMainPro.lblLogin.Text.Trim.ToString 'ใช้ User

     With Conn
          If .State Then Close()
             .ConnectionString = strConnAdodb

             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .CommandTimeout = 30
             .Open()

     End With

        strSqlSelc = "SELECT  * FROM v_rpt_delvr2 (NOLOCK)" _
                            & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'" _
                            & " ORDER BY no  "

        Rsd = New ADODB.Recordset

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                If .Fields("ps02_result").Value = True Then
                    btnAcp2.Enabled = False
                    lblApp2.Visible = True
                End If

                If .Fields("ps03_result").Value = True Then
                    btnAcp3.Enabled = False
                    lblApp3.Visible = True
                End If

                If .Fields("ps04_result").Value = True Then
                    btnAcp4.Enabled = False
                    lblApp4.Visible = True
                End If


                ds.Clear()
                da.Fill(ds, Rsd, "delv")

                cryRpt.Load(Application.StartupPath & "\BillDelvr.rpt")
                cryRpt.SetDatabaseLogon("sa", "sa2008", "ADDASRV03", "DBequipmnt")
                cryRpt.ReportOptions.EnableSaveDataWithReport = False
                cryRpt.SetDataSource(ds.Tables("delv"))

                Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject               'ตัวแปรส่งค่าให้กับ CrystalReport user
                    cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryTxtUsr")   'cryuser1
                    cryuser1.Text = strUser

                    cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                    cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Portrait            'กำหนดกระดาษเเนวนอน

                    CrystalReportViewer1.ReportSource = cryRpt
                    CrystalReportViewer1.DisplayStatusBar = True
                    CrystalReportViewer1.Refresh()
                    CrystalReportViewer1.Zoom(100)

             Else

                MsgBox("ไม่มีข้อมูลสำหรับพิมพ์เอกสาร!!" & vbNewLine _
                             & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")
             End If

           .ActiveConnection = Nothing
           ' .Close()
        End With

    Rsd = Nothing

    ds.Clear()
    ds.Dispose()

    da.Dispose()

    ds = Nothing
    da = Nothing

Conn.Close()
Conn = Nothing

End Sub

Private Sub PrePrintdata()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim RsdPic As New ADODB.Recordset

 Dim strSqlSelc As String
 Dim strSqlCmdPic As String

     With Conn

              If .State Then .Close()

                 .ConnectionString = strConnAdodb
                 .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                 .ConnectionTimeout = 90
                 .Open()
     End With

     strSqlSelc = "SELECT * FROM delv_approve (NOLOCK)" _
                                    & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

            '--------------- บันทึกข้อมูลใน tmp_delv_approve ---------------------------------------

             strSqlCmdPic = "SELECT * FROM tmp_delv_approve (NOLOCK)"

             RsdPic = New ADODB.Recordset
             RsdPic.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
             RsdPic.LockType = ADODB.LockTypeEnum.adLockOptimistic
             RsdPic.Open(strSqlCmdPic, Conn, , , )

                                       RsdPic.AddNew()
                                       RsdPic.Fields("doc_id").Value = .Fields("doc_id").Value
                                       RsdPic.Fields("app_sta").Value = .Fields("app_sta").Value
                                       RsdPic.Fields("ps01_result").Value = .Fields("ps01_result").Value
                                       RsdPic.Fields("ps01_date").Value = .Fields("ps01_date").Value
                                       RsdPic.Fields("ps01_usr").Value = .Fields("ps01_usr").Value
                                       RsdPic.Fields("ps02_result").Value = .Fields("ps02_result").Value
                                       RsdPic.Fields("ps02_date").Value = .Fields("ps02_date").Value
                                       RsdPic.Fields("ps02_usr").Value = .Fields("ps02_usr").Value
                                       RsdPic.Fields("ps03_result").Value = .Fields("ps03_result").Value
                                       RsdPic.Fields("ps03_date").Value = .Fields("ps03_date").Value
                                       RsdPic.Fields("ps03_usr").Value = .Fields("ps03_usr").Value
                                       RsdPic.Fields("ps04_result").Value = .Fields("ps04_result").Value
                                       RsdPic.Fields("ps04_date").Value = .Fields("ps04_date").Value
                                       RsdPic.Fields("ps04_usr").Value = .Fields("ps04_usr").Value
                                       RsdPic.Fields("note").Value = .Fields("note").Value


                                       Dim RsdSteam As New ADODB.Stream
                                       Dim strPicSign01 As String
                                       Dim strPicSign02 As String
                                       Dim strPicSign03 As String
                                       Dim strPicSign04 As String

                                       RsdSteam.Type = StreamTypeEnum.adTypeBinary
                                       RsdSteam.Open()

                                       '--------------------------------------โหลดรูปภาพลายเซ็นต์ ผู้บันทึก -------------------------------------

                                       If .Fields("ps01_usr").Value.ToString.Trim <> "" Then
                                           strPicSign01 = CallPathSignPicture(.Fields("ps01_usr").Value.ToString.Trim)

                                       Else
                                           strPicSign01 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                       End If

                                       RsdSteam.LoadFromFile(strPicSign01)
                                       RsdPic.Fields("sign_pic01").Value = RsdSteam.Read


                                       '--------------------------------------โหลดรูปภาพลายเซ็นต์ หน.ส่วนเทคนิคอุปกรณ์ -------------------------------------

                                       If .Fields("ps02_usr").Value.ToString.Trim <> "" Then
                                           strPicSign02 = CallPathSignPicture(.Fields("ps02_usr").Value.ToString.Trim)

                                       Else
                                           strPicSign02 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                       End If

                                       RsdSteam.LoadFromFile(strPicSign02)
                                       RsdPic.Fields("sign_pic02").Value = RsdSteam.Read


                                      '--------------------------------------โหลดรูปภาพลายเซ็นต์ หน.ส่วนเทคนิคอุปกรณ์ -------------------------------------

                                       If .Fields("ps03_usr").Value.ToString.Trim <> "" Then
                                           strPicSign03 = CallPathSignPicture(.Fields("ps03_usr").Value.ToString.Trim)

                                       Else
                                           strPicSign03 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                       End If

                                       RsdSteam.LoadFromFile(strPicSign03)
                                       RsdPic.Fields("sign_pic03").Value = RsdSteam.Read


                                      '--------------------------------------โหลดรูปภาพลายเซ็นต์ หน.ส่วนเทคนิคอุปกรณ์ -------------------------------------

                                       If .Fields("ps04_usr").Value.ToString.Trim <> "" Then
                                           strPicSign04 = CallPathSignPicture(.Fields("ps04_usr").Value.ToString.Trim)

                                       Else
                                           strPicSign04 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                       End If

                                       RsdSteam.LoadFromFile(strPicSign04)
                                       RsdPic.Fields("sign_pic04").Value = RsdSteam.Read

                                      '-----------------------------------------------------------------------------------------------------------

                                       RsdPic.Fields("user_id").Value = frmMainPro.lblLogin.Text

                   RsdSteam.Close()
                   RsdSteam = Nothing

                   RsdPic.Update()

                   RsdPic.ActiveConnection = Nothing
                   RsdPic.Close()
                   RsdPic = Nothing

             Else
                  MsgBox("ไม่มีข้อมูล")

             End If

            .ActiveConnection = Nothing
            .Close()
     End With

 Rsd = Nothing

Conn.Close()
Conn = Nothing
End Sub

Private Sub ClearData(ByVal strCase As String) 'เคลียร์ tmp table
 Dim Conn As New ADODB.Connection
 Dim strSqlCmd As String

     With Conn
          .ConnectionString = strConnAdodb
          .CursorLocation = ADODB.CursorLocationEnum.adUseClient
          .ConnectionTimeout = 90
          .Open()

     End With


         Select Case strCase

                Case Is = "0"

                           '-------------- Clear tmp_delv_apporve -------------------------------

                           strSqlCmd = "DELETE FROM tmp_delv_approve " _
                                         & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

                           Conn.Execute(strSqlCmd)


                           '-------------- Clear tmp_rpt_delvr ----------------------------------

                           strSqlCmd = "DELETE FROM tmp_rpt_delvr " _
                                         & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

                           Conn.Execute(strSqlCmd)


                Case Is = "1"

                            '-------------- Clear tmp_rpt_delvr ก่อนเพิ่มข้อมูล ------------------------

                            strSqlCmd = "DELETE FROM tmp_delv_approve " _
                                           & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

                            Conn.Execute(strSqlCmd)


          End Select

 Conn.Close()
 Conn = Nothing
End Sub

Private Sub HidePanel()
 Dim z As String
 Dim x() As String

     ScreenResolution()   'อ่านค่า screen resolution W x H ของเครื่อง

     z = CStr(ScreenResolution())    'เเปลงค่า W x H เป็น สตริง
     x = z.Split(" x ")              'ต้ด  x ออก แล้วเเปลงเป็น Array โดยใช้ฟังก์ชั่น Split()

     intWidth = CInt(x(0))          'เก็บค่าในตัวเเปร intWidth ของฟอร์ม

     spnRpt.SplitterDistance = intWidth
     btnFeed.Text = ">"

End Sub

Private Sub btnFeed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFeed.Click

 With btnFeed

      If .Text = ">" Then
         .Text = "<"
          spnRpt.SplitterDistance = 900

      Else
          .Text = ">"
          spnRpt.SplitterDistance = intWidth

      End If
  End With
End Sub

Private Function ChkDocidExist(ByVal strDocid As String) As Boolean       'ตรวจสอบว่ามีการเพิ่มรหัส doc_id ใน table แล้วหรือไม่
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlSelc As String

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 150
              .Open()
      End With

      strSqlSelc = "SELECT * FROM delv_approve (NOLOCK) " _
                                     & " WHERE doc_id = '" & strDocid & "'"

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

Private Sub SaveApprove_ps01()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim datSave As Date = Now()
 Dim strdateNull As String = "NULL"
 Dim strDate As String = datSave.Date.ToString("yyyy-MM-dd")

      strDate = "'" & SaveChangeEngYear(strDate) & "'"

     With Conn
          If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

          strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                    & "," & "ps01_result ='" & "1" & "'" _
                    & "," & "ps01_date = " & strDate _
                    & "," & "ps01_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                    & "," & "ps02_result ='" & "0" & "'" _
                    & "," & "ps02_date = " & strdateNull _
                    & "," & "ps02_usr = '" & "" & "'" _
                    & "," & "ps03_result ='" & "0" & "'" _
                    & "," & "ps03_date = " & strdateNull _
                    & "," & "ps03_usr = '" & "" & "'" _
                    & "," & "ps04_result ='" & "0" & "'" _
                    & "," & "ps04_date = " & strdateNull _
                    & "," & "ps04_usr = '" & "" & "'" _
                    & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

        Conn.Execute(strSqlCmd)

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub SaveApprove_ps02()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim datSave As Date = Now()
 Dim strdateNull As String = "NULL"
 Dim strDate As String = datSave.Date.ToString("yyyy-MM-dd")

      strDate = "'" & SaveChangeEngYear(strDate) & "'"

     With Conn
          If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

          strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                    & "," & "ps02_result ='" & "1" & "'" _
                    & "," & "ps02_date = " & strDate _
                    & "," & "ps02_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                    & "," & "ps03_result ='" & "0" & "'" _
                    & "," & "ps03_date = " & strdateNull _
                    & "," & "ps03_usr = '" & "" & "'" _
                    & "," & "ps04_result ='" & "0" & "'" _
                    & "," & "ps04_date = " & strdateNull _
                    & "," & "ps04_usr = '" & "" & "'" _
                    & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

          Conn.Execute(strSqlCmd)

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub SaveApprove_ps03()
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim datSave As Date = Now()
 Dim strdateNull As String = "NULL"
 Dim strDate As String = datSave.Date.ToString("yyyy-MM-dd")

      strDate = "'" & SaveChangeEngYear(strDate) & "'"

     With Conn
          If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

          strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                    & "," & "ps03_result ='" & "1" & "'" _
                    & "," & "ps03_date = " & strDate _
                    & "," & "ps03_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                    & "," & "ps04_result ='" & "0" & "'" _
                    & "," & "ps04_date = " & strdateNull _
                    & "," & "ps04_usr = '" & "" & "'" _
                    & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

          Conn.Execute(strSqlCmd)

  Conn.Close()
  Conn = Nothing
End Sub

Private Sub SaveApprove_ps04(ByVal user As String)
 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim datSave As Date = Now()
 Dim strdateNull As String = "NULL"
 Dim strDate As String = datSave.Date.ToString("yyyy-MM-dd")

      strDate = "'" & SaveChangeEngYear(strDate) & "'"

     With Conn
          If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

          Select Case user         ' Approve แยกตามแผนก

                 Case Is = "PEERA"    'ผช.ผจก. แผนกฉีด PVC

                         strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                                   & "," & "ps04_result ='" & "1" & "'" _
                                   & "," & "ps04_date = " & strDate _
                                   & "," & "ps04_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                   & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

                         Conn.Execute(strSqlCmd)

                 Case Is = "ITHISAK"            'หน.ส่วนอาวุโส แผนกฉีด EVA

                         strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                                   & "," & "ps04_result ='" & "1" & "'" _
                                   & "," & "ps04_date = " & strDate _
                                   & "," & "ps04_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                   & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

                        Conn.Execute(strSqlCmd)

                 Case Is = "PRADIST"             'ผจก.แผนกตัดชิ้นส่วน

                         strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                                   & "," & "ps04_result ='" & "1" & "'" _
                                   & "," & "ps04_date = " & strDate _
                                   & "," & "ps04_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                   & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

                         Conn.Execute(strSqlCmd)

                 Case Is = "TECHIN"         'หน.ส่วนอาวุโส แผนกผลิตโฟม

                         strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                                   & "," & "ps04_result ='" & "1" & "'" _
                                   & "," & "ps04_date = " & strDate _
                                   & "," & "ps04_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                   & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

                         Conn.Execute(strSqlCmd)

                 Case Is = "TODSAPORN"      'ผจก.แผนกฉีด PU

                          strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                                   & "," & "ps04_result ='" & "1" & "'" _
                                   & "," & "ps04_date = " & strDate _
                                   & "," & "ps04_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                   & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

                          Conn.Execute(strSqlCmd)

                 Case Is = "SATHID"            'หน.ส่วนอวุโส แผนกเย็บ

                          strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "0" & "'" _
                                   & "," & "ps04_result ='" & "1" & "'" _
                                   & "," & "ps04_date = " & strDate _
                                   & "," & "ps04_usr = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                   & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

                          Conn.Execute(strSqlCmd)
          End Select


  Conn.Close()
  Conn = Nothing

End Sub

Private Sub InsertNewDocid()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String
 Dim strSqlCmd As String

 Dim strdateNull As String = "NULL"
 Dim datSave As Date = Now()
 Dim strDate As String = datSave.Date.ToString("yyyy-MM-dd")

     strDate = "'" & SaveChangeEngYear(strDate) & "'"

     With Conn
          If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

      strSqlSelc = "SELECT *  FROM v_rpt_delvr (NOLOCK)"

      With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

         If ChkDocidExist(frmMainPro.lblRptDesc.Text) Then        'เช็คว่ามีการบันทึก รหัสเอกสารแล้วหรือยัง

            strSqlCmd = " INSERT INTO delv_approve " _
                      & "(doc_id,app_sta,ps01_result,ps01_date,ps01_usr,ps02_result," _
                      & "ps02_date,ps02_usr,ps03_result,ps03_date,ps03_usr,ps04_result," _
                      & "ps04_date,ps04_usr,note" _
                      & ")" _
                      & " VALUES (" _
                      & "'" & frmMainPro.lblRptDesc.Text & "'" _
                      & ",'" & "0" & "'" _
                      & ",'" & "0" & "'" _
                      & "," & strdateNull _
                      & ",'" & "" & "'" _
                      & ",'" & "0" & "'" _
                      & "," & strdateNull _
                      & ",'" & "" & "'" _
                      & ",'" & "0" & "'" _
                      & "," & strdateNull _
                      & ",'" & "" & "'" _
                      & ",'" & "0" & "'" _
                      & "," & strdateNull _
                      & ",'" & "" & "'" _
                      & ",'" & "" & "'" _
                      & ")"

                 Conn.Execute(strSqlCmd)

                                '---------------------- เพิ่มข้อมูลใน tmp_rpt_delvr ---------------------------

                                 strSqlCmd = " INSERT INTO tmp_rpt_delvr " _
                                                 & " SELECT *" _
                                                 & " FROM v_rpt_delvr" _
                                                 & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

                                 Conn.Execute(strSqlCmd)
                                 PrePrintdata() 'นำเข้ารูปลายเซ็นต์


            Else

                                 '---------------------- เพิ่มข้อมูลใน tmp_rpt_delvr ---------------------------

                                 strSqlCmd = " INSERT INTO tmp_rpt_delvr " _
                                                 & " SELECT *" _
                                                 & " FROM v_rpt_delvr" _
                                                 & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

                                 Conn.Execute(strSqlCmd)
                                 PrePrintdata() 'นำเข้ารูปลายเซ็นต์

            End If

        .ActiveConnection = Nothing
        .Close()

      End With


  Conn.Close()
  Conn = Nothing

End Sub

Private Sub btnAcp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcp2.Click

 If frmMainPro.lblLogin.Text = "SUTID" Then          'ตรวจสอบชื่อผู้ใช้

     If chkApprove_ps01() Then       'ตรวจสอบว่าผู้จัดของ Approve แล้วหรือยัง

        SaveApprove_ps02()  'เพิ่มรายการใน table approve_delv
        ClearData(1)        'เคลี่ยร์ข้อมูลใน tmp_rpt_delv2 
        PrePrintdata()
        Viewdata()

        btnAcp2.Enabled = False

     Else
          MsgBox("ไม่สามารถดำเนินการได้  ผู้บันทึกยังไม่ได้อนุมัติ!...")

     End If

  Else
      MsgBox("คุณไม่มีสิทธิใช้งานส่วนนี้")
  End If

End Sub

Private Function chkApprove_ps01() As Boolean         'ฟังก์ชั่นตรวจสอบการเซ็น approve

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

     strSqlSelc = "SELECT * FROM delv_approve (NOLOCK)" _
                          & " WHERE ps01_result = 'True' AND doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

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

Private Function chkApprove_ps02() As Boolean         'ฟังก์ชั่นตรวจสอบการเซ็น approve

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

     strSqlSelc = "SELECT * FROM delv_approve (NOLOCK)" _
                          & " WHERE ps02_result = 'True' AND doc_id = '" & frmMainPro.lblRptDesc.Text & "' "

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

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

Private Function chkApprove_ps03() As Boolean         'ฟังก์ชั่นตรวจสอบการเซ็น approve

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

     strSqlSelc = "SELECT * FROM delv_approve (NOLOCK)" _
                          & " WHERE ps03_result = 'True' AND doc_id = '" & frmMainPro.lblRptDesc.Text & "' "

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .RecordCount <> 0 Then

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

Private Sub btnAcp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcp3.Click

 If frmMainPro.lblLogin.Text = "BOONTUM" Then          'ตรวจสอบชื่อผู้ใช้

    If chkApprove_ps02() Then                          'ตรวจสอบว่า หน.ส่วนอุปกรณ์ Approve แล้วหรือยัง

       SaveApprove_ps03()  'เพิ่มรายการใน table approve_delv
       ClearData(1)      'เคลี่ยร์ข้อมูลใน tmp_rpt_delv2 
       PrePrintdata()
       Viewdata()

       btnAcp3.Enabled = False

    Else

        MsgBox("ไม่สามารถดำเนินการได้  หน.ส่วนเทคนิคอุปกรณ์ยังไม่ได้อนุมัติ!...")
    End If

  Else
      MsgBox("คุณไม่มีสิทธิใช้งานส่วนนี้")
  End If

End Sub

Private Sub btnAcp4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcp4.Click

 If (frmMainPro.lblLogin.Text = "PEERA" Or frmMainPro.lblLogin.Text = "ITHISAK") Or _
                     (frmMainPro.lblLogin.Text = "PRADIST" Or frmMainPro.lblLogin.Text = "TECHIN") Or _
                     (frmMainPro.lblLogin.Text = "TODSAPORN" Or frmMainPro.lblLogin.Text = "SATHID") Then ' ตรวจสอบ user ที่สามารถอนุมัติได้

    If SignCompare() Then

          If chkApprove_ps03() Then         'ตรวจสอบว่า ผจก.เทคนิคอุปกรณ์ Approve แล้วหรือยัง

             SaveApprove_ps04(frmMainPro.lblLogin.Text)         'เพิ่มรายการใน table approve_delv
             SaveCopleteApprove()
             ClearData(1)                   'เคลี่ยร์ข้อมูลใน tmp_rpt_delv2 
             PrePrintdata()
             Viewdata()

             btnAcp3.Enabled = False

          Else

             MsgBox("ไม่สามารถดำเนินการได้  ผจก.แผนกเทคนิคอุปกรณ์ยังไม่ได้อนุมัติ!...")

          End If

    Else
        MsgBox("คุณไม่มีสิทธิใช้งานส่วนนี้")

    End If

 Else
      MsgBox("คุณไม่มีสิทธิใช้งานส่วนนี้")
 End If

End Sub

Private Function SignCompare() As Boolean       'ฟังก์ชั่นเปรียบเทียบ รหัส login กับ ผู้รับโอนอุปกรณ์

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlSelc As String

 Dim strDept As String
 Dim strMerg As String = ""
 Dim strDeptid As String = ""

     With Conn
           If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

     strSqlSelc = " SELECT * FROM v_rpt_delvr2 (NOLOCK)" _
                                   & " WHERE doc_id = '" & frmMainPro.lblRptDesc.Text & "'"

     Conn.Execute(strSqlSelc)

       With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )


            If .RecordCount <> 0 Then

               strDept = Mid(.Fields("rvc_dep_nm").Value.ToString.Trim, 1, 6)      'แผนกที่รับ
               strDept = Mid(strDept, 1, 6)    'ตัดเอารหัสแผนก

                    Select Case frmMainPro.lblLogin.Text

                           Case Is = "PEERA"     'ฉีด PVC
                                strDeptid = "124000"

                           Case Is = "ITHISAK"   'ฉีด EVA
                                strDeptid = "125000"

                           Case Is = "PRADIST"   'ตัดชิ้นส่วน
                                strDeptid = "122000"

                           Case Is = "SATHID"     'เย็บ
                                strDeptid = "123000"

                           Case Is = "TECHIN"     'ผลิตโฟม
                                strDeptid = "121000"

                           Case Is = "TODSAPORN"  'ฉีด PU
                                strDeptid = "126000"
                    End Select


               If String.Compare(strDept, strDeptid) = 0 Then      'ตรวจสอบว่าแผนกรับโอน กับ ผู้อนุมัติรับโอนอยู่แผนกเดียวกันหรือไม่
                   Return True

                Else
                   Return False

                End If

            End If

         .ActiveConnection = Nothing
         .Close()

       End With

  Conn.Close()
  Conn = Nothing

End Function

Private Sub SaveCopleteApprove()            'อัพเดทสถานะ app_sta (Approve status)

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim datSave As Date = Now()
 Dim strdateNull As String = "NULL"
 Dim strDate As String = datSave.Date.ToString("yyyy-MM-dd")

     strDate = "'" & SaveChangeEngYear(strDate) & "'"

     With Conn
          If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

          strSqlCmd = "UPDATE delv_approve SET app_sta = '" & "1" & "'" _
                                 & " WHERE doc_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

         Conn.Execute(strSqlCmd)

  Conn.Close()
  Conn = Nothing

End Sub

End Class