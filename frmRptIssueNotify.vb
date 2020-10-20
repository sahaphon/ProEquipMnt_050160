Imports ADODB
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Data.OleDb.OleDbDataAdapter
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared

Public Class frmRptIssueNotify
 Dim cryRpt As New ReportDocument
 Dim strUser As String
 Dim intWidth As Integer

Private Sub frmRptIssueNotify_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearData(0)
  lblComplete.Text = frmMainPro.lblRptDesc.Text
  frmMainPro.lblRptDesc.Text = ""
  Me.Dispose()

  frmMainPro.Show()
  frmApproveIssue.Show()

End Sub

Private Sub frmRptIssueNotify_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Me.WindowState = FormWindowState.Maximized
  StdDateTimeThai()        'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module

  PrePrintData(frmMainPro.lblRptDesc.Text)
  HidePanel()
  Viewdata()

End Sub

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

        strSqlSelc = "SELECT  * FROM tmp_notifyissue (NOLOCK)" _
                                    & " WHERE req_id = '" & frmMainPro.lblRptDesc.Text & "'"

        Rsd = New ADODB.Recordset

        With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                If .Fields("person2").Value.ToString <> "" Then
                    btnAcp1.Enabled = False
                    lblApp1.Visible = True
                End If

                If .Fields("person4").Value.ToString <> "" Then
                    btnAcp3.Enabled = False
                    lblApp3.Visible = True
                End If


                ds.Clear()
                da.Fill(ds, Rsd, "issue")

                cryRpt.Load(Application.StartupPath & "\issuenotify.rpt")
                cryRpt.SetDatabaseLogon("sa", "sa2008", "ADDASRV03", "DBequipmnt")
                cryRpt.ReportOptions.EnableSaveDataWithReport = False
                cryRpt.SetDataSource(ds.Tables("issue"))

                Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                    cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryTxtUsr") ' cryuser1
                    cryuser1.Text = strUser

                    cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                    cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Portrait           'กำหนดกระดาษเเนวนอน

                    CrystalReportViewer1.ReportSource = cryRpt
                    CrystalReportViewer1.DisplayStatusBar = True
                    CrystalReportViewer1.Refresh()
                    CrystalReportViewer1.Zoom(100)

             Else

                MsgBox("ไม่มีข้อมูลสำหรับพิมพ์เอกสาร!!" & vbNewLine _
                             & "โปรดปิดหน้าจอนี้ แล้วสั่งพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")
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
          spnRpt.SplitterDistance = intWidth           'เท่ากับ Resolution ของเครื่อง

      End If
  End With
End Sub


Private Sub PrePrintData(ByVal strSelectCode As String)

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

     strSqlSelc = "SELECT * " _
                          & " FROM notifyissue (NOLOCK)" _
                          & " WHERE req_id = '" & strSelectCode & "'"


     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

             If .RecordCount <> 0 Then

                 For i As Integer = 1 To .RecordCount

                                       strSqlCmdPic = "SELECT * " _
                                                               & " FROM tmp_notifyissue (NOLOCK)" _
                                                               & " WHERE req_id = '" & strSelectCode & "'"

                                       RsdPic = New ADODB.Recordset
                                       RsdPic.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                                       RsdPic.LockType = ADODB.LockTypeEnum.adLockOptimistic
                                       RsdPic.Open(strSqlCmdPic, Conn, , , )

                                                     RsdPic.AddNew()
                                                     RsdPic.Fields("user_id").Value = frmMainPro.lblLogin.Text.ToString.Trim
                                                     RsdPic.Fields("req_id").Value = .Fields("req_id").Value
                                                     RsdPic.Fields("req_sta").Value = .Fields("req_sta").Value
                                                     RsdPic.Fields("group").Value = .Fields("group").Value
                                                     RsdPic.Fields("to_dep").Value = .Fields("to_dep").Value
                                                     RsdPic.Fields("from_notify").Value = .Fields("from_notify").Value
                                                     RsdPic.Fields("dep_notify").Value = .Fields("dep_notify").Value
                                                     RsdPic.Fields("order").Value = .Fields("order").Value
                                                     RsdPic.Fields("eqpnm").Value = .Fields("eqpnm").Value
                                                     RsdPic.Fields("shoe").Value = .Fields("shoe").Value
                                                     RsdPic.Fields("size").Value = .Fields("size").Value
                                                     RsdPic.Fields("amount").Value = .Fields("amount").Value
                                                     RsdPic.Fields("issue").Value = .Fields("issue").Value
                                                     RsdPic.Fields("cause").Value = .Fields("cause").Value
                                                     RsdPic.Fields("needdate").Value = .Fields("needdate").Value
                                                     RsdPic.Fields("needtime").Value = .Fields("needtime").Value
                                                     RsdPic.Fields("fxissue").Value = .Fields("fxissue").Value
                                                     RsdPic.Fields("wantdate").Value = .Fields("wantdate").Value
                                                     RsdPic.Fields("wanttime").Value = .Fields("wanttime").Value
                                                     RsdPic.Fields("pic_Issue").Value = .Fields("pic_Issue").Value
                                                     RsdPic.Fields("person1_sta").Value = .Fields("person1_sta").Value
                                                     RsdPic.Fields("person1").Value = .Fields("person1").Value
                                                     RsdPic.Fields("person1_date").Value = .Fields("person1_date").Value
                                                     RsdPic.Fields("person2_sta").Value = .Fields("person2_sta").Value
                                                     RsdPic.Fields("person2").Value = .Fields("person2").Value
                                                     RsdPic.Fields("person2_date").Value = .Fields("person2_date").Value
                                                     RsdPic.Fields("person3_sta").Value = .Fields("person3_sta").Value
                                                     RsdPic.Fields("person3").Value = .Fields("person3").Value
                                                     RsdPic.Fields("person3_date").Value = .Fields("person3_date").Value
                                                     RsdPic.Fields("person4_sta").Value = .Fields("person4_sta").Value
                                                     RsdPic.Fields("person4").Value = .Fields("person4").Value
                                                     RsdPic.Fields("person4_date").Value = .Fields("person4_date").Value
                                                     RsdPic.Fields("recordby").Value = .Fields("recordby").Value
                                                     RsdPic.Fields("record_date").Value = .Fields("record_date").Value
                                                     RsdPic.Fields("lastby").Value = .Fields("lastby").Value
                                                     RsdPic.Fields("last_date").Value = .Fields("last_date").Value
                                                     RsdPic.Fields("remark").Value = .Fields("remark").Value


                                                     Dim RsdSteam As New ADODB.Stream
                                                     Dim strPicSign02 As String
                                                     Dim strPicSign03 As String
                                                     Dim strPicSign04 As String

                                                     RsdSteam.Type = StreamTypeEnum.adTypeBinary
                                                     RsdSteam.Open()


                                                     '--------------------------------------โหลดรูปภาพลายเซ็นต์ ผู้อนุมัติแจ้ง -------------------------------------

                                                     If .Fields("person2").Value.ToString.Trim <> "" Then
                                                         strPicSign02 = CallPathSignPicture(.Fields("person2").Value.ToString.Trim)
                                                     Else
                                                         strPicSign02 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                                     End If

                                                     RsdSteam.LoadFromFile(strPicSign02)
                                                     RsdPic.Fields("sign_approve2").Value = RsdSteam.Read


                                                     '--------------------------------------โหลดรูปภาพลายเซ็นต์ ผู้รับแจ้ง -------------------------------------

                                                     If .Fields("person3").Value.ToString.Trim <> "" Then
                                                         strPicSign03 = CallPathSignPicture(.Fields("person3").Value.ToString.Trim)
                                                     Else
                                                         strPicSign03 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                                     End If

                                                     RsdSteam.LoadFromFile(strPicSign03)
                                                     RsdPic.Fields("sign_approve3").Value = RsdSteam.Read

                                                     RsdPic.Update()


                                                     '--------------------------------------โหลดรูปภาพลายเซ็นต์ ผู้อนุมัติรับแจ้ง -------------------------------------

                                                     If .Fields("person4").Value.ToString.Trim <> "" Then
                                                         strPicSign04 = CallPathSignPicture(.Fields("person4").Value.ToString.Trim)
                                                     Else
                                                         strPicSign04 = "\\10.32.0.14\SignPicture\sign_bnk.jpg"
                                                     End If

                                                     RsdSteam.LoadFromFile(strPicSign04)
                                                     RsdPic.Fields("sign_approve4").Value = RsdSteam.Read

                                                     RsdPic.Update()


                                        RsdPic.ActiveConnection = Nothing
                                        RsdPic.Close()
                                        RsdPic = Nothing
                                  .MoveNext()     'เลื่อนไปที่ Record ถัดไป
                  Next i

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

                           '-------------- Clear tmp_notifyissue -------------------------------

                           strSqlCmd = "DELETE FROM tmp_notifyissue " _
                                         & " WHERE req_id = '" & frmMainPro.lblRptDesc.Text & "'"

                           Conn.Execute(strSqlCmd)

          End Select

 Conn.Close()
 Conn = Nothing
End Sub

Private Sub btnAcp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcp1.Click

   If SignCompare() Then

      SaveApprove_st()
      ClearData(0)
      PrePrintData(frmMainPro.lblRptDesc.Text)
      Viewdata()

   Else
      MessageBox.Show("คุณไม่มีสิทธิใช้งานส่วนนี้", "Access denie!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)

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

     strSqlSelc = " SELECT * FROM tmp_notifyissue (NOLOCK)" _
                                   & " WHERE req_id = '" & frmMainPro.lblRptDesc.Text & "'"

     Conn.Execute(strSqlSelc)

       With Rsd

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlSelc, Conn, , , )


            If .RecordCount <> 0 Then

               strDept = .Fields("dep_notify").Value.ToString.Trim       'แผนกที่รับ
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

                           Case Else

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

Private Sub SaveApprove_st()
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

          strSqlCmd = "UPDATE notifyissue SET req_sta = '" & "1" & "'" _
                                       & "," & "person2_sta = '" & True & "'" _
                                       & "," & "person2 = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                       & "," & "person2_date = " & strDate _
                                       & " WHERE req_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

          Conn.Execute(strSqlCmd)

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub btnAcp3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAcp3.Click

  If frmMainPro.lblLogin.Text = "BOONTUM" Then          'ตรวจสอบชื่อผู้ใช้

     If chkApprove_rd() Then       'ตรวจสอบว่า  Approve แล้วหรือยัง

        SaveApprove_rd()           'เพิ่มรายการใน table approve_delv
        ClearData(0)               'เคลี่ยร์ข้อมูลใน tmp_rpt_delv2 

        PrePrintData(frmMainPro.lblRptDesc.Text)
        Viewdata()

     Else
         MessageBox.Show("ผู้รับแจ้งยังไม่ได้ดำเนินการ", "ไม่สามารถดำเนินการได้", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

     End If

  Else
      MessageBox.Show("คุณไม่มีสิทธิใช้งานส่วนนี้", "Access denie!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)

  End If

End Sub

Private Function chkApprove_rd() As Boolean         'ฟังก์ชั่นตรวจสอบการเซ็น approve

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

     strSqlSelc = "SELECT person3 FROM notifyissue (NOLOCK)" _
                          & " WHERE req_id ='" & frmMainPro.lblRptDesc.Text & "' "

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .Fields("person3").Value.ToString.Trim <> "" Then

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

Private Sub SaveApprove_rd()

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

          strSqlCmd = "UPDATE notifyissue SET req_sta = '3' " _
                                       & "," & "person4_sta = '" & True & "'" _
                                       & "," & "person4 = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                       & "," & "person4_date = " & strDate _
                                       & " WHERE req_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

          Conn.Execute(strSqlCmd)

  Conn.Close()
  Conn = Nothing

End Sub

End Class