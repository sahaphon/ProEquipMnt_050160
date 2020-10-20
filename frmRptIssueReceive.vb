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

Public Class frmRptIssueReceive
 Dim cryRpt As New ReportDocument
 Dim strUser As String
 Dim intWidth As Integer

Private Sub frmRptIssueReceive_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearData(0)
  frmMainPro.lblRptDesc.Text = ""
  Me.Dispose()

  frmMainPro.Show()
  frmNotifyIssue.Show()

End Sub

Private Sub frmRptIssueReceive_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Me.WindowState = FormWindowState.Maximized
  StdDateTimeThai()        'ซับรูทีนเเปลงวันที่เป็นวดป.ไทย อยู่ใน Module

  PrePrintData(frmMainPro.lblRptDesc.Text)
  HidePanel()
  Viewdata()

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

                If .Fields("person3").Value.ToString <> "" Then
                    btnAcp.Enabled = False
                    lblApp.Visible = True

                End If

                ds.Clear()
                da.Fill(ds, Rsd, "issue")

                cryRpt.Load(Application.StartupPath & "\ReceiveIssue.rpt")
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

         If chkApprove() Then           'ตรวจสอบว่าผจก.แผนกที่แจ้ง Approve แล้วหรือยัง


                    If chkApprove03() Then
                       gpbRecvNotify.Visible = False

                    Else

                         If frmMainPro.lblLogin.Text = "SUTID" Then
                            gpbRecvNotify.Visible = True
                            txtFxissue.Focus()

                         End If

                    End If

         Else
             gpbRecvNotify.Visible = False

         End If

      Else
          .Text = ">"
          spnRpt.SplitterDistance = intWidth           'เท่ากับ Resolution ของเครื่อง

      End If
  End With

End Sub

Private Sub btnAcp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAcp.Click

   If frmMainPro.lblLogin.Text = "SUTID" Then          'ตรวจสอบชื่อผู้ใช้

       If chkApprove() Then           'ตรวจสอบว่าผจก.แผนกที่แจ้ง Approve แล้วหรือยัง
            gpbRecvNotify.Visible = True

              If chkApprove03() Then
                 PrePrintData(frmMainPro.lblRptDesc.Text)
                 Viewdata()
                 btnAcp.Enabled = False
                 lblApp.Visible = True
                 gpbRecvNotify.Visible = False

              Else
                  btnAcp.Enabled = True
                  gpbRecvNotify.Visible = True
                  txtFxissue.Focus()

                      If txtFxissue.Text <> "" Then
                         SaveApprove()
                         ClearData(0)

                         PrePrintData(frmMainPro.lblRptDesc.Text)
                         Viewdata()
                         ClearDataInput()        'เคลี่ยร์ข้อมูลรายการแก้ไข
                         btnAcp.Enabled = False
                         gpbRecvNotify.Visible = False

                      Else
                         MsgBox("กรุณาระบุ รายละเอียดการแก้ไข")
                         txtFxissue.Focus()
                      End If

              End If

         Else
            gpbRecvNotify.Visible = False
            MessageBox.Show("แผนกที่แจ้งยังไม่ได้อนุมัติ...", "ไม่สามารถดำเนินการได้", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

         End If


    Else
      MessageBox.Show("คุณไม่มีสิทธิใช้งานส่วนนี้", "Access denie!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End If

End Sub

Sub ClearDataInput()
  txtFxissue.Text = ""
  txtWanttime.Text = ""
  txtWantDate.Text = ""
End Sub

Private Function chkApprove() As Boolean         'ฟังก์ชั่นตรวจสอบการเซ็น approve

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

     strSqlSelc = "SELECT person2_sta FROM notifyissue (NOLOCK)" _
                             & " WHERE req_id ='" & frmMainPro.lblRptDesc.Text & "' "

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .Fields("person2_sta").Value Then

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

Private Function chkApprove03() As Boolean         'ฟังก์ชั่นตรวจสอบการเซ็น approve

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

     strSqlSelc = "SELECT person3_sta FROM notifyissue (NOLOCK)" _
                             & " WHERE req_id ='" & frmMainPro.lblRptDesc.Text & "' "

     Rsd = New ADODB.Recordset

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlSelc, Conn, , , )

          If .Fields("person3_sta").Value Then

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

Private Sub SaveApprove()

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmd As String

 Dim datSave As Date = Now()
 Dim strDate As String = datSave.Date.ToString("yyyy-MM-dd")
 Dim strWantDate As String

     strDate = "'" & SaveChangeEngYear(strDate) & "'"

     With Conn

          If .State Then Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()

     End With

            '---------------------------------------- วดป.เปิดใบสั่งซื้อ -------------------------------------------------
                   If txtWantDate.Text <> "__/__/____" Then

                      strWantDate = Mid(txtWantDate.Text.ToString, 7, 4) & "-" _
                                    & Mid(txtWantDate.Text.ToString, 4, 2) & "-" _
                                    & Mid(txtWantDate.Text.ToString, 1, 2)
                      strWantDate = "'" & SaveChangeEngYear(strWantDate) & "'"

                   Else
                      strWantDate = "NULL"

                   End If


          strSqlCmd = "UPDATE notifyissue SET req_sta = '2'" _
                                       & "," & "fxissue = '" & ReplaceQuote(txtFxissue.Text) & "'" _
                                       & "," & "wantdate = " & strWantDate _
                                       & "," & "wanttime = '" & ReplaceQuote(txtWanttime.Text) & "'" _
                                       & "," & "person3_sta =  '" & True & "'" _
                                       & "," & "person3 = '" & ReplaceQuote(frmMainPro.lblLogin.Text) & "'" _
                                       & "," & "person3_date = " & strDate _
                                       & " WHERE req_id ='" & ReplaceQuote(frmMainPro.lblRptDesc.Text) & "'"

          Conn.Execute(strSqlCmd)

  Conn.Close()
  Conn = Nothing

End Sub

Private Sub txtFxissue_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFxissue.KeyDown

  Dim intChkpoint As Integer

        With txtFxissue

            Select Case e.KeyCode

                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย
                   Case Is = 38 'ปุ่มลูกศรขึ้น  
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

Private Sub txtFxissue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFxissue.KeyPress
 If e.KeyChar = Chr(13) Then
    txtWantDate.Focus()

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
                           txtFxissue.Focus()
                        End If

                   Case Is = 38 'ปุ่มลูกศรขึ้น  
                        txtFxissue.Focus()

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

Private Sub txtWanttime_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWanttime.KeyPress
 If e.KeyChar = Chr(13) Then
    btnAcp.Focus()
 End If
End Sub

Private Sub spnRpt_Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles spnRpt.Panel2.Paint

End Sub
End Class