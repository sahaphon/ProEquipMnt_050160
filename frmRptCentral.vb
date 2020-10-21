Imports ADODB
Imports System.Data.OleDb.OleDbDataAdapter
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared


Public Class frmRptCentral

Dim cryRpt As New ReportDocument

Dim strUser As String
Dim strPfsID As String

Dim strPrintCode As String
Dim strDataReport As String
Dim strDataReportEx As String

Private Sub frmRptCentral_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    With frmMainPro
         .lblRptCentral.Text = ""
         .lblRptDesc.Text = ""
         .Show()
    End With

    ClearTmpTableUser("tmp_eqpmst")
    ClearTmpTableUser("tmp_mst_trn")
    ClearTmpTableUser("tmp_v_fixeqptrn")
    ClearTmpTableUser("tmp_eqptrn")
    ClearTmpTableUser("print_view_allmold")

    Me.Dispose()
End Sub

    Private Sub frmRptCentral_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        StdDateTimeThai()
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        strUser = frmMainPro.lblLogin.Text.Trim.ToString 'ใช้ User

        With frmMainPro

            strDataReport = .lblRptDesc.Text.ToString.Trim
            Select Case .lblRptCentral.Text

                Case Is = "A" 'ใบตรวจสอบอุปกรณ์

                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    InputMoldInjData()

                Case Is = "B" 'ใบโอนอุปกรณ์
                    CRviewer.ShowPrintButton = False
                    btnPrint.Visible = True
                    InputBillDeliverData()

                Case Is = "C" 'โมลด์อัดลาย
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    InputEqpSheetmoldData()

                Case Is = "D" 'รายงานบล็อคสกรีน
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    InputBlockScreenData()

                Case Is = "E" 'รายงานบล็อคอาร์ค
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    InputBlockArkData()

                Case Is = "F"  'รายงานส่งซ่อมอุปกรณ์
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    inputFixEqpData()

                Case Is = "G"  'รายงานรับกลับส่งซ่อมอุปกรณ์
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    inputRecvFixEqp()

                Case Is = "H"  'รายงานส่งซ่อม - รับกลับส่งซ่อม
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    FixAndRecv()

                Case Is = "I" 'รายการ Mold ทั้งหมด
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    PrintAllMold()

                Case Is = "J" 'รายการ Mold อัดลาย
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    PrintEqpMold()

                Case Is = "K" 'รายงานโมล์ดทั้งหมด คุณชลินทร์
                    CRviewer.ShowPrintButton = True
                    btnPrint.Visible = False
                    PrepairReport()

            End Select

        End With

    End Sub

    Sub PrepairReport()
        Dim Conn As New ADODB.Connection
        Dim RsdDvl As New ADODB.Recordset
        Dim strSqlCmdSelc As String

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

        strSqlCmdSelc = "SELECT * FROM v_molds (NOLOCK) ORDER BY moldtype, eqp_id"

        RsdDvl = New ADODB.Recordset

        With RsdDvl

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlCmdSelc, Conn, , , )

            If .RecordCount <> 0 Then

                ds.Clear()
                da.Fill(ds, RsdDvl, "qty_cd")

                cryRpt.Load(Application.StartupPath & "\rptMolds.rpt")
                cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                cryRpt.ReportOptions.EnableSaveDataWithReport = False
                cryRpt.SetDataSource(ds.Tables("qty_cd"))

                'Dim cryTextDoc As CrystalDecisions.CrystalReports.Engine.TextObject
                Dim cryTxtUsr As CrystalDecisions.CrystalReports.Engine.TextObject

                'cryTextDoc = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryTxtDoc")
                cryTxtUsr = cryRpt.ReportDefinition.Sections(1).ReportObjects("cryuser")

                'cryTextDoc.Text = strDataReportEx
                cryTxtUsr.Text = strUser

                '------------------------------ กำหนดขนาดกระดาษเองโดยตั้งขนาดที่เครื่อง Client ก่อน ---------------------------------------

                Dim printDoc As New System.Drawing.Printing.PrintDocument
                Dim pkSize As PaperSize

                Dim strNewPaper As String = "FolderControl_20x14" 'ขนาดจริงคือ 20.40x14.00 cm ใช้หน่วย Mertric 
                'Dim strNewPaper As String = "PaperTest"

                Dim sngPaperW As Single = 204 'ซม.
                Dim sngPaperH As Single = 140 'ซม.

                Dim strFindNewPaper As String

                Dim i As Integer
                Dim x As Byte

                For i = 0 To printDoc.PrinterSettings.PaperSizes.Count - 1

                    strFindNewPaper = printDoc.PrinterSettings.PaperSizes.Item(i).PaperName
                    If strNewPaper = strFindNewPaper Then
                        pkSize = printDoc.PrinterSettings.PaperSizes.Item(i).RawKind
                        x = 1
                        Exit For
                    End If

                Next i


                If x = 1 Then
                    cryRpt.PrintOptions.PaperSize = CType(pkSize, CrystalDecisions.Shared.PaperSize)
                Else
                    cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4
                End If

                CRviewer.ReportSource = cryRpt
                CRviewer.DisplayStatusBar = True
                CRviewer.Refresh()
                CRviewer.Zoom(100)

            Else

                MsgBox("ไม่มีข้อมูลที่คุณต้องการค้นหา_2!!" & vbNewLine _
                          & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

            End If

            .ActiveConnection = Nothing
            ' .Close()

        End With
        RsdDvl = Nothing

        ds.Clear()
        ds.Dispose()

        da.Dispose()

        ds = Nothing
        da = Nothing


        Conn.Close()
        Conn = Nothing

    End Sub

    Private Sub InputMoldInjData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset
Dim RsdSub As New ADODB.Recordset

Dim strSqlCmdSelc As String
Dim strSqlSelcSub As String

Dim da As New System.Data.OleDb.OleDbDataAdapter
Dim adap As New System.Data.OleDb.OleDbDataAdapter
Dim ds As New DataSet
Dim dsSub As New DataSet


    With Conn

         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .CommandTimeout = 30
            .Open()
    End With

       strSqlCmdSelc = "SELECT * FROM v_tmp_eqpmst (NOLOCK)" _
                                     & " WHERE " & strDataReport

       With Rsd

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , , )

               If .RecordCount <> 0 Then

                  ds.Clear()
                  da.Fill(ds, Rsd, "qty_cd")

                  cryRpt.Load(Application.StartupPath & "\BillChkLst.rpt")
                  cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                  cryRpt.ReportOptions.EnableSaveDataWithReport = False
                  cryRpt.SetDataSource(ds.Tables("qty_cd"))

                  '------------------ Set DataSource Subreports ------------------------

                  strSqlSelcSub = "SELECT * FROM pre_tmp_eqptrn_newsize (NOLOCK)" _
                                                   & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                                                   & " ORDER BY size_desc, tmp_newsize"

                  RsdSub = New ADODB.Recordset
                  RsdSub.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                  RsdSub.LockType = ADODB.LockTypeEnum.adLockOptimistic
                  RsdSub.Open(strSqlSelcSub, Conn, , , )

                  dsSub.Clear()
                  adap.Fill(dsSub, RsdSub, "eqp_id")
                  cryRpt.OpenSubreport("sbChkList").SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                  cryRpt.OpenSubreport("sbChkList").SetDataSource(dsSub.Tables("eqp_id"))

                  RsdSub.ActiveConnection = Nothing
                  RsdSub = Nothing

                  cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4
                  cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Landscape

                  CRviewer.ReportSource = cryRpt
                  CRviewer.DisplayStatusBar = True
                  CRviewer.Refresh()
                  CRviewer.Zoom(80)

               Else

                    MsgBox("ไม่มีข้อมูลที่คุณต้องการค้นหา!!" & vbNewLine _
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


Private Sub InputBillDeliverData()

        Dim Conn As New ADODB.Connection
        Dim RsdDvl As New ADODB.Recordset
        Dim strSqlCmdSelc As String

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

        strSqlCmdSelc = "SELECT * FROM  v_rpt_delvr (NOLOCK)" _
                                     & " WHERE " & strDataReport

        RsdDvl = New ADODB.Recordset

        With RsdDvl

            .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            .LockType = ADODB.LockTypeEnum.adLockOptimistic
            .Open(strSqlCmdSelc, Conn, , , )

            If .RecordCount <> 0 Then

                ds.Clear()
                da.Fill(ds, RsdDvl, "qty_cd")

                cryRpt.Load(Application.StartupPath & "\BillDelvr.rpt")
                cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                cryRpt.ReportOptions.EnableSaveDataWithReport = False
                cryRpt.SetDataSource(ds.Tables("qty_cd"))

                'Dim cryTextDoc As CrystalDecisions.CrystalReports.Engine.TextObject
                Dim cryTxtUsr As CrystalDecisions.CrystalReports.Engine.TextObject

                'cryTextDoc = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryTxtDoc")
                cryTxtUsr = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryTxtUsr")

                'cryTextDoc.Text = strDataReportEx
                cryTxtUsr.Text = strUser

                '------------------------------ กำหนดขนาดกระดาษเองโดยตั้งขนาดที่เครื่อง Client ก่อน ---------------------------------------

                Dim printDoc As New System.Drawing.Printing.PrintDocument
                Dim pkSize As PaperSize

                Dim strNewPaper As String = "FolderControl_20x14" 'ขนาดจริงคือ 20.40x14.00 cm ใช้หน่วย Mertric 
                'Dim strNewPaper As String = "PaperTest"

                Dim sngPaperW As Single = 204 'ซม.
                Dim sngPaperH As Single = 140 'ซม.

                Dim strFindNewPaper As String

                Dim i As Integer
                Dim x As Byte

                For i = 0 To printDoc.PrinterSettings.PaperSizes.Count - 1

                    strFindNewPaper = printDoc.PrinterSettings.PaperSizes.Item(i).PaperName
                    If strNewPaper = strFindNewPaper Then
                        pkSize = printDoc.PrinterSettings.PaperSizes.Item(i).RawKind
                        x = 1
                        Exit For
                    End If

                Next i


                If x = 1 Then
                    cryRpt.PrintOptions.PaperSize = CType(pkSize, CrystalDecisions.Shared.PaperSize)
                Else
                    cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4
                End If

                CRviewer.ReportSource = cryRpt
                CRviewer.DisplayStatusBar = True
                CRviewer.Refresh()
                CRviewer.Zoom(100)

            Else

                MsgBox("ไม่มีข้อมูลที่คุณต้องการค้นหา_2!!" & vbNewLine _
                          & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

            End If

            .ActiveConnection = Nothing
            ' .Close()

        End With
        RsdDvl = Nothing

    ds.Clear()
    ds.Dispose()

        da.Dispose()

    ds = Nothing
    da = Nothing


    Conn.Close()
    Conn = Nothing

    End Sub


Private Sub InputEqpSheetmoldData()  'โมล์ดอัดลาย

   Dim Conn As New ADODB.Connection
   Dim RsdPst As New ADODB.Recordset
   Dim strSqlCmdSelc As String

   Dim da As New System.Data.OleDb.OleDbDataAdapter
   Dim ds As New DataSet

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .CommandTimeout = 30
               .Open()

       End With

       strSqlCmdSelc = " SELECT * FROM printEqpMold" _
                                      & " WHERE " & strDataReport _
                                      & " ORDER BY size_id"


       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "eqpSheet")

                     cryRpt.Load(Application.StartupPath & "\Eqpsheet.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("eqpSheet"))

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                          'set ขนาดกระดาษ
                cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Portrait           'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                              & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing


 Conn.Close()
 Conn = Nothing

End Sub

Private Sub InputBlockScreenData()
   Dim Conn As New ADODB.Connection
   Dim RsdPst As New ADODB.Recordset
   Dim strSqlCmdSelc As String

   Dim da As New System.Data.OleDb.OleDbDataAdapter
   Dim ds As New DataSet

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .CommandTimeout = 30
               .Open()

       End With

       strSqlCmdSelc = " SELECT * FROM tmp_mst_trn(NOLOCK) " _
                                      & " WHERE " & strDataReport   'WHERE Userid AND Eqpid  

       RsdPst = New ADODB.Recordset

       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "blk_screen")

                     cryRpt.Load(Application.StartupPath & "\BlockScreen.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("blk_screen"))

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Portrait           'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                              & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing

 Conn.Close()
 Conn = Nothing
End Sub

Private Sub InputBlockArkData()
  Dim Conn As New ADODB.Connection
   Dim RsdPst As New ADODB.Recordset
   Dim strSqlCmdSelc As String

   Dim da As New System.Data.OleDb.OleDbDataAdapter
   Dim ds As New DataSet

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .CommandTimeout = 30
               .Open()

       End With

       strSqlCmdSelc = " SELECT * FROM tmp_mst_trn(NOLOCK) " _
                                      & " WHERE " & strDataReport _
                                      & " ORDER BY size_id"

       RsdPst = New ADODB.Recordset
       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "blk_ark")

                     cryRpt.Load(Application.StartupPath & "\blockArk.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("blk_ark"))

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Portrait           'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                              & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing

 Conn.Close()
 Conn = Nothing
End Sub

Private Sub inputFixEqpData()

   Dim Conn As New ADODB.Connection
   Dim RsdPst As New ADODB.Recordset
   Dim strSqlCmdSelc As String

   Dim da As New System.Data.OleDb.OleDbDataAdapter
   Dim ds As New DataSet

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .CommandTimeout = 30
               .Open()

       End With

       strSqlCmdSelc = " SELECT * FROM tmp_v_fixeqptrn(NOLOCK) "

       RsdPst = New ADODB.Recordset

       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "fixeqp")

                     cryRpt.Load(Application.StartupPath & "\fxreport.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("fixeqp"))

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Landscape            'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                                 & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing

 Conn.Close()
 Conn = Nothing

End Sub

Private Sub FixAndRecv()   'รายงาน ส่งซ่อม + รับเข้า อุปกรณ์
   Dim Conn As New ADODB.Connection
   Dim RsdPst As New ADODB.Recordset
   Dim strSqlCmdSelc As String

   Dim da As New System.Data.OleDb.OleDbDataAdapter
   Dim ds As New DataSet

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .CommandTimeout = 30
               .Open()

       End With

       strSqlCmdSelc = " SELECT * FROM tmp_v_fixeqptrn(NOLOCK) "

       RsdPst = New ADODB.Recordset
       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "rcvfix")

                     cryRpt.Load(Application.StartupPath & "\RFeqpmnt.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("rcvfix"))

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Landscape            'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                              & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing


 Conn.Close()
 Conn = Nothing


End Sub

Private Sub inputRecvFixEqp()

   Dim Conn As New ADODB.Connection
   Dim RsdPst As New ADODB.Recordset
   Dim strSqlCmdSelc As String

   Dim da As New System.Data.OleDb.OleDbDataAdapter
   Dim ds As New DataSet

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .CommandTimeout = 30
               .Open()

       End With

       strSqlCmdSelc = " SELECT * FROM tmp_v_fixeqptrn(NOLOCK) "

       RsdPst = New ADODB.Recordset
       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "rcvfix")

                     cryRpt.Load(Application.StartupPath & "\rcvreport.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("rcvfix"))

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Landscape            'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                              & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing


 Conn.Close()
 Conn = Nothing

End Sub

Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

    With frmMainPro

            Select Case .lblRptCentral.Text

                      Case Is = "A" 'ใบตรวจสอบอุปกรณ์                              
                      Case Is = "B" 'ใบโอนอุปกรณ์ paper
                                PrintHalfPaper() 'พิมพ์กระดาษครึ่งหนึ่ง                                
                      Case Is = "C"
                      Case Is = "D"
                      Case Else

            End Select

  End With

End Sub

Private Sub PrintHalfPaper()
 Dim printDoc As New System.Drawing.Printing.PrintDocument
  Dim pkSize As Printing.PaperKind
  Dim strNewPaper As String = "FolderControl_20x14" 'ขนาดจริงคือ 20.40x14.00 cm ใช้หน่วย Mertric 

  'Dim strNewPaper As String = "PaperTest"
  Dim sngPaperW As Single = 204 'ซม.
  Dim sngPaperH As Single = 140 'ซม.
  Dim strFindNewPaper As String

  Dim i As Integer
  Dim x As Byte

                        'CustomPrintForm.AddCustomPaperSizeToDefaultPrinter(strNewPaper, sngPaperW, sngPaperH)

                        For i = 0 To printDoc.PrinterSettings.PaperSizes.Count - 1
                              strFindNewPaper = printDoc.PrinterSettings.PaperSizes.Item(i).PaperName
                               If strNewPaper = strFindNewPaper Then
                                      pkSize = printDoc.PrinterSettings.PaperSizes.Item(i).RawKind
                                       x = 1
                                       Exit For
                               End If

                        Next i

                       'Dim MyNormalFont As New Font("Arial", 8, FontStyle.Regular)
                        Dim PaperK As New Printing.PaperKind
                        Dim PaperS As New Printing.PaperSize

                              PaperK = pkSize
                              PaperS.RawKind = PaperK


                            'printDoc.DefaultPageSettings.Landscape = True

                            printDoc.DefaultPageSettings.PaperSize = PaperS

                            Dim pageDialog1 As New PageSetupDialog ' This Dialog can set the paper size or kind
                            Dim printDialog1 As New PrintDialog ' This is the dialog to setting the printer options

                            pageDialog1.Document = printDoc
                            pageDialog1.PageSettings.PaperSize = PaperS

                            printDialog1.Document = pageDialog1.Document

                            Dim SetPrinterName As Boolean = True

                            While SetPrinterName

                                If printDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

                                    With printDialog1.PrinterSettings

                                        'cryRpt.PrintOptions.PaperSize = CType(pkSize, CrystalDecisions.Shared.PaperSize)
                                        'cryRpt.PrintOptions.PaperSource = PaperSource.Manual

                                        .DefaultPageSettings.PaperSize = PaperS
                                        .DefaultPageSettings.PaperSource.SourceName = strNewPaper

                                        cryRpt.PrintOptions.PrinterName = .PrinterName
                                        cryRpt.PrintOptions.PaperSize = GetPapersizeID(.PrinterName.ToString, strNewPaper)

                                        'ที่สำหรับต้องปรับ Feed กระดาษที่เครื่องพิมพ์ด้วย ตาที่ระบุขนาดไว้

                                        InputBillDeliverData()
                                        cryRpt.PrintToPrinter(.Copies, .Collate, .FromPage, .ToPage)
                                        SetPrinterName = False 'Done

                                    End With

                                    'Printed = True

                                Else
                                    SetPrinterName = False
                                End If

                        End While

End Sub

Public Function GetPapersizeID(ByVal PrinterName As String, ByVal PaperSizeName As String) As Integer
     Dim doctoprint As New System.Drawing.Printing.PrintDocument()
     Dim PaperSizeID As Integer = 0
     Dim ppname As String = ""
     Dim s As String = ""

     doctoprint.PrinterSettings.PrinterName = PrinterName  '(ex. "Epson SQ-1170 ESC/P 2")
     For i As Integer = 0 To doctoprint.PrinterSettings.PaperSizes.Count - 1
         Dim rawKind As Integer

         ppname = PaperSizeName

         If doctoprint.PrinterSettings.PaperSizes(i).PaperName = ppname Then
            rawKind = CInt(doctoprint.PrinterSettings.PaperSizes(i).GetType().GetField("kind", _
                Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic).GetValue(doctoprint.PrinterSettings.PaperSizes(i)))
                PaperSizeID = rawKind
                Exit For
         End If
     Next
     Return PaperSizeID

End Function

Private Sub PrintAllMold()

   Dim Conn As New ADODB.Connection
   Dim RsdPst As New ADODB.Recordset
   Dim strSqlCmdSelc As String

   Dim da As New System.Data.OleDb.OleDbDataAdapter
   Dim ds As New DataSet
       Try

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

       strSqlCmdSelc = "SELECT * FROM print_view_allmold " _
                           & " ORDER BY eqp_id,size_desc,tmp_size"

       RsdPst = New ADODB.Recordset

       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "allmold")

                     cryRpt.Load(Application.StartupPath & "\printAllMold.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("allmold"))

                     'Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject           'ตัวแปรส่งค่าให้กับ CrystalReport user
                     'cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     'cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Portrait            'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                              & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing


 Conn.Close()
 Conn = Nothing


       Catch ex As Exception
             MsgBox(ex.Message)
       End Try
End Sub

Private Sub PrintEqpMold()

  Dim Conn As New ADODB.Connection
  Dim RsdPst As New ADODB.Recordset
  Dim strSqlCmdSelc As String

  Dim da As New System.Data.OleDb.OleDbDataAdapter
  Dim ds As New DataSet

      Try

       With Conn
            If .State Then Close()
               .ConnectionString = strConnAdodb
               .CursorLocation = ADODB.CursorLocationEnum.adUseClient
               .ConnectionTimeout = 90
               .Open()

       End With

       strSqlCmdSelc = "SELECT * FROM print_view_allmold " _
                           & " WHERE [group]='D' " _
                           & " ORDER BY eqp_id,size_desc,tmp_size"

       RsdPst = New ADODB.Recordset

       With RsdPst

               .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
               .LockType = ADODB.LockTypeEnum.adLockOptimistic
               .Open(strSqlCmdSelc, Conn, , )

                 If .RecordCount <> 0 Then

                     ds.Clear()
                     da.Fill(ds, RsdPst, "EQPMold")

                     cryRpt.Load(Application.StartupPath & "\printEqpMold.rpt")
                     cryRpt.SetDatabaseLogon("Sa", "Sa2008", "ADDASRV03", "DBequipmnt")
                     cryRpt.ReportOptions.EnableSaveDataWithReport = False
                     cryRpt.SetDataSource(ds.Tables("EQPMold"))

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4                           'set ขนาดกระดาษ
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Portrait            'กำหนดกระดาษเเนวนอน

                     CRviewer.ReportSource = cryRpt
                     CRviewer.DisplayStatusBar = True
                     CRviewer.Refresh()
                     CRviewer.Zoom(100)

                 Else

                     MsgBox("ไม่มีข้อมูลที่ต้องการพิมพ์!!" & vbNewLine _
                              & "โปรดปิดหน้าจอนี้ แล้วเลือกพิมพ์ใหม่!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                 End If
                 .ActiveConnection = Nothing


       End With
       RsdPst = Nothing

 ds.Clear()
 ds.Dispose()

 da.Dispose()
 ds = Nothing
 da = Nothing


 Conn.Close()
 Conn = Nothing


       Catch ex As Exception
             MsgBox(ex.Message)
       End Try
End Sub

End Class