Imports ADODB

Public Class frmMainPro
Dim intX As Integer = Screen.PrimaryScreen.Bounds.Width
Dim intY As Integer = Screen.PrimaryScreen.Bounds.Height

    Private Sub frmMainPro_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'ChngResltWinwnd(intX, intY)      'คืนค่า Resolution เดิมของเครื่อง(อยู่ใน ResChanger)
        Timer1.Enabled = False
        frmLogin.Close()
        Me.Dispose()
    End Sub

    Private Sub frmMainPro_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'ChngResltWinwnd(intX, intY)      'คืนค่า Resolution เดิมของเครื่อง(อยู่ใน ResChanger)
    End Sub

    Private Sub frmMainPro_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        StdDateTimeThai()
        Me.WindowState = FormWindowState.Maximized

        PreMainBar()
        ImageBackground()  'แสดงรูปกลางหน้าจอ

        With Timer1
            .Interval = 100        '1s = 1,000 ms ( 60,000 ms = 1 minute )
            .Enabled = True
        End With

    End Sub

    Private Sub mnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnExit.Click
        Me.Close()
    End Sub

    Private Sub mnTileHor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnTileHor.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub mnVer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnVer.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub mnCasd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnCasd.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim dteComputer As Date = Now()
        Dim strCurrentDate As String
        Dim strCurrentTime As String

        strCurrentDate = dteComputer.Date.ToString("dd/MM/yyyy")
        strCurrentTime = Format(Now(), "HH:mm:ss")

        lblCurrentDate.Text = "วันที่ : " & strCurrentDate
        lblCurrentTime.Text = "เวลา : " & strCurrentTime

    End Sub

    Private Function CheckFormChildOpen(ByVal strFormName As String) As Boolean

        Dim IsOpenExist As Boolean
        Dim f As Form

        For Each f In Me.MdiChildren

            If f.Name = strFormName Then
                IsOpenExist = True
            End If

        Next

        CheckFormChildOpen = IsOpenExist

    End Function

    Private Sub mnDocFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnDocFile.Click

        Dim strReturnLevel As String
        strReturnLevel = CheckUserLevel(lblLogin.Text.ToString)

        Select Case strReturnLevel

            Case Is = "A" 'สิทธิ Admin

                If CheckFormChildOpen("frmDocFile") Then
                    frmDocFile.Activate()
                Else

                    With frmDocFile
                        .MdiParent = Me
                        .Show()
                    End With

                End If

            Case Is = "U" 'สิทธิ User
                MsnAdmin()

            Case Else
                MsnAdmin()

        End Select

    End Sub

    Private Sub mnUsrFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnUsrFile.Click

        Dim strReturnLevel As String
        strReturnLevel = CheckUserLevel(lblLogin.Text.ToString)

        Select Case strReturnLevel

            Case Is = "A" 'สิทธิ Admin

                If CheckFormChildOpen("frmUserPermit") Then
                    frmUserPermit.Activate()
                Else

                    With frmUserPermit
                        .MdiParent = Me
                        .Show()
                    End With

                End If

            Case Is = "U" 'สิทธิ User
                MsnAdmin()

            Case Else
                MsnAdmin()

        End Select

    End Sub

    Private Sub mnNewLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnNewLogIn.Click
        frmLogin.Show()
        Me.Enabled = False
    End Sub

    Private Sub mnChangePass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnChangePass.Click
        frmChangePass.ShowDialog()
    End Sub

    Private Sub PreMainBar()

        With lstBarMain
            .Groups.Add("ข้อมูลหลัก")
            .Groups.Add("บันทึกรายการ")

            .Groups(0).Items.Add("โมล์ดฉีด,พ่น,หยอด", 0)
            .Groups(0).Items.Add("โมล์ดแผงอัดลาย", 1)
            .Groups(0).Items.Add("มีดตัดชิ้นส่วน", 2)
            .Groups(0).Items.Add("บล็อคสกรีน", 3)
            .Groups(0).Items.Add("บล็อคอาร์ค", 4)

            .Groups(1).Items.Add("โอนอุปกรณ์ลงผลิต", 5)
            .Groups(1).Items.Add("บันทึกส่งซ่อม - รับเข้าอุปกรณ์", 6)

            '-------------------- Set Font Color ---------------------

            .Groups(0).Items(0).ForeColor = Color.White
            .Groups(0).Items(1).ForeColor = Color.White
            .Groups(0).Items(2).ForeColor = Color.White
            .Groups(0).Items(3).ForeColor = Color.White
            .Groups(0).Items(4).ForeColor = Color.White

            .Groups(1).Items(0).ForeColor = Color.White
            .Groups(1).Items(1).ForeColor = Color.White

        End With

    End Sub

    Private Sub lstBarMain_ItemClicked(ByVal sender As Object, ByVal e As vbAccelerator.Components.ListBarControl.ItemClickedEventArgs) Handles lstBarMain.ItemClicked

        Dim blnReturn As Boolean
        Dim strDocCode As String = "F"        'รหัสเอกสาร
        Dim strUserPermiss As String = ""

        Select Case e.Item.IconIndex            'จะต้องตรวจสอบให้รูปภาพต้องไม่ซ้ำกันใน imagelist

            Case Is = 0 '

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim      'F0
                blnReturn = CheckUserEntry(strDocCode, "act_open")            'ตรวจสอบสิทธิ์เข้าใช้ (ไฟล์ไอคอน, ฟิวด์ที่ตรวจสอบ)
                If blnReturn Then

                    If CheckFormChildOpen("frmMoldInj") Then             'เช็คฟอร์มเปิดอยู่หรือไม่
                        frmMoldInj.Activate()

                    Else

                        With frmMoldInj
                            .MdiParent = Me
                            .Show()
                        End With

                    End If

                Else
                    MsnAdmin()
                End If

            Case Is = 1 'แฟ้มโมลด์อัดลายแผง

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim
                blnReturn = CheckUserEntry(strDocCode, "act_open")
                If blnReturn Then

                    If CheckFormChildOpen("frmEqpSheet") Then
                        frmEqpSheet.Activate()

                    Else

                        With frmEqpSheet
                            .MdiParent = Me
                            .Show()
                        End With

                    End If

                Else
                    MsnAdmin()

                End If



            Case Is = 2 'แฟ้มมีดตัดชิ้นส่วน

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim
                blnReturn = CheckUserEntry(strDocCode, "act_open")
                If blnReturn Then

                    If CheckFormChildOpen("frmCutting") Then
                        frmCutting.Activate()

                    Else
                        With frmCutting
                            .MdiParent = Me
                            .Show()

                        End With

                    End If

                Else
                    MsnAdmin()
                End If

            Case Is = 3 'แฟ้มบล็อคสกรีน

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim
                blnReturn = CheckUserEntry(strDocCode, "act_open")
                If blnReturn Then

                    If CheckFormChildOpen("frmScreenBlk") Then
                        frmScreenBlk.Activate()

                    Else

                        With frmScreenBlk
                            .MdiParent = Me
                            .Show()
                        End With

                    End If

                Else
                    MsnAdmin()
                End If


            Case Is = 4 'แฟ้มบล็อคอาร์ค

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim
                blnReturn = CheckUserEntry(strDocCode, "act_open")
                If blnReturn Then

                    If CheckFormChildOpen("frmArkBlk") Then
                        frmArkBlk.Activate()

                    Else

                        With frmArkBlk
                            .MdiParent = Me
                            .Show()
                        End With

                    End If

                Else
                    MsnAdmin()
                End If


            Case Is = 5 'โอนอุปกรณ์ลงฝ่ายผลิต

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim
                blnReturn = CheckUserEntry(strDocCode, "act_open")

                If blnReturn Then

                    If CheckFormChildOpen("frmDelv") Then
                        frmDelv.Activate()

                    Else

                        With frmDelv
                            .MdiParent = Me
                            .Show()
                        End With

                    End If

                Else
                    MsnAdmin()
                End If


            Case Is = 6 'บันทึกส่ง - รับเข้า อุปกรณ์

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim
                blnReturn = CheckUserEntry(strDocCode, "act_open")

                If blnReturn Then

                    If CheckFormChildOpen("frmFixEqpmnt") Then
                        frmFixEqpmnt.Activate()

                    Else

                        With frmFixEqpmnt
                            .MdiParent = Me
                            .Show()
                        End With

                    End If

                Else
                    MsnAdmin()
                End If

            Case Is = 10  'แจ้งปัญหาอุปกรณ์

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim
                blnReturn = CheckUserEntry(strDocCode, "act_open")

                If blnReturn Then


                    If CheckFormChildOpen("frmNotifyIssue") Then
                        frmNotifyIssue.Activate()

                    Else

                        With frmNotifyIssue
                            .MdiParent = Me
                            .Show()

                        End With

                    End If

                Else
                    MsnAdmin()

                End If

            Case Is = 11  'แจ้งปัญหาอุปกรณ์

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim

                blnReturn = CheckUserEntry(strDocCode, "act_open")

                If blnReturn Then

                    If CheckFormChildOpen("frmApproveIssue") Then
                        frmApproveIssue.Activate()

                    Else

                        With frmApproveIssue
                            .MdiParent = Me
                            .Show()

                        End With

                    End If

                Else
                    MsnAdmin()

                End If


            Case Is = 12  'แฟ้มข้อมูลอนุมัติการโอนอุปกรณ์ลงผลิต

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim

                blnReturn = CheckUserEntry(strDocCode, "act_open")

                If blnReturn Then


                    If CheckFormChildOpen("frmApproveDelv") Then
                        frmApproveDelv.Activate()

                    Else

                        With frmApproveDelv
                            .MdiParent = Me
                            .Show()

                        End With

                    End If

                Else
                    MsnAdmin()

                End If
        End Select

    End Sub

    Private Sub mnRptdvl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnRptdvl.Click
        frmRptEqpTrnsf.Show()
    End Sub

    Private Sub mnWipnewImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnWipnewImport.Click
        frmWipImp.ShowDialog()
    End Sub

End Class
