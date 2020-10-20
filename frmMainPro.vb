Imports ADODB

Public Class frmMainPro
Dim intX As Integer = Screen.PrimaryScreen.Bounds.Width
Dim intY As Integer = Screen.PrimaryScreen.Bounds.Height

    Private Sub frmMainPro_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'ChngResltWinwnd(intX, intY)      '�׹��� Resolution ����ͧ����ͧ(����� ResChanger)
        Timer1.Enabled = False
        frmLogin.Close()
        Me.Dispose()
    End Sub

    Private Sub frmMainPro_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'ChngResltWinwnd(intX, intY)      '�׹��� Resolution ����ͧ����ͧ(����� ResChanger)
    End Sub

    Private Sub frmMainPro_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        StdDateTimeThai()
        Me.WindowState = FormWindowState.Maximized

        PreMainBar()
        ImageBackground()  '�ʴ��ٻ��ҧ˹�Ҩ�

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

        lblCurrentDate.Text = "�ѹ��� : " & strCurrentDate
        lblCurrentTime.Text = "���� : " & strCurrentTime

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

            Case Is = "A" '�Է�� Admin

                If CheckFormChildOpen("frmDocFile") Then
                    frmDocFile.Activate()
                Else

                    With frmDocFile
                        .MdiParent = Me
                        .Show()
                    End With

                End If

            Case Is = "U" '�Է�� User
                MsnAdmin()

            Case Else
                MsnAdmin()

        End Select

    End Sub

    Private Sub mnUsrFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnUsrFile.Click

        Dim strReturnLevel As String
        strReturnLevel = CheckUserLevel(lblLogin.Text.ToString)

        Select Case strReturnLevel

            Case Is = "A" '�Է�� Admin

                If CheckFormChildOpen("frmUserPermit") Then
                    frmUserPermit.Activate()
                Else

                    With frmUserPermit
                        .MdiParent = Me
                        .Show()
                    End With

                End If

            Case Is = "U" '�Է�� User
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
            .Groups.Add("��������ѡ")
            .Groups.Add("�ѹ�֡��¡��")

            .Groups(0).Items.Add("���촩մ,��,��ʹ", 0)
            .Groups(0).Items.Add("����ἧ�Ѵ���", 1)
            .Groups(0).Items.Add("�մ�Ѵ�����ǹ", 2)
            .Groups(0).Items.Add("���ͤʡ�չ", 3)
            .Groups(0).Items.Add("���ͤ����", 4)

            .Groups(1).Items.Add("�͹�ػ�ó�ŧ��Ե", 5)
            .Groups(1).Items.Add("�ѹ�֡�觫��� - �Ѻ����ػ�ó�", 6)

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
        Dim strDocCode As String = "F"        '�����͡���
        Dim strUserPermiss As String = ""

        Select Case e.Item.IconIndex            '�е�ͧ��Ǩ�ͺ����ٻ�Ҿ��ͧ����ӡѹ� imagelist

            Case Is = 0 '

                strDocCode = strDocCode & e.Item.IconIndex.ToString.Trim      'F0
                blnReturn = CheckUserEntry(strDocCode, "act_open")            '��Ǩ�ͺ�Է�������� (����ͤ͹, ��Ǵ����Ǩ�ͺ)
                If blnReturn Then

                    If CheckFormChildOpen("frmMoldInj") Then             '�礿�����Դ�����������
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

            Case Is = 1 '�����Ŵ��Ѵ���ἧ

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



            Case Is = 2 '����մ�Ѵ�����ǹ

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

            Case Is = 3 '������ͤʡ�չ

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


            Case Is = 4 '������ͤ����

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


            Case Is = 5 '�͹�ػ�ó�ŧ���¼�Ե

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


            Case Is = 6 '�ѹ�֡�� - �Ѻ��� �ػ�ó�

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

            Case Is = 10  '�駻ѭ���ػ�ó�

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

            Case Is = 11  '�駻ѭ���ػ�ó�

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


            Case Is = 12  '���������͹��ѵԡ���͹�ػ�ó�ŧ��Ե

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
