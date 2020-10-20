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
    Me.Dispose() '����� close Form ������¿����������� ���Ǥ׹������˹��¤�����
End Sub

Private Sub frmRptEqpTrnsf_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlCmdSelc As String

  Dim da As New System.Data.OleDb.OleDbDataAdapter
  Dim ds As New DataSet


      Me.WindowState = FormWindowState.Maximized
      strUser = frmMainPro.lblLogin.Text.Trim.ToString '�� User


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
                         '------------------- ��������ػ�ó�ŧ� cboEqpid -------------------
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

Private Function StartDate() As String    '�駡����ŧ�� �.�.(DatetimePicker_ �����)

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

Private Function EndDate() As String   '�駡����ŧ�� �.�. (DatetimePicker_�֧)
      Dim Edate, Edate2, Edate3, DateEn As String

      Edate = Mid(dtpEnd.Value.ToString("yyyy-MM-dd"), 1, 4)  '�Ѵʵ�ԧ���੾�л�
      Edate2 = Mid(dtpEnd.Value.ToString("yyyy-MM-dd"), 6, 2) '�Ѵʵ�ԧ���੾����͹
      Edate3 = Mid(dtpEnd.Value.ToString("yyyy-MM-dd"), 9, 2) '�Ѵʵ�ԧ���੾���ѹ���

      If CInt(Edate) > 2500 Then       '��ҷ�����繻� �.�.
             Edate = CInt(Edate) - 543
            DateEn = Edate & "-" & Edate2 & "-" & Edate3

      Else                             ' ��ҷ�����繻� �.�.
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
Dim SDThai, EDThai As String   '��������ʵ�ԧ�ѹ���ҡ�ѧ���� StartDate() ��� EndDate()


        With Conn

            If .State Then .Close()

                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .CommandTimeout = 30
                .Open()

        End With
                    '-------------------------�����������(�ó����͡�ء�ػ�ó� ��� ���͡�ء�ѹ���)---------------------------

                    If ChkAllEqp.Checked = True And ChkTime.Checked = True Then

                                 strSqlCmdSelc = "SELECT * FROM v_rpt_delvr (NOLOCK)"

                    '-------------------------�����������(�ó����͡�ء�ػ�ó� ���кت�ǧ����)-----------------------------

                         ElseIf ChkAllEqp.Checked = True And ChkTime.Checked = False Then

                                 strSqlCmdSelc = "SELECT * FROM v_rpt_delvr (NOLOCK)" _
                                               & " WHERE  (doc_date >= '" & StartDate() _
                                               & "' AND doc_date <= '" & EndDate() & "')"

                    '-------------------------�����������(�ó��к������ػ�ó� �����͡�ء�ѹ���)-----------------------------

                         ElseIf ChkAllEqp.Checked = False And ChkTime.Checked = True Then

                                 strSqlCmdSelc = "SELECT * FROM v_rpt_delvr (NOLOCK)" _
                                               & " WHERE eqp_id = '" & cboEqpid.Text & "' "

                         Else

                    '-------------------------�����������(�ó��к������ػ�ó� ����кت�ǧ����)-----------------------------

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

                '-------------------------�觤�ҵ�������� Crystal Report (�ó������͡�ء�ػ�ó� ��� �����͡�ء�ѹ���)------------------------  

                If ChkAllEqp.Checked = True And ChkTime.Checked = True Then

                    Dim cryeqp As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryeqp = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryeqp")
                    cryeqp.Text = "�ء�����ػ�ó�"

                    Dim cryDf As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDf = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDf")
                    cryDf.Text = "-"

                    Dim cryDe As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDe = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDe")
                    cryDe.Text = "-"

                '-------------------------�觤�ҵ�������� Crystal Report (�ó������͡�ء�����ػ�ó�  ��������͡�ء�ѹ���)----------------------

                ElseIf ChkAllEqp.Checked = True And ChkTime.Checked = False Then

                    Dim cryeqp As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryeqp = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryeqp")
                    cryeqp.Text = "�ء�����ػ�ó�"

                    Dim cryDf As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDf = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDf")
                    SDThai = StartDate()
                    cryDf.Text = DateThai(SDThai)

                    Dim cryDe As CrystalDecisions.CrystalReports.Engine.TextObject
                    cryDe = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryDe")
                    EDThai = DateThai(EndDate)
                    cryDe.Text = EDThai

                '-------------------------�觤�ҵ�������� Crystal Report (�ó�������͡�ء�ػ�ó� �������͡�ء�ѹ���)------------------------  

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

                '--------------------- �觤�ҵ�������� Crystal Report (�ó���������͡�ء�ػ�ó� ��� ��������͡�ء�ѹ���) -------------------

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

               '----------------------- ����èҡ Crystal Report �Ѻ��Ҩҡ����������---- ------------------------

                     Dim cryuser1 As CrystalDecisions.CrystalReports.Engine.TextObject
                     cryuser1 = cryRpt.ReportDefinition.Sections(0).ReportObjects("cryuser1")
                     cryuser1.Text = strUser

                     cryRpt.PrintOptions.PaperSize = PaperSize.PaperA4
                     cryRpt.PrintOptions.PaperOrientation = PaperOrientation.Landscape

                     Viewer1.ReportSource = cryRpt
                     Viewer1.DisplayStatusBar = True
                     Viewer1.Refresh()
                     Viewer1.Zoom(80)

                     LockOptions(False)    '��� GroupboxMain ��� GroupboxSub  ��ҹ����� ��лԴ���� OK,Cancle

            Else
                
                MsgBox("��辺�����ŷ��س��ͧ��ä���!!" & vbNewLine _
                       & "��س��кآ���������!!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Data Empty!!")

                LockOptions(True)
               
                Exit Sub

                End If
                .ActiveConnection = Nothing
       End With

        btnOK.Enabled = False
        btnCancle.Enabled = True
        btnExit.Enabled = True

    Rsd1 = Nothing

    ds.Clear()    '���������� Dataset
    ds.Dispose()  '����µ����� DataSet

    da.Dispose() '����µ����� DataAdapter

    ds = Nothing  '������������ Dataset
    da = Nothing  '������������� adapter

  Conn.Close()
  Conn = Nothing

End Sub

    Function DateThai(ByVal str As String)  '�ѧ������ŧ�ѹ���ҡ ("2013-10-XX") �� dd/MM/yyyy

        Dim yyyy, MM, dd As String
        Dim d As String

        '�Ѻ����(yyyy - MM - dd) �շ����� 14 ����˹�

        yyyy = Mid(str, 1, 4)  '�Ѵ��
        dd = Mid(str, 9, 2)  '�Ѵ�ѹ��
        MM = Mid(str, 6, 2)  '�Ѵ��͹

        d = dd & "/" & MM & "/" & yyyy            '�ŧ�ѹ����� dd/MM/yyyy
        Return d
    End Function


    '--------------------------- Event ����� CheckBox ���͡�ء�����ػ�ó�  ---------------------------------------

Private Sub chkAllEqp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkAllEqp.CheckedChanged
  If ChkAllEqp.Checked Then
     cboEqpid.Enabled = False
  Else
     cboEqpid.Enabled = True
  End If
  End Sub

    '--------------------------- Event �����CheckBox ���͡���������ҷ�����  ---------------------------------------

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
    Viewer1.ReportSource = Nothing   '������ Report �����Ŵ�����
    LockOptions(True)   '�óա����� btnCancle ������ btnOK, GroupboxMain ��� GroupboxSub  ��ҹ��
 End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
   Me.Close()
End Sub

Private Sub LockOptions(ByVal blnSta As Boolean)   '�Ѻ�ٷչ��ͤ������ ���Ѻ��� boolean ��
    gpbMain.Enabled = blnSta
    gpbSub.Enabled = blnSta
    If blnSta Then   '����繨�ԧ
        btnOK.Enabled = blnSta       '���� btnOK ����ö��ҹ��
        btnCancle.Enabled = False    '���� btnCancle �١��ͤ
    Else                                  ' �ó�����
        btnOK.Enabled = blnSta       '������ btnOK �ء��ͤ
        btnCancle.Enabled = blnSta   '������ btnCancle ��ͤ
    End If
End Sub

End Class