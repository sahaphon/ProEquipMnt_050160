Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class frmWipImp

Dim ConnSQL As New SqlConnection
Dim ConnDbase As New OleDbConnection
Dim blnInterrupt As Boolean

Private Sub frmPayImpt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  StdDateTime()
  blnInterrupt = False
  pgbSta.Value = 0
  lblCurentRd.Text = "0"
  StateUnlock()
End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
  Me.Close()
End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
  blnInterrupt = True
End Sub

Sub StateUnlock()
  btnImport.Enabled = True
  btnCancel.Enabled = False
  btnExit.Enabled = True
End Sub

Sub StateLock()
  btnImport.Enabled = False
  btnCancel.Enabled = True
  btnExit.Enabled = False
End Sub

Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click

 Dim bytConSave As Byte

     bytConSave = MsgBox("คุณต้องการนำเข้าข้อมูล" & vbNewLine & "ใช่หรือไม่!" _
                            , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Exclamation, "IMPORT DATA!!!")

     If bytConSave = 6 Then
        Me.Cursor = Cursors.WaitCursor
        StateLock()
        ImportWipData()
     Else
            MsnAdmin()
     End If

End Sub

Sub ImportWipData()

 Dim adap As New OleDbDataAdapter
 Dim cmd As OleDbCommand = Nothing
 Dim sqlCmd As SqlCommand = Nothing
 Dim Trans As SqlTransaction = Nothing
 Dim ds As New DataSet
 Dim dt As DataTable
 Dim SQL As String

 Dim i As Integer = 0

     'Connect dbase
     ConnDbase = New OleDbConnection(dbase)
     If ConnDbase.State = 0 Then ConnDbase.Open()

        SQL = "SELECT * FROM pilotrl"
        cmd = New OleDbCommand(SQL, ConnDbase)
        adap.SelectCommand = cmd
        adap.Fill(ds, "cs")
        dt = ds.Tables("cs")
        ConnDbase.Close()

        If dt.Rows.Count > 0 Then

           With ConnSQL

                If .State Then .Close()
                   .ConnectionString = sqlclint
                   .Open()
                   Trans = ConnSQL.BeginTransaction
           End With

          'Delete prodd1
          SQL = "DELETE FROM pilotrl"
          sqlCmd = New SqlCommand(SQL, ConnSQL, Trans)
          sqlCmd.ExecuteNonQuery()
          sqlCmd.Dispose()

          pgbSta.Minimum = 0
          pgbSta.Maximum = dt.Rows.Count
          lblCurentRd.Text = "0"
          'MsgBox("จำนวน : " & dt.Rows.Count)

          Try

               For Each r As Object In dt.Rows

                   SQL = "INSERT INTO pilotrl(o_code, o_date, en_date, d_ship, p_code" _
                             & ", run_id, style, colorcode, size, qty, pre_date)" _
                             & "VALUES(" _
                             & "'" & r("o_code") & "'" _
                             & ",'" & IIf(CDate(r("o_date")) > CDate("1899-12-30"), CDate(r("o_date")).ToString("yyyy-MM-dd"), "1899-12-30") & "'" _
                             & ",'" & IIf(CDate(r("en_date")) > CDate("1899-12-30"), CDate(r("en_date")).ToString("yyyy-MM-dd"), "1899-12-30") & "'" _
                             & ",'" & IIf(CDate(r("d_ship")) > CDate("1899-12-30"), CDate(r("d_ship")).ToString("yyyy-MM-dd"), "1899-12-30") & "'" _
                             & ",'" & r("p_code") & "'" _
                             & ",'" & r("run_id") & "'" _
                             & ",'" & r("style") & "'" _
                             & ",'" & r("colorcode") & "'" _
                             & ",'" & r("size") & "'" _
                             & "," & r("qty") _
                             & ",'" & IIf(CDate(r("pre_date")) > CDate("1899-12-30"), CDate(r("pre_date")).ToString("yyyy-MM-dd"), "1899-12-30") & "'" _
                             & ")"

                    sqlCmd = New SqlCommand(SQL, ConnSQL, Trans)
                    sqlCmd.ExecuteNonQuery()

                    i = i + 1
                    pgbSta.Value = i
                    lblCurentRd.Text = i

                    Application.DoEvents()
               Next

               If blnInterrupt = True Then
                  Me.Cursor = Cursors.Arrow
                  pgbSta.Value = 0
                  lblCurentRd.Text = "0"
                  Trans.Rollback()
                  StateUnlock()
               Else
                     Me.Cursor = Cursors.Arrow
                     Trans.Commit()
                     StateUnlock()
                     MsgBox("นำเข้าข้อมูล Wipnew ประสบผลสำเร็จ", MsgBoxStyle.Information, "Success!!")
               End If

          Catch ex As Exception
                Me.Cursor = Cursors.Arrow
                pgbSta.Value = 0
                lblCurentRd.Text = "0"
                Trans.Rollback()
                MsgBox("ERROR : " & SQL & ex.Message)
                StateUnlock()
                Exit Sub
          End Try
     End If

    ds.Clear()
    sqlCmd.Dispose()
    ConnSQL.Close()
End Sub

End Class