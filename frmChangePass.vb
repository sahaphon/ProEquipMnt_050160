Imports ADODB

Public Class frmChangePass

Private Sub frmChangePass_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
    txtPass.Focus()
End Sub

Private Sub frmChangePass_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    Me.Dispose()
End Sub

Private Sub frmChangePass_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  lblUsrLogIn.Text = frmMainpro.lblLogin.Text
End Sub

Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
  Me.Close()
End Sub

Private Sub ChangePassWord()

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String

Dim btyConsider As Byte
Dim strPassWord As String = txtPass.Text.ToUpper.Trim

Dim blnDataComplete As Boolean

        If Len(strPassWord) <> 0 Then


                btyConsider = MsgBox("ผู้ใช้งาน : " & lblUsrLogIn.Text & vbNewLine _
                                                & "คุณต้องการเปลี่ยนรหัสผ่าน" & vbNewLine _
                                                & "ใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                + MsgBoxStyle.Exclamation, "Confirm Change Password")

                If btyConsider = 6 Then

                        With Conn

                                 If .State Then .Close()
                                    .ConnectionString = strConnAdodb
                                    .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                    .ConnectionTimeout = 90
                                    .Open()

                                    .BeginTrans()
                                     strSqlCmd = "UPDATE usermst SET pass ='" & strPassWord & "'" _
                                                    & " WHERE user_id ='" & lblUsrLogIn.Text & "'"
                                    .Execute(strSqlCmd)
                                    .CommitTrans()
                                    .Close()

                         End With
                         Conn = Nothing

                         blnDataComplete = True

                Else
                    txtPass.Focus()
                End If

        Else

            MsgBox("โปรดระบุรหัสผ่านใหม่ ก่อนบันทึกข้อมูล!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "New Password Is Null!!")
            txtPass.Focus()

        End If

If blnDataComplete Then
    Me.Close()
End If

End Sub

Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
  ChangePassWord()
End Sub

Private Sub txtPass_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPass.KeyPress

If e.KeyChar = Chr(13) Then
    ChangePassWord()
End If

End Sub

End Class