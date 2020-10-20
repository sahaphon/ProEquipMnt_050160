Public Class frmAeFixEqp

Private Sub txtRep_ID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRep_ID.KeyDown
 Dim intChkPoint As Integer
 With txtRep_ID
     Select Case e.KeyCode
            Case Is = 35 'ปุ่ม End 
            Case Is = 36 'ปุ่ม Home
            Case Is = 37 'ลูกศรซ้าย
            Case Is = 38 'ปุ่มลูกศรขึ้น
            Case Is = 39   'ปุ่มลูกศรขวา
                 If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                    cmbType.DroppedDown = True
                    cmbType.Focus()
                 Else
                     intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            cmbType.DroppedDown = True
                            cmbType.Focus()
                        End If
                 End If

            Case Is = 40 'ปุ่มลง
                      txtEqp_id.Focus()
            Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
     End Select
  End With
End Sub

Private Sub txtRep_ID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRep_ID.KeyPress
  If e.KeyChar = Chr(13) Then
     cmbType.DroppedDown = True
     cmbType.Focus()
  End If
End Sub

Private Sub txtEqp_id_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqp_id.KeyDown
 Dim intChkPoint As Integer
 With txtEqp_id
     Select Case e.KeyCode
            Case Is = 35 'ปุ่ม End 
            Case Is = 36 'ปุ่ม Home
            Case Is = 37 'ลูกศรซ้าย
                      If .SelectionStart = 0 Then
                          cmbType.DroppedDown = True
                          cmbType.Focus()
                      End If
            Case Is = 38 'ปุ่มลูกศรขึ้น
                      txtRep_ID.Focus()
            Case Is = 39   'ปุ่มลูกศรขวา
                 If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                    txtEqpnm.Focus()
                 Else
                     intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtEqpnm.Focus()
                        End If
                 End If

            Case Is = 40 'ปุ่มลง
                      txtRemark.Focus()
            Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
     End Select
  End With
End Sub

Private Sub txtEqp_id_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqp_id.KeyPress
  If e.KeyChar = Chr(13) Then
     txtEqpnm.Focus()
  End If
End Sub

Private Sub txtEqpnm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqpnm.KeyDown
 Dim intChkPoint As Integer
 With txtEqpnm
     Select Case e.KeyCode
            Case Is = 35 'ปุ่ม End 
            Case Is = 36 'ปุ่ม Home
            Case Is = 37 'ลูกศรซ้าย
                      If .SelectionStart = 0 Then
                         txtEqp_id.Focus()
                      End If
            Case Is = 38 'ปุ่มลูกศรขึ้น
                      cmbType.DroppedDown = True
                      cmbType.Focus()
            Case Is = 39   'ปุ่มลูกศรขวา
                 If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate
                    txtRemark.Focus()
                 Else
                     intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate
                            txtRemark.Focus()
                        End If
                 End If

            Case Is = 40 'ปุ่มลง
                      txtRemark.Focus()
            Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
     End Select
  End With
End Sub

Private Sub txtEqpnm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqpnm.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRemark.Focus()
  End If
End Sub

Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
 Dim intChkPoint As Integer
 With txtRemark
     Select Case e.KeyCode
            Case Is = 35 'ปุ่ม End 
            Case Is = 36 'ปุ่ม Home
            Case Is = 37 'ลูกศรซ้าย
                      If .SelectionStart = 0 Then
                         txtEqpnm.Focus()
                      End If
            Case Is = 38 'ปุ่มลูกศรขึ้น
                      txtEqpnm.Focus()
            Case Is = 39   'ปุ่มลูกศรขวา
                 If .SelectionLength = .Text.Trim.Length Then  'ถ้าความยาวตำแหน่งปัจจุบัน = ความยาวของ mskLdate

                 Else
                     intChkPoint = .Text.Trim.Length     'ให้ InChkPoint = ความยาวของ  mskLdate
                        If .SelectionStart = intChkPoint Then    'ถ้า Pointer ชี้ไปที่ตำแหน่งสุดท้ายของ mskLdate

                        End If
                 End If

            Case Is = 40 'ปุ่มลง

            Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
     End Select
  End With
End Sub

Private Sub txtRemark_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRemark.KeyPress

End Sub

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click

End Sub
End Class