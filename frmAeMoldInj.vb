Imports ADODB
Imports System.IO
Imports System.Drawing

Public Class frmAeMoldInj

Dim IsShowSeek As Boolean
Dim strDateDefault As String

Public Const DrvName As String = "\\10.32.0.15\data1\EquipPicture\"
Public Const PthName As String = "\\10.32.0.15\data1\EquipPicture"
Private tt As ToolTip = New ToolTip 'แสดงทุูลทิป ในรูปภาพเวลาเลื่อนเคอร์เซอร์

Protected Overrides ReadOnly Property CreateParams() As CreateParams 'ป้องกันการปิดโดยใช้ปุ่ม Close Button

 Get
      Dim cp As CreateParams = MyBase.CreateParams
      Const CS_DBLCLKS As Int32 = &H8
      Const CS_NOCLOSE As Int32 = &H200
            cp.ClassStyle = CS_DBLCLKS Or CS_NOCLOSE
      Return cp
 End Get

End Property

Private Sub frmAeMoldInj_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearTmpTable(0, "")
  ClearTmpTable(2, "") 'ล้างข้อมูลตาราง tmp_fixeqptrn
  ClearTmpTable(3, "")  'ล้างข้อมูลตาราง tmp_eqptrn_newsize
  frmMoldInj.lblCmd.Text = "0" 'เคลียร์สถานะ
  Me.Dispose()
End Sub

Private Sub ClearTmpTable(ByVal bytOption As Byte, ByVal strPsID As String)

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String = ""

    With Conn

         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()

            Select Case bytOption

                   Case Is = 0 'ลบข้อมูลหลังจากปิดฟอร์ม

                        strSqlCmd = "Delete FROM tmp_eqptrn" _
                                           & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"
                        .Execute(strSqlCmd)

                   Case Is = 1

                       strSqlCmd = "Delete FROM tmp_eqptrn" _
                                         & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                         & " AND docno ='" & strPsID.ToString.Trim & "'"
                      .Execute(strSqlCmd)

                   Case Is = 2

                       strSqlCmd = "Delete FROM tmp_fixeqptrn" _
                                         & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"
                      .Execute(strSqlCmd)

                  Case Is = 3

                       strSqlCmd = "Delete FROM tmp_eqptrn_newsize" _
                                         & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"
                      .Execute(strSqlCmd)

           End Select

       End With

   Conn.Close()
   Conn = Nothing

End Sub

Private Sub frmAeMoldInj_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

  Dim dteComputer As Date = Now()
  Dim strCurrentDate As String

      StdDateTimeThai()
      strCurrentDate = dteComputer.Date.ToString("dd/MM/yyyy")

      PreGpSeek()
      PreTypeSeek()
      PrePartSeek()
      PreMoldStatus() 'สถานะโมล์ด

       Select Case frmMoldInj.lblCmd.Text.ToString

              Case Is = "0" 'เพิ่มข้อมูล

                   ClearAllData()

                   With txtBegin
                        .Text = strCurrentDate
                         strDateDefault = strCurrentDate
                   End With

                   With Me
                        .Text = "เพิ่มข้อมูล"
                        txtExpId.Focus()
                   End With

              Case Is = "1" 'แก้ไขข้อมูล

                   With Me
                        .Text = "เเก้ไขข้อมูล"
                   End With

                     LockEditData()
                     LoadHistory_Fixmold()            'โหลดประวัติการซ่อมโมล์ด
                     txtEqpId.ReadOnly = True
                     cmbGp.Enabled = False

              Case Is = "2" 'มุมมองข้อมูล

                   With Me
                        .Text = "มุมมองข้อมูล"
                   End With

                     LockEditData()
                     LoadHistory_Fixmold()            'โหลดประวัติการซ่อมโมล์ด
                     txtEqpId.ReadOnly = True
                     cmbGp.Enabled = False
                     btnSaveData.Enabled = False

      End Select

End Sub

Private Sub LockEditData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlSelcSelc As String
Dim strCmd As String

Dim strLoadFilePicture As String
Dim strPathPicture As String = "\\10.32.0.15\data1\EquipPicture\"

Dim blnHaveData As Boolean
Dim dteComputer As Date = Now()

Dim strSqlSelc As String = ""
Dim strPart As String = ""

Dim strCode As String = frmMoldInj.dgvShoe.Rows(frmMoldInj.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
   End With

           strSqlSelcSelc = "SELECT * FROM v_moldinj_hd (NOLOCK)" _
                                 & " WHERE eqp_id ='" & strCode & "'"

   With Rsd

        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        .LockType = ADODB.LockTypeEnum.adLockOptimistic
        .Open(strSqlSelcSelc, Conn, , , )

        If .RecordCount <> 0 Then
           cmbGp.Text = .Fields("desc_eng").Value.ToString.Trim
           txtBegin.Text = Mid(.Fields("creat_date").Value.ToString.Trim, 1, 10)
           strDateDefault = Mid(.Fields("creat_date").Value.ToString.Trim, 1, 10)

           txtEqpId.Text = .Fields("eqp_id").Value.ToString.Trim
           txtEqpName.Text = .Fields("eqp_name").Value.ToString.Trim
           txtStyle.Text = .Fields("shoe").Value.ToString.Trim
           txtOrd.Text = .Fields("pi").Value.ToString.Trim

           txtRef.Text = .Fields("doc_ref").Value.ToString.Trim
           txtSet.Text = Format(.Fields("set_qty").Value, "#,##0.0")
           txtRemark.Text = .Fields("remark").Value.ToString.Trim

           lblPicName1.Text = .Fields("pic_ctain").Value.ToString.Trim
           lblPicName2.Text = .Fields("pic_io").Value.ToString.Trim
           lblPicName3.Text = .Fields("pic_part").Value.ToString.Trim

           lblPicPath1.Text = PthName
           lblPicPath2.Text = PthName
           lblPicPath3.Text = PthName

           '-----------------------------  ใส่ชิ้นงาน  ----------------------------

           Select Case .Fields("part").Value.ToString.Trim

                  Case Is = "UPPER"
                       cmbPart.Text = "หนังหน้า"

                  Case Is = "ACCSSY"
                       cmbPart.Text = "อุปกรณ์"

                  Case Is = "SOLE1"
                       cmbPart.Text = "พื้นบน"

                  Case Is = "SOLE2"
                       cmbPart.Text = "พื้นล่าง"

                  Case Is = "SOLE3"
                       cmbPart.Text = "รองเท้าสำเร็จรูป"

                  Case Is = "GUARD"
                       cmbPart.Text = "การ์ด"

                  Case Is = "BUTTON"
                       cmbPart.Text = "กระดุม"

                  Case Is = "NOSE"
                       cmbPart.Text = "จมูก"

                  Case Is = "LEG"
                       cmbPart.Text = "ขา"

                  Case Is = "TPR"
                       cmbPart.Text = "TPR โลโก้"

          End Select

                '----------------------------- Load รูปภาพบรรจุอุปกรณ์  ----------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_ctain").Value.ToString.Trim
                If File.Exists(strLoadFilePicture) Then
                   Dim img1 As Image
                   img1 = Image.FromFile(strLoadFilePicture)
                   picEqp1.Image = ScaleImage(img1, picEqp1.Height, picEqp2.Width)
                Else
                    picEqp1.Image = Nothing
                End If
                strLoadFilePicture = ""

                '--------------------------- Load รูปภาพภายนอก ภายใน ----------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_io").Value.ToString.Trim
                If File.Exists(strLoadFilePicture) Then
                   Dim img2 As Image
                       img2 = Image.FromFile(strLoadFilePicture)
                       picEqp2.Image = ScaleImage(img2, picEqp2.Height, picEqp2.Width)
                Else
                     picEqp2.Image = Nothing
                End If
                strLoadFilePicture = ""

               '----------------------------- Load รูปภาพชิ้นงาน --------------------------

                strLoadFilePicture = strPathPicture & .Fields("pic_part").Value.ToString.Trim
                If File.Exists(strLoadFilePicture) Then  'Exists ใช้ตรวจสอบไฟล์ซ้ำ
                   Dim img3 As Image
                   img3 = Image.FromFile(strLoadFilePicture)
                   picEqp3.Image = ScaleImage(img3, picEqp3.Height, picEqp3.Width)
                Else
                   picEqp3.Image = Nothing
                End If
                strLoadFilePicture = ""

                        strCmd = frmMoldInj.lblCmd.Text.ToString
                        Select Case strCmd
                               Case Is = "1" 'ให้ล็อคตอนแก้ไข
                               Case Is = "2" 'ให้ล็อคตอนมุมมอง
                                    btnSaveData.Enabled = False
                        End Select

                        '--------------------- โหลดข้อมูลใส่ตาราง tmp_fixeqptrn ---------------

                         strSqlSelc = "INSERT INTO tmp_fixeqptrn " _
                                     & "SELECT user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                     & " FROM fixeqptrn " _
                                     & " WHERE eqp_id = '" & strCode & "'" _
                                     & " AND fix_sta= '" & "1" & "'"

                         Conn.Execute(strSqlSelc)

                        '--------------------- บันทึกข้อมูลในตาราง tmp_eqptrn -----------------

                        strSqlSelc = "INSERT INTO tmp_eqptrn" _
                                   & " SELECT user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "',*" _
                                   & " FROM eqptrn" _
                                   & " WHERE eqp_id ='" & strCode & "'"

                        Conn.Execute(strSqlSelc)

                        '-------------------- เรียง SIZE ใหม่-------------------------------

                        ReSizeSort(strCode)   'จัดเรียง size เสียใหม่

                        blnHaveData = True

                        Else
                            blnHaveData = False
                        End If

                    .ActiveConnection = Nothing
                    .Close()

             End With
             Rsd = Nothing

   Conn.Close()
   Conn = Nothing

If blnHaveData Then
   ShowScrapItem()
End If
    End Sub

Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
   Me.Close()
End Sub

Private Sub PreGpSeek()

Dim strGpTopic(2) As String
Dim i As Byte

      strGpTopic(0) = "MOLD EVA"
      strGpTopic(1) = "MOLD PVC"
      strGpTopic(2) = "MOLD PU"

      With cmbGp

              For i = 0 To 2
                 .Items.Add(strGpTopic(i))
              Next i

      End With

End Sub

Private Sub cmbGp_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbGp.KeyPress
  If e.KeyChar = Chr(13) Then
     txtEqpId.Focus()
  End If
End Sub

Private Sub cmbGp_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGp.TextChanged

   Select Case cmbGp.SelectedIndex

          Case Is = 0
                  lblGpName.Text = "โมล์ดฉีด EVA INJECTION"
          Case Is = 1
                  lblGpName.Text = "โมล์ดฉีด PVC INJECTION"
          Case Is = 2
                  lblGpName.Text = "โมล์ดหยอด PU"
          Case Else
                  lblGpName.Text = ""

   End Select

End Sub

Private Sub PreTypeSeek()

Dim strGpTopic(1) As String
Dim i As Byte

    strGpTopic(0) = "โมลด์ต่างประเทศ"
    strGpTopic(1) = "โมลด์ในประเทศ"

    With cmbType

         For i = 0 To 1
             .Items.Add(strGpTopic(i))
         Next i

    End With

End Sub

Private Sub PrePartSeek()

Dim strGpTopic(10) As String
Dim i As Byte

      strGpTopic(0) = "หนังหน้า"
      strGpTopic(1) = "อุปกรณ์"
      strGpTopic(2) = "พื้นบน"
      strGpTopic(3) = "พื้นล่าง"
      strGpTopic(4) = "รองเท้าสำเร็จรูป"
      strGpTopic(5) = "การ์ด"
      strGpTopic(6) = "กระดุม"
      strGpTopic(7) = "จมูก"
      strGpTopic(8) = "ขา"
      strGpTopic(9) = "TPR โลโก้"

      With cmbPart
           For i = 0 To 9
               .Items.Add(strGpTopic(i))
           Next i
      End With
End Sub

Private Sub txtBegin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBegin.GotFocus
   With mskBegin
        txtBegin.SendToBack()
        .BringToFront()
        .Focus()
   End With
End Sub

Private Sub mskBegin_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskBegin.GotFocus

Dim i, x As Integer

Dim strTmp As String = ""
Dim strMerge As String = ""

    With mskBegin

         If txtBegin.Text <> "__/__/____" Then
            x = Len(txtBegin.Text.ToString)

            For i = 1 To x
                strTmp = Mid(txtBegin.Text.ToString, i, 1)

                Select Case strTmp
                       Case Is = "_"
                       Case Else

                           If InStr("0123456789/", strTmp) > 0 Then
                              strMerge = strMerge & strTmp
                           End If

                 End Select

                         Next i

                        Select Case strMerge.ToString.Length

                               Case Is = 10
                                    .SelectionStart = 0

                        End Select

                   .SelectedText = strMerge
                End If

        .SelectAll()

End With

End Sub

Private Sub mskBegin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskBegin.KeyDown

Dim intChkPoint As Integer

    With mskBegin

         Select Case e.KeyCode
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                     If .SelectionStart = 0 Then
                     End If

                Case Is = 38 'ปุ่มลูกศรขึ้น
                     cmbGp.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา

                     If .SelectionLength = .Text.Trim.Length Then
                     Else
                         intChkPoint = .Text.Trim.Length
                         If .SelectionStart = intChkPoint Then
                             End If
                         End If

               Case Is = 40 'ปุ่มลง
                    txtEqpId.Focus()
               Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub mskBegin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskBegin.KeyPress

   Select Case e.KeyChar

          Case Is = Chr(13)
               txtEqpId.Focus()

          Case Is = Chr(46) 'เครื่องหมายจุลภาค(.)
                  'mskBegin.SelectionStart = 6
  End Select

End Sub

Private Sub mskBegin_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskBegin.LostFocus

Dim i, x As Integer
Dim z As Date

Dim strTmp As String = ""
Dim strMerge As String = ""

    With mskBegin

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

       mskBegin.Text = ""
       strMerge = "#" & strMerge & "#"
       z = CDate(strMerge)

       If Year(z) < 2500 Then 'ปีคริสต์ < 2100                        
          txtBegin.Text = Mid(z.ToString("dd/MM/yyyy"), 1, 6) & Trim(Str(Year(z) + 543))
       Else
           txtBegin.Text = z.ToString("dd/MM/yyyy")
       End If

   Catch ex As Exception
         txtBegin.Text = strDateDefault
         mskBegin.Text = ""
   End Try

  mskBegin.SendToBack()
  txtBegin.BringToFront()

End With

End Sub

Private Sub txtEqpId_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpId.GotFocus
  txtEqpId.SelectAll()
End Sub

Private Sub txtEqpId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqpId.KeyDown

Dim intChkPoint As Integer

    With txtEqpId

         Select Case e.KeyCode

                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย

                     If .SelectionStart = 0 Then
                     End If

                Case Is = 38 'ปุ่มลูกศรขึ้น
                     cmbGp.Focus()
                Case Is = 39 'ปุ่มลูกศรขวา

                     If .SelectionLength = .Text.Trim.Length Then
                        txtEqpName.Focus()
                     Else

                        intChkPoint = .Text.Trim.Length
                        If .SelectionStart = intChkPoint Then
                           txtEqpName.Focus()
                           End If

                       End If

               Case Is = 40 'ปุ่มลง
                    txtStyle.Focus()
               Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtEqpId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqpId.KeyPress

          Select Case e.KeyChar
                 Case "0" To "9"
                            e.Handled = False

                 Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" _
                            , "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
                            e.Handled = False

                 Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n" _
                            , "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
                            e.Handled = False

                 Case "-"
                            e.Handled = False

                 Case Is = Chr(13)
                            e.Handled = False
                            txtEqpName.Focus()

                 Case Chr(8), Chr(46)
                            e.Handled = False

                 Case Else
                            e.Handled = True

        End Select

End Sub

Private Sub txtEqpId_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpId.LostFocus
    With txtEqpId
            .Text = .Text.ToString.Trim.ToUpper
    End With
End Sub

Private Sub txtEqpIdNm_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpName.GotFocus
    txtEqpName.SelectAll()
End Sub

Private Sub txtEqpIdNm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEqpName.KeyDown

Dim intChkPoint As Integer

    With txtEqpName

               Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                        txtEqpId.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                   
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtStyle.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtStyle.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง
                                 txtOrd.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtEqpIdNm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEqpName.KeyPress
If e.KeyChar = Chr(13) Then
        txtStyle.Focus()
End If
End Sub

Private Sub txtEqpIdNm_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEqpName.LostFocus
  With txtEqpName
       .Text = .Text.ToString.Trim.ToUpper
  End With
End Sub

Private Sub txtStyle_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStyle.GotFocus
 txtStyle.SelectAll()
End Sub

Private Sub txtStyle_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStyle.KeyDown
 Dim intChkPoint As Integer

     With txtStyle

               Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                        txtEqpName.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น 
                                  txtEqpId.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtOrd.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                             txtOrd.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง
                                 txtRemark.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtStyle_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStyle.KeyPress

    Select Case e.KeyChar
           Case "0" To "9"
                 e.Handled = False
           Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" _
                            , "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
                 e.Handled = False

           Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n" _
                     , "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
                 e.Handled = False
           Case "-"
                 e.Handled = False
           Case Is = Chr(13)
                     e.Handled = False
                     txtOrd.Focus()
           Case Chr(8), Chr(46)
                e.Handled = False
           Case Else
                e.Handled = True
  End Select
End Sub

Private Sub txtStyle_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStyle.LostFocus
   With txtStyle
        .Text = .Text.ToString.Trim.ToUpper
   End With
End Sub

Private Sub txtOrd_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrd.GotFocus
    txtOrd.SelectAll()
End Sub

Private Sub txtOrd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOrd.KeyDown

Dim intChkPoint As Integer

    With txtOrd

            Select Case e.KeyCode
                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย
                        If .SelectionStart = 0 Then
                           txtStyle.Focus()
                        End If
                   Case Is = 38 'ปุ่มลูกศรขึ้น 
                        txtEqpName.Focus()
                   Case Is = 39 'ปุ่มลูกศรขวา

                        If .SelectionLength = .Text.Trim.Length Then
                           cmbPart.Focus()
                        Else
                             intChkPoint = .Text.Trim.Length
                             If .SelectionStart = intChkPoint Then
                                cmbPart.Focus()
                             End If
                        End If
                  Case Is = 40 'ปุ่มลง
                       txtRemark.Focus()
                  Case Is = 113 'ปุ่ม F2
                       .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub txtOrd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOrd.KeyPress

    Select Case e.KeyChar
           Case "0" To "9"
                e.Handled = False
           Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" _
                   , "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
                e.Handled = False
           Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n" _
                            , "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
                e.Handled = False
           Case "-"
                e.Handled = False
           Case Is = Chr(13)
                e.Handled = False
                cmbPart.DroppedDown = True
           Case Chr(8), Chr(46)
                e.Handled = False
           Case Else
                e.Handled = True
   End Select
End Sub

Private Sub txtOrd_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOrd.LostFocus
  With txtOrd
       .Text = .Text.ToString.Trim.ToUpper
  End With
End Sub

Private Sub cmbPart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbPart.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRemark.Focus()
  End If
End Sub

Private Sub cmbType_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbType.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRecvDate.Focus()
  End If
End Sub

Private Sub txtSuplier_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuplier.GotFocus
  txtSuplier.SelectAll()
End Sub

Private Sub txtSuplier_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSuplier.KeyDown

Dim intChkPoint As Integer

    With txtSuplier

            Select Case e.KeyCode
                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย
                        If .SelectionStart = 0 Then
                           txtImpt.Focus()
                        End If
                  Case Is = 38 'ปุ่มลูกศรขึ้น 
                       txtPrice.Focus()
                  Case Is = 39 'ปุ่มลูกศรขวา
                       If .SelectionLength = .Text.Trim.Length Then
                          cmbType.Focus()
                       Else
                            intChkPoint = .Text.Trim.Length
                            If .SelectionStart = intChkPoint Then
                               cmbType.Focus()
                            End If
                       End If
                Case Is = 40 'ปุ่มลง
                     txtInvoice.Focus()
                Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub txtSuplier_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSuplier.KeyPress
 If e.KeyChar = Chr(13) Then
    cmbType.Focus()
 End If
End Sub

Private Sub txtSuplier_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSuplier.LostFocus
   With txtSuplier
        .Text = .Text.ToString.Trim.ToUpper
   End With
End Sub

Private Sub txtRef_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRef.GotFocus
  txtRef.SelectAll()
End Sub

Private Sub txtRef_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRef.KeyDown
Dim intChkPoint As Integer

    With txtSuplier

            Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                        txtRef.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น 
                                  cmbPart.Focus()

                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtSet.Focus()

                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtSet.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง
                                 txtSet.Focus()

                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length

            End Select

    End With
End Sub

Private Sub txtRef_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRef.KeyPress
   If e.KeyChar = Chr(13) Then
      txtSet.Focus()
   End If
End Sub

Private Sub txtRef_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRef.LostFocus
  With txtRef
       .Text = .Text.ToString.Trim.ToUpper
  End With
End Sub

Private Sub txtSet_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSet.GotFocus
  txtSet.SelectAll()
End Sub

Private Sub mskQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskQty.GotFocus
   Dim i, x As Integer
   Dim strTmp As String = ""
   Dim strMerge As String = ""

       With mskQty

           If txtSet.Text.ToString.Trim <> "" Then

                        x = Len(txtSet.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtSet.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = "_"
                                       Case Else

                                            If InStr("0123456789.", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If
                                End Select

                         Next i

                         Select Case strMerge.IndexOf(".")

                                Case Is = -1
                                     .SelectionStart = 0
                                Case Is = 1
                                     .SelectionStart = 1
                                Case Is = 2
                                     .SelectionStart = 0
                                Case Is = 3
                                     .SelectionStart = 0
                                Case Else
                                     .SelectionStart = 0
                        End Select

                        .SelectedText = strMerge

                End If
            .SelectAll()

     End With
End Sub

Private Sub mskQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskQty.KeyDown

Dim intChkPoint As Integer

    With mskQty

            Select Case e.KeyCode
                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                               End If

                   Case Is = 38 'ปุ่มลูกศรขึ้น 
                              cmbType.Focus()
                   Case Is = 39 'ปุ่มลูกศรขวา

                           If .SelectionLength = .Text.Trim.Length Then
                              txtRemark.Focus()
                           Else
                               intChkPoint = .Text.Trim.Length
                               If .SelectionStart = intChkPoint Then
                                  txtRemark.Focus()
                               End If
                           End If
                   Case Is = 40 'ปุ่มลง                                                            
                   Case Is = 113 'ปุ่ม F2
                             .SelectionStart = .Text.Trim.Length

            End Select
    End With
End Sub

Private Sub mskQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskQty.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRemark.Focus()
  End If
End Sub

Private Sub mskQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskQty.LostFocus
Dim i, x As Integer
Dim z As Double

Dim strTmp As String = ""
Dim strMerge As String = ""

    With mskQty

           x = .Text.Length

               For i = 1 To x

                     strTmp = Mid(.Text.ToString, i, 1)
                     Select Case strTmp
                            Case Is = ","
                            Case Is = "+"
                            Case Is = "_"
                            Case Else
                                 If InStr("0123456789.", strTmp) > 0 Then
                                    strMerge = strMerge & strTmp
                                 End If

                    End Select
                    strTmp = ""
              Next i
              Try

                    mskQty.Text = ""
                    z = CDbl(strMerge)
                    txtSet.Text = z.ToString("#,##0.0")

                Catch ex As Exception
                    txtSet.Text = "0.0"
                    mskQty.Text = ""
               End Try

    mskQty.SendToBack()
    txtSet.BringToFront()

End With
End Sub

Private Sub txtRemark_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.GotFocus
  txtRemark.SelectAll()
End Sub

Private Sub txtRemark_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRemark.KeyDown
  Dim intChkPoint As Integer

      With txtRemark

            Select Case e.KeyCode
                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย

                        If .SelectionStart = 0 Then
                           cmbPart.Focus()
                        End If

                   Case Is = 38 'ปุ่มลูกศรขึ้น 
                              txtOrd.Focus()
                   Case Is = 39 'ปุ่มลูกศรขวา

                           If .SelectionLength = .Text.Trim.Length Then
                              cmbGp.Focus()
                           Else
                                intChkPoint = .Text.Trim.Length
                                If .SelectionStart = intChkPoint Then
                                   cmbGp.Focus()
                                End If
                           End If

                   Case Is = 40 'ปุ่มลง     
                         cmbGp.Focus()
                   Case Is = 113 'ปุ่ม F2
                         .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtRemark_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRemark.KeyPress
   If e.KeyChar = Chr(13) Then
      cmbGp.Focus()
   End If
End Sub

Private Sub txtRemark_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRemark.LostFocus
    With txtRemark
         .Text = .Text.ToString.Trim.ToUpper
    End With
End Sub

Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
  ClearSubData()
  ShowResvrd()
  gpbSeek.Text = "เพิ่มข้อมูล"
  txtSize.ReadOnly = False
  txtSizeDesc.ReadOnly = False
  txtSize.Focus()
End Sub

Sub ClearSubData()
    txtSize.Text = ""
    txtSizeDesc.Text = ""
    txtSetQty.Text = "0.0"
    txtSizeQty.Text = "0.0"
    txtInvoice.Text = ""
End Sub

Private Sub ShowResvrd()

  tabMain.SelectedTab = tabSize

  IsShowSeek = Not IsShowSeek
  If IsShowSeek Then

     With gpbSeek
          .Visible = True
          .Left = 8
          .Top = 252
          .Height = 450
          .Width = 1006
    End With

    StateLockFindDept(False)

  Else
     StateLockFindDept(True)

  End If

End Sub


Private Sub StateLockFindDept(ByVal Sta As Boolean)

Dim strMode As String = frmMoldInj.lblCmd.Text.ToString
    btnAdd.Enabled = Sta
    gpbHead.Enabled = Sta
    tabMain.Enabled = Sta
    btnSaveData.Enabled = Sta

    Select Case strMode
           Case Is = "1" 'แก้ไขข้อมูล                        
           Case Is = "2" 'มุมมองข้อมูล
                btnSaveData.Enabled = False
    End Select

End Sub

Private Sub btnSeekExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekExit.Click
  StateLockFindDept(True)
  gpbSeek.Visible = False
  IsShowSeek = False
End Sub

Private Sub txtSize_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSize.GotFocus
  txtSize.SelectAll()
End Sub

Private Sub txtSize_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSize.KeyDown

Dim intChkPoint As Integer
    With txtSize
         Select Case e.KeyCode
                Case Is = 35 'ปุ่ม End 
                Case Is = 36 'ปุ่ม Home
                Case Is = 37 'ลูกศรซ้าย
                     If .SelectionStart = 0 Then
                     End If
                Case Is = 38 'ปุ่มลูกศรขึ้น                                  
                Case Is = 39 'ปุ่มลูกศรขวา
                     If .SelectionLength = .Text.Trim.Length Then
                        txtSizeDesc.Focus()
                     Else
                          intChkPoint = .Text.Trim.Length
                          If .SelectionStart = intChkPoint Then
                             txtSizeDesc.Focus()
                          End If
                     End If
               Case Is = 40 'ปุ่มลง
                    txtPrDate.Focus()
               Case Is = 113 'ปุ่ม F2
                    .SelectionStart = .Text.Trim.Length
         End Select
    End With
End Sub

Private Sub txtSize_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSize.KeyPress

  Select Case Asc(e.KeyChar)
         Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข 0 57คือเลข 9ตามลำดับ
              e.Handled = False

         Case 8, 46, 45  ' ปุ่ม Backspace = 8, ปุ่มDelete = 46 , ปุ่ม 45 = ขีดกลาง
              e.Handled = False

         Case 13 'ปุ่ม Enter = 13
              e.Handled = False
              txtSizeDesc.Focus()

         Case Else
              e.Handled = True
              MessageBox.Show("สามารถกดได้แค่ตัวเลข")
  End Select
End Sub

Private Sub txtSize_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSize.LostFocus
  With txtSize
       .Text = .Text.ToString.Trim.ToUpper
  End With
End Sub

Private Sub txtSizeDesc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeDesc.GotFocus
   txtSizeDesc.SelectAll()
End Sub

Private Sub txtSizeDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSizeDesc.KeyDown
Dim intChkPoint As Integer

    With txtSizeDesc

               Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                        txtSize.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                   
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtSetQty.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                            txtSetQty.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง     
                                  txtRecvDate.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub txtSizeDesc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSizeDesc.KeyPress
    Select Case Asc(e.KeyChar)
           Case 45 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
                e.Handled = False

           Case 8, 46 ' Backspace = 8,  Delete = 46
                e.Handled = False
           Case 13   'Enter = 13
                e.Handled = False
                txtSetQty.Focus()
           Case Else
                e.Handled = True
                MessageBox.Show("กรุณาระบุข้อมูลเป็นภาษาอังกฤษ")
   End Select
End Sub

Private Sub txtSizeDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeDesc.LostFocus
    With txtSizeDesc
            .Text = .Text.ToString.Trim.ToUpper
    End With
End Sub

Private Sub txtSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSizeQty.GotFocus
    With mskSizeQty
         txtSizeQty.SendToBack()
         .BringToFront()
         .Focus()
    End With
End Sub

Private Sub mskSizeQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.GotFocus

Dim i, x As Integer
Dim strTmp As String = ""
Dim strMerge As String = ""

    With mskSizeQty

          If txtSizeQty.Text.ToString.Trim <> "0" Then
             x = Len(txtSizeQty.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtSizeQty.Text.ToString, i, 1)

                                Select Case strTmp
                                       Case Is = "_"
                                       Case Else

                                            If InStr("0123456789.", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If

                                End Select

                         Next i


                         Select Case strMerge.IndexOf(".")

                                Case Is = -1
                                     .SelectionStart = 0
                                Case Is = 1
                                     .SelectionStart = 1
                                Case Is = 2
                                     .SelectionStart = 0
                                Case Is = 3
                                     .SelectionStart = 0
                                Case Else
                                     .SelectionStart = 0
                        End Select

                 .SelectedText = strMerge
             End If

        .SelectAll()
   End With

End Sub

Private Sub mskSizeQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSizeQty.KeyDown
Dim intChkPoint As Integer

    With mskSizeQty

            Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                    txtSetQty.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtWeight.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                            txtWeight.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง    
                                  txtPrDate.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length

            End Select

    End With

End Sub

Private Sub mskSizeQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSizeQty.KeyPress

 Select Case e.KeyChar

                Case Is = Chr(13)
                            txtWeight.Focus()

               Case Is = Chr(46) 'เครื่องหมายจุลภาค(.)                                      
                           mskSizeQty.SelectionStart = 3

    End Select

End Sub

Private Sub mskSizeQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSizeQty.LostFocus

  Dim i, x As Integer
  Dim z As Double

  Dim strTmp As String = ""
  Dim strMerge As String = ""
  Dim intFull As Integer

      With mskSizeQty

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = ","
                                       Case Is = "+"
                                       Case Is = "_"
                                       Case Else

                                            If InStr("0123456789.", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If

                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskSizeQty.Text = ""
                    z = CDbl(strMerge)

                    intFull = Int(z)

                    If (z - intFull) > 0 Then
                       txtSizeQty.Text = z.ToString("#,##0.0")
                    Else
                        txtSizeQty.Text = z.ToString("0.0")
                    End If

                Catch ex As Exception
                    txtSizeQty.Text = "0.0"
                    mskQty.Text = ""
               End Try

       mskQty.SendToBack()
       txtSizeQty.BringToFront()
    End With
End Sub

Private Sub txtWeight_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWeight.GotFocus
   With mskWeight
          txtWeight.SendToBack()
         .BringToFront()
         .Focus()
    End With
End Sub

Private Sub mskWeight_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskWeight.GotFocus
Dim i, x As Integer

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskWeight

        If txtWeight.Text <> "0.00" Then

                        x = Len(txtWeight.Text.ToString)

                        For i = 1 To x

                             strTmp = Mid(txtWeight.Text.ToString, i, 1)
                              Select Case strTmp
                                     Case Is = "_"
                                     Case Else

                                         If InStr(",.0123456789", strTmp) > 0 Then
                                            strMerge = strMerge & strTmp
                                         End If

                                End Select
                         Next i

                        Select Case strMerge.IndexOf(".")
                               Case Is = 5
                                     .SelectionStart = 0
                               Case Is = 3
                                     .SelectionStart = 2
                               Case Is = 2
                                     .SelectionStart = 3
                               Case Is = 1
                                     .SelectionStart = 4
                               Case Else
                                     .SelectionStart = 0
                        End Select
                        .SelectedText = strMerge

                End If
        .SelectAll()
End With
End Sub

Private Sub mskWeight_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskWeight.KeyDown
Dim intChkPoint As Integer

    With mskWeight

            Select Case e.KeyCode
                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย
                        If .SelectionStart = 0 Then
                           txtSizeQty.Focus()
                        End If
                  Case Is = 38 'ปุ่มลูกศรขึ้น    
                         txtSizeDesc.Focus()
                  Case Is = 39 'ปุ่มลูกศรขวา
                       If .SelectionLength = .Text.Trim.Length Then
                          txtPrDate.Focus()
                       Else
                            intChkPoint = .Text.Trim.Length
                            If .SelectionStart = intChkPoint Then
                               txtPrDate.Focus()
                            End If
                       End If
                 Case Is = 40 'ปุ่มลง    
                        txtPrDate.Focus()
                 Case Is = 113 'ปุ่ม F2
                        .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub mskWeight_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskWeight.KeyPress
   Select Case e.KeyChar
          Case Is = Chr(13)
                  txtPrDate.Focus()

          Case Is = Chr(46)   'เครื่องหมายจุลภาค(.)                                      
                  mskWeight.SelectionStart = 6
  End Select
End Sub

Private Sub mskWeight_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskWeight.LostFocus

Dim i, x As Integer
Dim z As Double

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskWeight

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = ","
                                       Case Is = "+"
                                       Case Is = "_"
                                       Case Else
                                            If InStr(".0123456789", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If
                               End Select
                               strTmp = ""
                        Next i
                Try

                    mskWeight.Text = ""
                    z = CDbl(strMerge)
                    txtWeight.Text = z.ToString("#,##0.00")

                Catch ex As Exception
                    txtWeight.Text = "0.00"
                    mskWeight.Text = ""
               End Try

    mskWeight.SendToBack()
    txtWeight.BringToFront()
End With
End Sub

Private Sub txtPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrice.GotFocus
    With mskPrice
         txtPrice.SendToBack()
         .BringToFront()
         .Focus()
    End With
End Sub

Private Sub mskPrice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrice.GotFocus

Dim i, x As Integer

Dim strTmp As String = ""
Dim strMerge As String = ""

    With mskPrice

        If txtPrice.Text <> "0.00" Then

                        x = Len(txtPrice.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtPrice.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = "_"
                                       Case Else
                                            If InStr(",.0123456789", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If
                               End Select

                         Next i


                        Select Case strMerge.IndexOf(".")   'หาตำแหน่งที่พบเป็นครั้งแรก

                               Case Is = 7
                                            .SelectionStart = 0
                               Case Is = 6
                                            .SelectionStart = 1
                               Case Is = 5
                                            .SelectionStart = 2
                               Case Is = 3
                                            .SelectionStart = 3
                               Case Is = 2
                                            .SelectionStart = 5
                               Case Is = 1
                                            .SelectionStart = 7
                               Case Else
                                            .SelectionStart = 7
                        End Select

                       .SelectedText = strMerge
                End If
        .SelectAll()
     End With

End Sub

Private Sub mskPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskPrice.KeyDown

    Dim intChkPoint As Integer
        With mskPrice

               Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                    txtPrDoc.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น    
                                  txtFcDate.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtWd.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtWd.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง    
                                  txtSuplier.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
             End Select
    End With
End Sub

Private Sub mskPrice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrice.KeyPress
   Select Case e.KeyChar

          Case Is = Chr(13)
               txtWd.Focus()
          Case Is = Chr(46) 'เครื่องหมายจุลภาค(.)                                      
               mskPrice.SelectionStart = 7

    End Select

End Sub

Private Sub mskPrice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrice.LostFocus

Dim i, x As Integer
Dim z As Double

Dim strTmp As String = ""
Dim strMerge As String = ""

    With mskPrice

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = ","
                                       Case Is = "+"
                                       Case Is = "_"
                                       Case Else
                                            If InStr(".0123456789", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If
                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskPrice.Text = ""
                    z = CDbl(strMerge)
                    txtPrice.Text = z.ToString("#,##0.00")


                Catch ex As Exception
                    txtPrice.Text = "0.00"
                    mskPrice.Text = ""
               End Try

        mskPrice.SendToBack()
        txtPrice.BringToFront()
      End With
End Sub

Private Sub txtPrDoc_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrDoc.GotFocus
   txtPrDoc.SelectAll()
End Sub

Private Sub txtPrDoc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPrDoc.KeyDown

  Dim intChkPoint As Integer

      With txtPrDoc

               Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย
                               If .SelectionStart = 0 Then
                                    txtFcDate.Focus()
                               End If
                      Case Is = 38 'ปุ่มลูกศรขึ้น    
                                  txtPrDate.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา
                                If .SelectionLength = .Text.Trim.Length Then
                                   txtPrice.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtPrice.Focus()
                                    End If

                                End If
                      Case Is = 40 'ปุ่มลง    
                                txtImpt.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select
    End With

End Sub

Private Sub txtPrDoc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPrDoc.KeyPress

  Select Case Asc(e.KeyChar)

         Case 45 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
             e.Handled = False

         Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

         Case 13   'Enter = 13
              e.Handled = False
              txtPrice.Focus()

         Case Else
              e.Handled = True
              MessageBox.Show("กรุณาระบุข้อมูลเป็นภาษาอังกฤษ")
  End Select

End Sub

Private Sub txtPrDoc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrDoc.LostFocus
  With txtPrDoc
       .Text = .Text.ToString.Trim.ToUpper
  End With
End Sub

Private Sub txtRmk_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRmk.GotFocus
  txtRmk.SelectAll()
End Sub

Private Sub txtRmk_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtRmk.KeyDown

Dim intChkPoint As Integer

    With txtRmk

               Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                    txtInvoice.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                  
                                 txtSuplier.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                            txtSize.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                            txtSize.Focus()
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง                              
                      Case Is = 113 'ปุ่ม F2
                               .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub txtRmk_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRmk.KeyPress
   If e.KeyChar = Chr(13) Then
      txtSize.Focus()
   End If
End Sub

Private Sub txtRmk_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRmk.LostFocus
    With txtRmk
            .Text = .Text.ToString.Trim.ToUpper
    End With
End Sub

Private Sub txtWd_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWd.GotFocus
   With mskWd
        txtWd.SendToBack()
        .BringToFront()
        .Focus()
   End With
End Sub

Private Sub mskWd_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskWd.GotFocus

  Dim i, x As Integer

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskWd

           If txtWd.Text <> "0.00" Then

                        x = Len(txtWd.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtWd.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = "_"
                                       Case Else
                                            If InStr(",.0123456789", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If
                                End Select
                         Next i

                        Select Case strMerge.IndexOf(".")
                               Case Is = 5
                                    .SelectionStart = 0
                               Case Is = 3
                                    .SelectionStart = 2
                               Case Is = 2
                                    .SelectionStart = 3
                               Case Is = 1
                                    .SelectionStart = 4
                               Case Else
                                    .SelectionStart = 0
                        End Select

                        .SelectedText = strMerge

                End If
          .SelectAll()
       End With
End Sub

Private Sub mskWd_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskWd.KeyDown

 Dim intChkPoint As Integer
     With mskWd

               Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                    txtHg.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtLg.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtLg.Focus()
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง                                      
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length

            End Select

    End With

End Sub

Private Sub mskWd_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskWd.KeyPress
    Select Case e.KeyChar
           Case Is = Chr(13)
                   txtLg.Focus()
           Case Is = Chr(46) 'เครื่องหมายจุลภาค(.)                                      
                   mskWd.SelectionStart = 6
    End Select
End Sub

Private Sub mskWd_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskWd.LostFocus

Dim i, x As Integer
Dim z As Double

Dim strTmp As String = ""
Dim strMerge As String = ""

    With mskWd

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = ","
                                          Case Is = "+"
                                          Case Is = "_"
                                          Case Else

                                                    If InStr(".0123456789", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskWd.Text = ""
                    z = CDbl(strMerge)
                    txtWd.Text = z.ToString("#,##0.00")


                Catch ex As Exception
                    txtWd.Text = "0.00"
                    mskWd.Text = ""
               End Try

mskWd.SendToBack()
txtWd.BringToFront()

End With

End Sub

Private Sub txtLg_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLg.GotFocus

     With mskLg
          txtLg.SendToBack()
         .BringToFront()
         .Focus()

    End With

End Sub

Private Sub mskLg_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskLg.GotFocus

Dim i, x As Integer

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskLg

        If txtLg.Text <> "0.00" Then

                        x = Len(txtLg.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtLg.Text.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = "_"
                                          Case Else

                                                    If InStr(",.0123456789", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select
                         Next i

                        Select Case strMerge.IndexOf(".")
                                  Case Is = 5
                                            .SelectionStart = 0
                                  Case Is = 3
                                            .SelectionStart = 2
                                 Case Is = 2
                                            .SelectionStart = 3
                                 Case Is = 1
                                            .SelectionStart = 4
                                 Case Else
                                            .SelectionStart = 0
                        End Select

                        .SelectedText = strMerge

                End If

        .SelectAll()

End With

End Sub

Private Sub mskLg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskLg.KeyDown
Dim intChkPoint As Integer

    With mskLg

            Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                  txtWd.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtHg.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtHg.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง                                      
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length

            End Select

    End With

End Sub

Private Sub mskLg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskLg.KeyPress
     Select Case e.KeyChar

                Case Is = Chr(13)
                            txtHg.Focus()

               Case Is = Chr(46) 'เครื่องหมายจุลภาค(.)                                      
                           mskLg.SelectionStart = 6

    End Select

End Sub

Private Sub mskLg_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskLg.LostFocus
Dim i, x As Integer
Dim z As Double

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskLg

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = ","
                                          Case Is = "+"
                                          Case Is = "_"
                                          Case Else

                                                    If InStr(".0123456789", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskLg.Text = ""
                    z = CDbl(strMerge)
                    txtLg.Text = z.ToString("#,##0.00")


                Catch ex As Exception
                    txtLg.Text = "0.00"
                    mskLg.Text = ""
               End Try

mskLg.SendToBack()
txtLg.BringToFront()

End With

End Sub

Private Sub txtHg_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHg.GotFocus

  With mskHg
          txtHg.SendToBack()
         .BringToFront()
         .Focus()

    End With

End Sub

Private Sub mskHg_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskHg.GotFocus
Dim i, x As Integer

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskHg

        If txtHg.Text <> "0.00" Then

                        x = Len(txtHg.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtHg.Text.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = "_"
                                          Case Else

                                                    If InStr(",.0123456789", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select
                         Next i

                        Select Case strMerge.IndexOf(".")
                                  Case Is = 5
                                            .SelectionStart = 0
                                  Case Is = 3
                                            .SelectionStart = 2
                                 Case Is = 2
                                            .SelectionStart = 3
                                 Case Is = 1
                                            .SelectionStart = 4
                                 Case Else
                                            .SelectionStart = 0
                        End Select

                        .SelectedText = strMerge

                End If

        .SelectAll()

End With

End Sub

Private Sub mskHg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskHg.KeyDown
Dim intChkPoint As Integer

    With mskHg

            Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                    txtLg.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtImpt.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtImpt.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง                                      
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length

            End Select

    End With

End Sub

Private Sub mskHg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskHg.KeyPress

    Select Case e.KeyChar

           Case Is = Chr(13)
                   txtImpt.Focus()

          Case Is = Chr(46) 'เครื่องหมายจุลภาค(.)                                      
                  mskHg.SelectionStart = 6

    End Select

End Sub

Private Sub mskHg_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskHg.LostFocus

Dim i, x As Integer
Dim z As Double

Dim strTmp As String = ""
Dim strMerge As String = ""

With mskHg

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = ","
                                       Case Is = "+"
                                       Case Is = "_"
                                       Case Else

                                            If InStr(".0123456789", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If

                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskHg.Text = ""
                    z = CDbl(strMerge)
                    txtHg.Text = z.ToString("#,##0.00")


                Catch ex As Exception
                    txtHg.Text = "0.00"
                    mskHg.Text = ""
               End Try

   mskHg.SendToBack()
   txtHg.BringToFront()

End With

End Sub

Private Sub btnSeekSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeekSave.Click
  CheckSubDataBefSave()
End Sub

Private Sub CheckSubDataBefSave()
  Dim i As Integer

                 If txtSize.Text.ToString.Trim <> "" Then

                           If txtSizeDesc.Text.ToString.Trim <> "" Then

                                   If cmbType.Text.ToString.Trim <> "" Then

                                            If gpbSeek.Text = "เพิ่มข้อมูล" Then
                                               SaveSubRecord()
                                               ReSizeSort(txtEqpId.Text.ToUpper.Trim)

                                            Else
                                               EditSubRecord()
                                               ReSizeSort(txtEqpId.Text.ToUpper.Trim)

                                            End If

                                            ShowScrapItem()
                                            'ShowSumItem()

                                            '------------------------------ค้นหารหัสที่เพิ่มเข้าไปใหม่------------------------------------------

                                             For i = 0 To dgvSize.Rows.Count - 1

                                                        If dgvSize.Rows(i).Cells(4).Value.ToString = txtSize.Text.ToString.Trim Then
                                                            dgvSize.CurrentCell = dgvSize.Item(5, i)
                                                            dgvSize.Focus()
                                                            Exit For

                                                         End If

                                            Next i

                                            StateLockFindDept(True)
                                            gpbSeek.Visible = False
                                            IsShowSeek = False

                                     Else

                                            MsgBox("โปรดระบุข้อมูลประเภทอุปกรณ์  " _
                                                            & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")

                                            cmbType.DroppedDown = True
                                            cmbType.Focus()

                                       End If


                            Else
                                 MsgBox("โปรดระบุข้อมูล ชุดโมล์ด/GroupSize" _
                                                  & " ก่อนเพิ่มข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                 txtSizeDesc.Focus()

                            End If

                 Else

                          MsgBox("โปรดระบุข้อมูล SIZE  " _
                                        & " ก่อนเพิ่มข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                          txtSize.Focus()

                 End If


End Sub

Private Function SaveSubRecord() As Boolean

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlSelc As String
Dim strSqlCmd As String

Dim datSave As Date = Now()
Dim strDate As String = ""

Dim strPrDate As String = ""
Dim strRecvDate As String = ""
Dim strFcDate As String = ""
Dim strEqpType As String = ""

     Try

        With Conn
             If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .Open()
         End With

       '------------------------------------------------------------เช็คข้อมูล่ก่อนว่ามีอยู่หรือเปล่า----------------------------------------------------------------------

        strSqlSelc = "SELECT size_id  FROM tmp_eqptrn " _
                            & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                            & " AND size_id = '" & txtSize.Text.ToString.Trim & "'" _
                            & " AND size_desc = '" & txtSizeDesc.Text.ToString.Trim & "'"

            With Rsd

                 .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                 .LockType = ADODB.LockTypeEnum.adLockOptimistic
                 .Open(strSqlSelc, Conn, , , )

                  If .RecordCount <> 0 Then
                     SaveSubRecord = False
                     MsgBox("SIZE : " & txtSize.Text.ToString.Trim & " มีในระบบเเล้ว โปรดระบุ SIZE อื่น", MsgBoxStyle.Critical, "ผิดพลาด")

                  Else

                        '---------------------------------------- วันที่เปิดซื้อ --------------------------------------------------

                        If txtPrDate.Text <> "__/__/____" Then

                                strPrDate = Mid(txtPrDate.Text.ToString, 7, 4) & "-" _
                                                    & Mid(txtPrDate.Text.ToString, 4, 2) & "-" _
                                                     & Mid(txtPrDate.Text.ToString, 1, 2)
                                strPrDate = "'" & SaveChangeEngYear(strPrDate) & "'"

                        Else
                                strPrDate = "NULL"
                        End If

                         '----------------------------------------วันที่รับเข้า---------------------------------------------------

                        If txtRecvDate.Text <> "__/__/____" Then

                                strRecvDate = Mid(txtRecvDate.Text.ToString, 7, 4) & "-" _
                                                          & Mid(txtRecvDate.Text.ToString, 4, 2) & "-" _
                                                          & Mid(txtRecvDate.Text.ToString, 1, 2)
                                strRecvDate = "'" & SaveChangeEngYear(strRecvDate) & "'"

                        Else
                                strRecvDate = "NULL"
                        End If

                          '----------------------------------------วันที่นัดเข้า---------------------------------------------------

                        If txtFcDate.Text <> "__/__/____" Then

                                strFcDate = Mid(txtFcDate.Text.ToString, 7, 4) & "-" _
                                                     & Mid(txtFcDate.Text.ToString, 4, 2) & "-" _
                                                     & Mid(txtFcDate.Text.ToString, 1, 2)
                                strFcDate = "'" & SaveChangeEngYear(strFcDate) & "'"

                        Else
                               strFcDate = "NULL"
                        End If

                        '------------------------------------กำหนดกลุ่มประเภท---------------------------------------------------------------------

                                 Select Case cmbType.Text.ToString.Trim

                                        Case Is = "โมลด์ต่างประเทศ"
                                             strEqpType = "EXP"
                                        Case Is = "โมลด์ในประเทศ"
                                             strEqpType = "LCA"

                                 End Select

                strSqlCmd = "INSERT INTO tmp_eqptrn " _
                                      & "(user_id,size_id,size_desc,size_qty,set_qty" _
                                      & ",weight,dimns,backgup,price,men_rmk,[group],eqp_id" _
                                      & ",delvr_sta,sent_sta,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                      & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                                      & ")" _
                                      & " VALUES (" _
                                      & "'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                      & ",'" & ReplaceQuote(txtSize.Text.ToString.Trim) & "'" _
                                      & ",'" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                                      & "," & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                                      & "," & ChangFormat(txtSetQty.Text.ToString.Trim) _
                                      & "," & ChangFormat(txtWeight.Text.ToString.Trim) _
                                      & ",'" & txtWd.Text.ToString.Trim & " x " & _
                                                    txtLg.Text.ToString.Trim & " x " & _
                                                    txtHg.Text.ToString.Trim & "'" _
                                      & ",'" & ReplaceQuote(txtInvoice.Text.ToString.Trim) & "'" _
                                      & "," & ChangFormat(txtPrice.Text.ToString.Trim) _
                                      & ",'" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                                      & ",'" & "" & "'" _
                                      & ",'" & "" & "'" _
                                      & ",'" & "0" & "'" _
                                      & ",'" & "0" & "'" _
                                      & "," & strPrDate _
                                      & ",'" & ReplaceQuote(txtPrDoc.Text.ToString.Trim) & "'" _
                                      & "," & strRecvDate _
                                      & "," & "0" _
                                      & "," & "0" _
                                      & "," & strFcDate _
                                      & ",'" & ReplaceQuote(txtImpt.Text.ToString.Trim) & "'" _
                                      & ",'" & ReplaceQuote(txtSuplier.Text.ToString.Trim) & "'" _
                                      & ",'" & strEqpType & "'" _
                                      & ",'" & "" & "'" _
                                      & ",'" & "" & "'" _
                                      & ",'" & "" & "'" _
                                      & ",'" & "" & "'" _
                                      & "," & ChangFormat(txtMouth_mold.Text.Trim) _
                                      & ")"

                         Conn.Execute(strSqlCmd)

                         SaveSubRecord = True

                  End If

                 .ActiveConnection = Nothing
                 .Close()

          End With

          Rsd = Nothing

    Conn.Close()
    Conn = Nothing

     Catch ex As Exception
           MsgBox("พบข้อผิดพลาดขณะทำการบันทึก โปรดดำเนินการใหม่อีกครั้ง", MsgBoxStyle.Critical, "ผิดพลาด")
           MsgBox(ex.Message)
     End Try

End Function

Private Function EditSubRecord() As Boolean

  Dim Conn As New ADODB.Connection
  Dim strSqlCmd As String

  Dim strPrDate As String = ""
  Dim strRecvDate As String = ""
  Dim strFcDate As String = ""
  Dim strEqpType As String = ""

     Try

        With Conn
             If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .Open()
         End With

                        '----------------------------------------วันที่เปิดซื้อ---------------------------------------------------

                        If txtPrDate.Text <> "__/__/____" Then

                           strPrDate = Mid(txtPrDate.Text.ToString, 7, 4) & "-" _
                                                    & Mid(txtPrDate.Text.ToString, 4, 2) & "-" _
                                                     & Mid(txtPrDate.Text.ToString, 1, 2)
                           strPrDate = "'" & SaveChangeEngYear(strPrDate) & "'"

                        Else
                           strPrDate = "NULL"

                        End If

                         '----------------------------------------วันที่รับเข้า---------------------------------------------------

                        If txtRecvDate.Text <> "__/__/____" Then

                           strRecvDate = Mid(txtRecvDate.Text.ToString, 7, 4) & "-" _
                                                          & Mid(txtRecvDate.Text.ToString, 4, 2) & "-" _
                                                          & Mid(txtRecvDate.Text.ToString, 1, 2)
                           strRecvDate = "'" & SaveChangeEngYear(strRecvDate) & "'"

                        Else
                           strRecvDate = "NULL"

                        End If

                         '----------------------------------------วันที่นัดเข้า---------------------------------------------------

                        If txtFcDate.Text <> "__/__/____" Then

                           strFcDate = Mid(txtFcDate.Text.ToString, 7, 4) & "-" _
                                                     & Mid(txtFcDate.Text.ToString, 4, 2) & "-" _
                                                     & Mid(txtFcDate.Text.ToString, 1, 2)
                           strFcDate = "'" & SaveChangeEngYear(strFcDate) & "'"

                        Else
                           strFcDate = "NULL"

                        End If

                        '------------------------------------กำหนดกลุ่มประเภท---------------------------------------------------------------------

                          Select Case cmbType.Text.ToString.Trim

                                 Case Is = "โมลด์ต่างประเทศ"
                                      strEqpType = "EXP"

                                 Case Is = "โมลด์ในประเทศ"
                                      strEqpType = "LCA"

                           End Select


               strSqlCmd = "UPDATE  tmp_eqptrn SET size_desc ='" & ReplaceQuote(txtSizeDesc.Text.ToString.Trim) & "'" _
                            & "," & "size_qty =" & ChangFormat(txtSizeQty.Text.ToString.Trim) _
                            & "," & "weight =" & ChangFormat(txtWeight.Text.ToString.Trim) _
                            & "," & "dimns ='" & txtWd.Text.ToString.Trim & " x " & _
                                                                  txtLg.Text.ToString.Trim & " x " & _
                                                                  txtHg.Text.ToString.Trim & "'" _
                            & "," & "pr_doc ='" & ReplaceQuote(txtPrDoc.Text.ToString.Trim) & "'" _
                            & "," & "price =" & ChangFormat(txtPrice.Text.ToString.Trim) _
                            & "," & "men_rmk ='" & ReplaceQuote(txtRmk.Text.ToString.Trim) & "'" _
                            & "," & "set_qty =" & ChangFormat(txtSetQty.Text.ToString.Trim) _
                            & "," & "pr_date =" & strPrDate _
                            & "," & "recv_date =" & strRecvDate _
                            & "," & "fc_date =" & strFcDate _
                            & "," & "backgup ='" & ReplaceQuote(txtInvoice.Text.ToString.Trim) & "'" _
                            & "," & "impt_id ='" & ReplaceQuote(txtImpt.Text.ToString.Trim) & "'" _
                            & "," & "sup_name ='" & ReplaceQuote(txtSuplier.Text.ToString.Trim) & "'" _
                            & "," & "lp_type ='" & strEqpType & "'" _
                            & "," & "mouth_long ='" & txtMouth_mold.Text.Trim & "'" _
                            & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                            & " AND size_id ='" & txtSize.Text.ToString.Trim & "'" _
                            & " AND size_desc = '" & txtSizeDesc.Text.ToString.Trim & "'"

             Conn.Execute(strSqlCmd)

 Conn.Close()
 Conn = Nothing

   Catch ex As Exception
         MsgBox("พบข้อผิดพลาดขณะทำการบันทึก โปรดดำเนินการใหม่อีกครั้ง", MsgBoxStyle.Critical, "ผิดพลาด")
         MsgBox(ex.Message)
   End Try

End Function

Private Sub ShowScrapItem()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlCmdSelc As String

Dim strSta As String = ""
Dim dubQty As Double
Dim dubAmt As Double
Dim user As String = frmMainPro.lblLogin.Text.ToString.Trim
Dim mold_id As String
Dim mold_size As String
Dim strArr() As String

Dim sngSetQty As Single   ' จำนวนเซ็ต

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
    End With
    'v_tmp_eqptrn  & " ORDER BY size_desc, tmp_newsize"

    strSqlCmdSelc = "SELECT * FROM v_tmpeqptrn_newsize (NOLOCK)" _
                              & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                              & " ORDER BY tmp_newsize"

         With Rsd

              .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
              .LockType = ADODB.LockTypeEnum.adLockOptimistic
              .Open(strSqlCmdSelc, Conn, , , )

              dgvSize.Rows.Clear()
              dgvSize.ScrollBars = ScrollBars.None 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

                  If .RecordCount <> 0 Then

                                Do While Not .EOF()
                                   mold_id = .Fields("eqp_id").Value.ToString.Trim
                                   strArr = Split(.Fields("size_desc").Value.ToString.Trim, "-")  'ตัด array ออกมา
                                   mold_size = .Fields("size_id").Value.ToString.Trim + strArr(0)
                                   dgvSize.Rows.Add( _
                                                         IIf(.Fields("delvr_sta").Value.ToString.Trim = "1", My.Resources.accept, My.Resources.blank), _
                                                         IIf(Find_fixmold(user, mold_id, mold_size) = "1", My.Resources.accept, My.Resources.blank), _
                                                         .Fields("size_id").Value.ToString.Trim, _
                                                         .Fields("size_act").Value.ToString.Trim, _
                                                         .Fields("size_desc").Value.ToString.Trim, _
                                                         .Fields("impt_id").Value.ToString.Trim, _
                                                         Format(.Fields("set_qty").Value, "#,##0.0"), _
                                                         Format(.Fields("size_qty").Value, "#,##0.0"), _
                                                         .Fields("weight").Value, _
                                                         .Fields("dimns").Value.ToString.Trim, _
                                                         .Fields("price").Value, _
                                                          Mid(.Fields("pr_date").Value.ToString.Trim, 1, 10), _
                                                         .Fields("pr_doc").Value.ToString.Trim, _
                                                         .Fields("sup_name").Value.ToString.Trim, _
                                                         .Fields("lptype").Value.ToString.Trim, _
                                                          Mid(.Fields("fc_date").Value.ToString.Trim, 1, 10), _
                                                          Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                                         .Fields("backgup").Value.ToString.Trim, _
                                                         .Fields("ord_rep").Value, _
                                                         .Fields("ord_qty").Value, _
                                                         .Fields("men_rmk").Value.ToString.Trim _
                                                     )

                                        sngSetQty = sngSetQty + .Fields("set_qty").Value
                                        dubQty = dubQty + .Fields("ord_qty").Value
                                        dubAmt = dubAmt + .Fields("price").Value

                                        .MoveNext()

                                Loop

                                        txtSet.Text = sngSetQty.ToString.Trim        'จำนวน SET
                                        lblOrdQty.Text = Format(dubQty, "#,##0")     'รวมคู่ลงผลิต
                                        lblAmt.Text = Format(dubAmt, "#,##0.00")     'รวมราคาอุปกรณ์

                      Else
                            txtSet.Text = "0"
                            lblOrdQty.Text = "0"
                            lblAmt.Text = "0.00"

                     End If

                    .ActiveConnection = Nothing
                    .Close()
                    Rsd = Nothing

                    dgvSize.ScrollBars = ScrollBars.Both 'กัน ScrollBars ของ DataGrid Refresh ไม่ทัน

        End With

    Conn.Close()
    Conn = Nothing

End Sub

Private Function Find_fixmold(ByVal user As String, ByVal idMold As String, ByVal mSize As String) As String

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim sqlSelc As String

     With Conn
          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

     sqlSelc = "SELECT fix_sta FROM tmp_fixeqptrn " _
                  & " WHERE user_id='" & user & "'" _
                  & " AND eqp_id='" & idMold & "'" _
                  & " AND size_id ='" & mSize & "'"

     With Rsd

         .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
         .LockType = ADODB.LockTypeEnum.adLockOptimistic
         .Open(sqlSelc, Conn, , , )

         If .RecordCount <> 0 Then
            Return .Fields("fix_sta").Value.ToString.Trim
         Else
             Return ""
         End If

       .ActiveConnection = Nothing
       .Close()
     End With

  Conn.Close()
End Function

Private Sub btnSaveData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData.Click
   CheckDataBeforeSave()
End Sub

Private Sub CheckDataBeforeSave()

  Dim intListWc As Integer = dgvSize.Rows.Count
  Dim strProd As String = ""
  Dim strProdNm As String = ""
  Dim bytConSave As Byte

      If cmbGp.Text.ToString.Trim <> "" Then

           If cmbPart.Text.ToString.Trim <> "" Then

                     If txtEqpId.Text.ToString.Trim <> "" Then

                                   If txtEqpName.Text.ToString.Trim <> "" Then

                                                   If intListWc > 0 Then
                                                                  bytConSave = MsgBox("คุณต้องการบันทึกข้อมูลใช่หรือไม่!" _
                                                                 , MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Information, "Save Data!!!")

                                                                    If bytConSave = 6 Then

                                                                          Select Case Me.Text

                                                                                 Case Is = "เพิ่มข้อมูล"
                                                                                      SaveNewRecord()

                                                                                 Case Else
                                                                                      SaveEditRecord()

                                                                          End Select

                                                                     Else
                                                                          dgvSize.Focus()

                                                                     End If

                                                    Else
                                                          MsgBox("โปรดระบุข้อมูลรายการ SIZE " & vbNewLine _
                                                                   & " โปรดกำหนดราคาขายก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")

                                                                 ClearSubData()
                                                                 ShowResvrd()
                                                                 gpbSeek.Text = "เพิ่มข้อมูล"
                                                                 txtSize.ReadOnly = False
                                                                 txtSizeDesc.ReadOnly = False

                                                    End If


                                           Else
                                                MsgBox("โปรดระบุข้อมูลรายละเอียดอุปกรณ์  " & vbNewLine _
                                                        & " โปรดกำหนดราคาขายก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                                        txtEqpName.Focus()
                                           End If


                                  Else
                                        MsgBox("โปรดระบุข้อมูลรหัสอุปกรณ์  " _
                                                & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")
                                                txtEqpId.Focus()

                                  End If


              Else

                             MsgBox("โปรดระบุข้อมูลชิ้นส่วนที่ผลิต  " _
                                              & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")

                             cmbPart.DroppedDown = True
                             cmbPart.Focus()

             End If


Else

     MsgBox("โปรดระบุข้อมูลกลุ่มอุปกรณ์  " _
                        & " ก่อนบันทึกข้อมูล", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Please Input Data First!")

    cmbGp.DroppedDown = True
    cmbGp.Focus()

End If

End Sub

Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click

   If dgvSize.Rows.Count > 0 Then
      ClearSubData()
      ShowResvrd()
      gpbSeek.Text = "แก้ไขข้อมูล"
      txtSize.ReadOnly = True
      txtSizeDesc.ReadOnly = True
      CallSizeEditData()
    Else
         MsgBox("ไม่มีรายการ SIZE ที่ต้องการแก้ไข!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
         dgvSize.Focus()
    End If

End Sub

Private Sub CallSizeEditData()

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlSelc As String

Dim strWd As String = ""
Dim strLg As String = ""
Dim strHg As String = ""

Dim strCode As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(2).Value.ToString.Trim
Dim strLot As String = dgvSize.Rows(dgvSize.CurrentRow.Index).Cells(4).Value.ToString.Trim

    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
   End With

              strSqlSelc = "SELECT * FROM v_tmp_eqptrn (NOLOCK)" _
                                   & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                   & " AND size_id ='" & strCode & "'" _
                                   & " AND size_desc = '" & strLot & "'"

              With Rsd

                   .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                   .LockType = ADODB.LockTypeEnum.adLockOptimistic
                   .Open(strSqlSelc, Conn, , , )

                   If .RecordCount <> 0 Then

                      txtSize.Text = .Fields("size_id").Value.ToString.Trim
                      txtSizeDesc.Text = .Fields("size_desc").Value.ToString.Trim
                      txtSetQty.Text = Format(.Fields("set_qty").Value, "#,##0.0")
                      txtSizeQty.Text = Format(.Fields("size_qty").Value, "#,##0.0")
                      txtWeight.Text = Format(.Fields("weight").Value, "#,##0.00")

                      If .Fields("pr_date").Value.ToString.Trim <> "" Then
                         txtPrDate.Text = Mid(.Fields("pr_date").Value.ToString.Trim, 1, 10)
                      Else
                         txtPrDate.Text = "__/__/____"
                      End If

                      If .Fields("recv_date").Value.ToString.Trim <> "" Then
                         txtRecvDate.Text = Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10)
                      Else
                         txtRecvDate.Text = "__/__/____"
                      End If

                      If .Fields("fc_date").Value.ToString.Trim <> "" Then
                         txtFcDate.Text = Mid(.Fields("fc_date").Value.ToString.Trim, 1, 10)
                      Else
                            txtFcDate.Text = "__/__/____"
                      End If

                      txtPrice.Text = Format(.Fields("price").Value, "#,##0.00")
                      txtPrDoc.Text = .Fields("pr_doc").Value.ToString.Trim
                      txtRmk.Text = .Fields("men_rmk").Value.ToString.Trim

                      cmbType.Text = .Fields("lptype").Value.ToString.Trim
                      txtSuplier.Text = .Fields("sup_name").Value.ToString.Trim
                      txtImpt.Text = .Fields("impt_id").Value.ToString.Trim
                      txtInvoice.Text = .Fields("backgup").Value.ToString.Trim

                      If .Fields("dimns").Value.ToString.Trim <> "" Then
                         RetrnDiamss(.Fields("dimns").Value.ToString.Trim, strWd, strLg, strHg)
                         txtWd.Text = strWd
                         txtLg.Text = strLg
                         txtHg.Text = strHg

                     Else
                           txtWd.Text = "0.00"
                           txtLg.Text = "0.00"
                           txtHg.Text = "0.00"
                     End If

                     If .Fields("mouth_long").Value.ToString.Trim <> "" Then
                        txtMouth_mold.Text = .Fields("mouth_long").Value.ToString.Trim

                     Else
                          txtMouth_mold.Text = "0.00"
                     End If


                End If

                .ActiveConnection = Nothing
                .Close()
             End With

             Rsd = Nothing

Conn.Close()
Conn = Nothing

End Sub

Private Sub dgvSize_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
  dgvSize.Rows(e.RowIndex).Height = 45
End Sub

Private Function RetrnDiamss(ByVal strDia As String, ByRef strW As String, ByRef strL As String, ByRef strH As String) As Boolean

Dim i, x As Integer
Dim strTmp As String = ""
Dim strMerge As String = ""

Dim strDiamns(1) As String
Dim y As Integer = 0

                        x = Len(strDia)

                        For i = 1 To x

                                strTmp = Mid(strDia, i, 1)

                                Select Case strTmp

                                             Case Is = "x"
                                                        strDiamns(y) = strMerge
                                                        y = y + 1
                                                        strMerge = ""

                                             Case Else

                                                    If InStr(",.0123456789", strTmp) > 0 Then
                                                            strMerge = strMerge & strTmp
                                                    End If

                                End Select

                         Next i

strW = strDiamns(0)
strL = strDiamns(1)
strH = strMerge

End Function

Private Sub DeleteSubData()

Dim btyConsider As Byte

Dim strSize As String
Dim strSizeAct As String
Dim strSizeDesc As String

With dgvSize

        If .Rows.Count > 0 Then

                strSizeAct = .Rows(.CurrentRow.Index).Cells(2).Value
                strSize = .Rows(.CurrentRow.Index).Cells(3).Value
                strSizeDesc = .Rows(.CurrentRow.Index).Cells(4).Value

                If strSizeAct <> "" Then

                                              btyConsider = MsgBox("SIZE : " & strSize.ToString.Trim & vbNewLine _
                                                                       & "LOT SIZE : " & strSizeDesc.ToString.Trim & vbNewLine _
                                                                       & "คุณต้องการลบใช่หรือไม่!!", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 _
                                                                        + MsgBoxStyle.Exclamation, "Confirm Delete Data")

                                              If btyConsider = 6 Then

                                                 Dim Conn As New ADODB.Connection
                                                 Dim strSqlCmd As String

                                                            If Conn.State Then Conn.Close()

                                                               Conn.ConnectionString = strConnAdodb
                                                               Conn.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                                               Conn.ConnectionTimeout = 90
                                                               Conn.Open()

                                                                strSqlCmd = "DELETE FROM tmp_eqptrn" _
                                                                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                                                     & " AND size_id ='" & strSizeAct.ToString.Trim & "'" _
                                                                                     & " AND size_desc = '" & strSizeDesc.ToString.Trim & "'"

                                                                Conn.Execute(strSqlCmd)

                                                                '------------------ ลบข้อมูลในตาราง tmp_eqptrn_newsize -------------------

                                                                strSqlCmd = "DELETE FROM tmp_eqptrn_newsize" _
                                                                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                                                                     & " AND size_id ='" & strSizeAct.ToString.Trim & "'" _
                                                                                     & " AND size_desc = '" & strSizeDesc.ToString.Trim & "'"

                                                                Conn.Execute(strSqlCmd)
                                                                Conn.Close()
                                                                Conn = Nothing

                                                               .Rows.RemoveAt(.CurrentRow.Index)
                                                                ShowScrapItem()


                                                Else
                                                   .Focus()

                                                End If


                End If

        Else
             MsgBox("ไม่มีรายการ SIZE ที่ต้องการลบข้อมูล!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Data Not Found!")
             dgvSize.Focus()

        End If


End With

End Sub

Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
  DeleteSubData()
End Sub

Private Sub SaveNewRecord()

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String

Dim dateSave As Date = Now()
Dim strDate As String = ""

Dim strDateDoc As String
Dim strGpType As String = ""
Dim strPartType As String = ""

Dim blnRetureCopyPic As Boolean

Dim Rsd As New ADODB.Recordset

  Try


    With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 150
            .Open()
    End With

        If CheckCodeDuplicate() Then 'ตรวจสอบรหัสซ้ำก่อน

           Conn.BeginTrans()  'เริ่มต้นทรานเซ็กชั่น

           '------------------------------------------------ บันทึกข้อมูลในตาราง eqpmst -------------------------------------------------------

           strDate = dateSave.Date.ToString("yyyy-MM-dd")
                     strDate = SaveChangeEngYear(strDate)

                     strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                          & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                          & Mid(txtBegin.Text.ToString, 1, 2)
                     strDateDoc = SaveChangeEngYear(strDateDoc)

                    '------------------------------------------------------- บันทึกรูปภาพการบรรจุ ------------------------------------------------

                     blnRetureCopyPic = CallCopyPicture(lblPicPath1.Text.ToString.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)
                     If blnRetureCopyPic Then
                        lblPicPath1.Text = PthName

                     Else
                          lblPicName1.Text = ""
                          lblPicPath1.Text = ""
                          picEqp1.Image = Nothing

                     End If

                    '------------------------------------------------บันทึกรูปภาพการภายนอก ภายใน-------------------------------------------------

                     blnRetureCopyPic = CallCopyPicture(lblPicPath2.Text.ToString.Trim, ReturnImageName(lblPicName2.Text.ToString.Trim), lblPicName2.Text.ToString.Trim)
                    If blnRetureCopyPic Then
                       lblPicPath2.Text = PthName

                    Else
                          lblPicName2.Text = ""
                          lblPicPath2.Text = ""
                          picEqp2.Image = Nothing

                    End If

                    '------------------------------------------------บันทึกรูปภาพการชิ้นงาน----------------------------------------------------------

                    blnRetureCopyPic = CallCopyPicture(lblPicPath3.Text.ToString.Trim, ReturnImageName(lblPicName3.Text.ToString.Trim), lblPicName3.Text.ToString.Trim)
                    If blnRetureCopyPic Then
                       lblPicPath3.Text = PthName

                    Else
                          lblPicName3.Text = ""
                          lblPicPath3.Text = ""
                          picEqp3.Image = Nothing

                    End If

                    '-------------------------------------------กำหนดกลุ่มของอุปกรณ์---------------------------------------------------------------

                    Select Case cmbGp.Text.ToString.Trim

                                  Case Is = "MOLD EVA"
                                              strGpType = "A"

                                  Case Is = "MOLD PVC"
                                              strGpType = "B"

                                  Case Is = "MOLD PU"
                                              strGpType = "C"

                    End Select

                    '------------------------------------กำหนดกลุ่มของชิ้นงาน---------------------------------------------------------------------

                    Select Case cmbPart.Text.ToString.Trim

                                  Case Is = "หนังหน้า"
                                              strPartType = "UPPER"
                                  Case Is = "อุปกรณ์"
                                              strPartType = "ACCSSY"
                                  Case Is = "พื้นบน"
                                              strPartType = "SOLE1"
                                  Case Is = "พื้นล่าง"
                                              strPartType = "SOLE2"
                                  Case Is = "รองเท้าสำเร็จรูป"
                                              strPartType = "SOLE3"
                                  Case Is = "การ์ด"
                                              strPartType = "GUARD"
                                  Case Is = "กระดุม"
                                              strPartType = "BUTTON"
                                  Case Is = "จมูก"
                                              strPartType = "NOSE"
                                  Case Is = "ขา"
                                              strPartType = "LEG"
                                  Case Is = "TPR โลโก้"
                                              strPartType = "TPR"

                    End Select


                    strSqlCmd = "INSERT INTO eqpmst " _
                                          & "(prod_sta,fix_sta,[group],eqp_id,eqp_name" _
                                          & ",pi,shoe,ap_code,ap_desc,doc_ref,set_qty" _
                                          & ",part,eqp_type" _
                                          & ",pic_ctain,pic_io,pic_part,remark" _
                                          & ",tech_desc,tech_thk,tech_lg,tech_sht,tech_eva,tech_warm" _
                                          & ",tech_time1,tech_time2,creat_date,pre_date,pre_by,pi_qty" _
                                          & ",eqp_amt,exp_id" _
                                          & ")" _
                                          & " VALUES (" _
                                          & "'" & "0" & "'" _
                                          & ",'" & "0" & "'" _
                                          & ",'" & strGpType.ToString.Trim & "'" _
                                          & ",'" & txtEqpId.Text.ToString.Trim & "'" _
                                          & ",'" & ReplaceQuote(txtEqpName.Text.ToString.Trim) & "'" _
                                          & ",'" & ReplaceQuote(txtOrd.Text.ToString.Trim) & "'" _
                                          & ",'" & ReplaceQuote(txtStyle.Text.ToString.Trim) & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & ReplaceQuote(txtRef.Text.ToString.Trim) & "'" _
                                          & "," & ChangFormat(txtSet.Text.ToString.Trim) _
                                          & ",'" & strPartType.ToString.Trim & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & ReplaceQuote(lblPicName1.Text.ToString.Trim) & "'" _
                                          & ",'" & ReplaceQuote(lblPicName2.Text.ToString.Trim) & "'" _
                                          & ",'" & ReplaceQuote(lblPicName3.Text.ToString.Trim) & "'" _
                                          & ",'" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & "" & "'" _
                                          & ",'" & strDateDoc & "'" _
                                          & ",'" & strDate & "'" _
                                          & ",'" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                          & "," & ChangFormat(lblOrdQty.Text.ToString.Trim) _
                                          & "," & RetrnAmount() _
                                          & ",'" & "" & "'" _
                                          & ")"
                      Conn.Execute(strSqlCmd)

            '------------------------------------------------บันทึกข้อมูลในตาราง eqptrn----------------------------------------------

            strSqlCmd = "INSERT INTO eqptrn " _
                                & " SELECT [group] ='" & strGpType.ToString.Trim & "'" _
                                & ",eqp_id ='" & txtEqpId.Text.ToString.Trim & "'" _
                                & ",size_id,size_desc,size_qty,weight,dimns,backgup,price,men_rmk" _
                                & ",delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                                & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long " _
                                & "FROM tmp_eqptrn" _
                                & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

            Conn.Execute(strSqlCmd)
            Conn.CommitTrans()

            frmMoldInj.lblCmd.Text = txtEqpId.Text.ToString.Trim   'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
            frmMoldInj.Activating()
            Me.Close()
           Else

                MsgBox("รหัสอุปกรณ์ซ้ำ " _
                                   & " โปรดระบุรหัสอุปกรณ์ใหม่", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Data Duplication!")
                txtEqpId.Focus()

           End If

   Conn.Close()
   Conn = Nothing

  Catch ex As Exception
        MsgBox(ex.Message)
  End Try

End Sub

Private Function RetrnAmount() As String

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlSelc As String

        With Conn

                If .State Then .Close()

                   .ConnectionString = strConnAdodb
                   .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                   .ConnectionTimeout = 90
                   .Open()

         End With

        strSqlSelc = "SELECT SUM(price)  AS sm_amt " _
                            & " FROM tmp_eqptrn " _
                            & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                            & " GROUP BY user_id"

         With Rsd

                 .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                 .LockType = ADODB.LockTypeEnum.adLockOptimistic
                 .Open(strSqlSelc, Conn, , , )

                  If .RecordCount <> 0 Then
                       RetrnAmount = .Fields("sm_amt").Value.ToString.Trim
                 Else
                       RetrnAmount = "0"
                 End If

                .ActiveConnection = Nothing
                .Close()

          End With

          Rsd = Nothing

    Conn.Close()
    Conn = Nothing

End Function

Private Function CheckCodeDuplicate() As Boolean

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset

Dim strSqlSelc As String

            With Conn

                    If .State Then .Close()

                            .ConnectionString = strConnAdodb
                            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                            .ConnectionTimeout = 90
                            .Open()

              End With

              strSqlSelc = "SELECT eqp_id " _
                                   & " FROM eqpmst " _
                                   & " WHERE eqp_id ='" & txtEqpId.Text.ToString.Trim & "'"

              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlSelc, Conn, , , )

                          If .RecordCount <> 0 Then
                                    CheckCodeDuplicate = False
                          Else
                                    CheckCodeDuplicate = True
                          End If

                         .ActiveConnection = Nothing
                         .Close()

             End With

             Rsd = Nothing

Conn.Close()
Conn = Nothing

End Function

Private Sub btnEditEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp1.Click

  Dim OpenFileDialog1 As New OpenFileDialog
  Dim strFileFullPath As String
  Dim strFileName As String
  Dim img As Image = Nothing

  Dim dateNow As Date = Now
  Dim typePic As String
  Dim strNamePic As String
  Dim lengPic, lengTypePic As Integer

      With OpenFileDialog1
           .CheckFileExists = True
           .ShowReadOnly = False
           .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png"
           .FilterIndex = 2

  Try

     If .ShowDialog = Windows.Forms.DialogResult.OK Then
         ClearBlankPicture1() 'เคลียร์รูปภาพ box
         strFileName = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์
         strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์

           img = ScaleImage(Image.FromFile(.FileName), picEqp1.Height, picEqp1.Width)  'ปรับขนาดรูปให้พอดีกับ picturebox ก่อน กรณีภาพใหญ่
           picEqp1.Image = img

           '----------- เพิ่มวันที่เวลาต่อชื่อไฟล์ ------------
           strFileName = Trim(strFileName)
           lengTypePic = strFileName.Length - 4
           typePic = Mid(strFileName, lengTypePic + 1, 4) ' ตัดเอา .jpg .png .gif 
           lengPic = strFileName.Length - 4   'เอาจำนวน charactor ทั้งหมดมาลบออกด้วยสกุล picture
           strNamePic = Mid(strFileName, 1, lengPic)     'ตัดเอาเฉพาะชื่อรูป
           strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           'เพิ่มวันที่ต่อท้ายชื่อไฟล์รูป

           lblPicPath1.Text = strFileFullPath
           lblPicName1.Text = strFileName

   End If

   Catch ex As Exception
         ClearBlankPicture1()
   End Try

End With

End Sub

Private Sub ClearBlankPicture1()
  picEqp1.Image = Nothing
  lblPicPath1.Text = ""
  lblPicName1.Text = ""
End Sub

Private Sub ClearBlankPicture2()
  picEqp2.Image = Nothing
  lblPicPath2.Text = ""
  lblPicName2.Text = ""

End Sub

Private Sub ClearBlankPicture3()
  picEqp3.Image = Nothing
  lblPicPath3.Text = ""
  lblPicName3.Text = ""

End Sub

Private Sub btnDelEqp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp1.Click
  ClearBlankPicture1()
End Sub

Private Sub btnDelEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp2.Click
    ClearBlankPicture2()
End Sub

Private Sub btnDelEqp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelEqp3.Click
    ClearBlankPicture3()
End Sub

Private Sub btnEditEqp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp2.Click

Dim OpenFileDialog2 As New OpenFileDialog
Dim strFileFullPath As String
Dim strFileName As String
Dim img As Image = Nothing

Dim dateNow As Date = Now
Dim typePic As String
Dim strNamePic As String
Dim lengPic, lengTypePic As Integer


With OpenFileDialog2
     .CheckFileExists = True
     .ShowReadOnly = False
     .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png"
     .FilterIndex = 2

     Try

            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                ' Load the specified file into a PictureBox control.
                ClearBlankPicture2()  'ล้าง picture box 2
                strFileName = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์
                strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์

                   img = ScaleImage(Image.FromFile(.FileName), picEqp2.Height, picEqp2.Width)
                   picEqp2.Image = img

                   '----------- เพิ่มวันที่เวลาต่อชื่อไฟล์ ------------
                   strFileName = Trim(strFileName)
                   lengTypePic = strFileName.Length - 4
                   typePic = Mid(strFileName, lengTypePic + 1, 4) ' ตัดเอา .jpg .png .gif 
                   lengPic = strFileName.Length - 4   'เอาจำนวน charactor ทั้งหมดมาลบออกด้วยสกุล picture
                   strNamePic = Mid(strFileName, 1, lengPic)     'ตัดเอาเฉพาะชื่อรูป
                   strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           'เพิ่มวันที่ต่อท้ายชื่อไฟล์รูป

                   lblPicPath2.Text = strFileFullPath
                   lblPicName2.Text = strFileName

            End If

        Catch ex As Exception
            ClearBlankPicture2()
        End Try

End With

End Sub

Private Sub btnEditEqp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditEqp3.Click

Dim OpenFileDialog3 As New OpenFileDialog
Dim strFileFullPath As String
Dim strFileName As String
Dim img As Image = Nothing

Dim dateNow As Date = Now
Dim typePic As String
Dim strNamePic As String
Dim lengPic, lengTypePic As Integer

With OpenFileDialog3
     .CheckFileExists = True
     .ShowReadOnly = False
     .Filter = "All Files|*.*|ไฟล์รูปภาพ (*)|*.bmp;*.gif;*.jpg;*.png"
     .FilterIndex = 2

    Try

       If .ShowDialog = Windows.Forms.DialogResult.OK Then
          ClearBlankPicture3()   'ล้าง picture box
          strFileName = New System.IO.FileInfo(.FileName).Name 'รับค่าเฉพาะชื่อไฟล์
          strFileFullPath = System.IO.Path.GetDirectoryName(.FileName) 'รับค่าเฉพาะพาธไฟล์

             img = ScaleImage(Image.FromFile(.FileName), picEqp3.Height, picEqp3.Width)
             picEqp3.Image = img

             '----------- เพิ่มวันที่เวลาต่อชื่อไฟล์ ------------
             strFileName = Trim(strFileName)
             lengTypePic = strFileName.Length - 4
             typePic = Mid(strFileName, lengTypePic + 1, 4) ' ตัดเอา .jpg .png .gif 
             lengPic = strFileName.Length - 4   'เอาจำนวน charactor ทั้งหมดมาลบออกด้วยสกุล picture
             strNamePic = Mid(strFileName, 1, lengPic)     'ตัดเอาเฉพาะชื่อรูป
             strFileName = strNamePic & "_" & DateTimeCutString(dateNow.ToString("yyyy/MM/dd hh:mm:ss")) & typePic           'เพิ่มวันที่ต่อท้ายชื่อไฟล์รูป

             lblPicPath3.Text = strFileFullPath
             lblPicName3.Text = strFileName

      End If

        Catch ex As Exception
            ClearBlankPicture3()
        End Try

End With

End Sub

Private Function chkPicName(ByVal fnames As String) As Boolean

 Dim di As New DirectoryInfo("\\10.32.0.15\data1\EquipPicture\")
 Dim aryFi As FileInfo() = di.GetFiles(fnames)
 Dim fi As FileInfo

    For Each fi In aryFi
        If fnames = fi.Name Then
           Exit Function
           Return False
        End If
    Next

    Return True

End Function

Private Sub picEqp1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp1.Click

  Dim strFilePicture As String = ""

      If Not picEqp1.Image Is Nothing Then
         strFilePicture = lblPicPath1.Text.ToString.Trim & "\" & lblPicName1.Text.ToString.Trim

         Try

            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(strFilePicture)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

         Catch ex As Exception
               MessageBox.Show("File " & strFilePicture & " Could not be found!", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
         End Try

     End If

End Sub

Private Sub picEqp1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp1.MouseHover

    With tt
         .Show("คลิกเพื่อดูรูปใหญ่", picEqp1)
         .AutomaticDelay = 500
         .AutoPopDelay = 5000
         .InitialDelay = 100
    End With

End Sub

Private Sub picEqp1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp1.MouseLeave
  tt.Hide(picEqp1)
End Sub

Private Sub picEqp2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp2.Click

  Dim strFilePicture As String = ""

      If Not picEqp2.Image Is Nothing Then
         strFilePicture = lblPicPath2.Text.ToString.Trim & "\" & lblPicName2.Text.ToString.Trim

         Try

            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(strFilePicture)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

         Catch ex As Exception
               MessageBox.Show("File " & strFilePicture & " Could not be found!", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
         End Try

     End If

End Sub

Private Sub picEqp2_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp2.MouseHover
   With tt
        .Show("คลิกเพื่อดูรูปใหญ่", picEqp2)
        .AutomaticDelay = 500
        .AutoPopDelay = 5000
        .InitialDelay = 100
   End With
End Sub

Private Sub picEqp2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp2.MouseLeave
  tt.Hide(picEqp2)
End Sub

Private Sub picEqp3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp3.Click

  Dim strFilePicture As String = ""

      If Not picEqp3.Image Is Nothing Then
         strFilePicture = lblPicPath3.Text.ToString.Trim & "\" & lblPicName3.Text.ToString.Trim

         Try

            Dim p As New System.Diagnostics.Process
            Dim s As New System.Diagnostics.ProcessStartInfo(strFilePicture)
            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

         Catch ex As Exception
               MessageBox.Show("File " & strFilePicture & " Could not be found!", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
         End Try

     End If

End Sub

Private Sub picEqp3_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp3.MouseHover
    With tt
         .Show("คลิกเพื่อดูรูปใหญ่", picEqp3)
         .AutomaticDelay = 500
         .AutoPopDelay = 5000
         .InitialDelay = 100
    End With
End Sub

Private Sub picEqp3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles picEqp3.MouseLeave
    tt.Hide(picEqp3)
End Sub

Private Function CallCopyPicture(ByVal strPicPath As String, ByVal strPicName As String, ByVal newPicname As String) As Boolean

Dim fName As String = String.Empty
Dim dFile As String = String.Empty
Dim dFilePath As String = String.Empty

Dim fServer As String = String.Empty
Dim intResult As Integer

On Error GoTo Err70

  fName = strPicPath & "\" & strPicName
  fServer = PthName & "\" & newPicname  'fServer = PthName & "\" & strPicName

If File.Exists(fServer) Then
   CallCopyPicture = True
Else

        If File.Exists(fName) Then

           dFile = Path.GetFileName(fName)
           dFilePath = DrvName + dFile

           intResult = String.Compare(fName.ToString.Trim, dFilePath.ToString.Trim)

           '------------------------------------ถ้าค่าเป็น 0 แสดงว่าโหลดไฟล์ใช้อยู่ ไม่สามารถ Copy ไฟล์ได้------------------------------

           If intResult = 1 Then 'ค่าที่ได้  1 ถึง copy รูปมาไว้ที่เครื่อง 10.32.0.15
              File.Copy(fName, dFilePath, True)
           End If

           My.Computer.FileSystem.RenameFile(dFilePath, newPicname)  'เปลี่ยนชื่อไฟล์รูปใหม่
           CallCopyPicture = True

        Else
               CallCopyPicture = True
        End If

End If

Err70:

If Err.Number <> 0 Then

    MsgBox("UserName ของคุณไม่มีสิทธิแก้ไขรูปภาพได้!!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Permission Can't Edit Picture")
    CallCopyPicture = False

End If

End Function

Private Sub SaveEditRecord()

Dim Conn As New ADODB.Connection
Dim strSqlCmd As String

Dim datSave As Date = Now()
Dim strDate As String = ""
Dim strDateDoc As String
Dim strGpType As String = ""
Dim strPartType As String = ""
Dim blnRetureCopyPic As Boolean

        With Conn
             If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 150
                .Open()
        End With

                    Conn.BeginTrans()

                    strDate = datSave.Date.ToString("yyyy-MM-dd")
                    strDate = SaveChangeEngYear(strDate)

                    strDateDoc = Mid(txtBegin.Text.ToString, 7, 4) & "-" _
                                            & Mid(txtBegin.Text.ToString, 4, 2) & "-" _
                                            & Mid(txtBegin.Text.ToString, 1, 2)
                    strDateDoc = SaveChangeEngYear(strDateDoc)

                    '------------------------------------กำหนดกลุ่มของชิ้นงาน---------------------------------------------------------------------

                    Select Case cmbPart.Text.ToString.Trim

                           Case Is = "หนังหน้า"
                                     strPartType = "UPPER"
                           Case Is = "อุปกรณ์"
                                     strPartType = "ACCSSY"
                           Case Is = "พื้นบน"
                                     strPartType = "SOLE1"
                           Case Is = "พื้นล่าง"
                                     strPartType = "SOLE2"
                           Case Is = "รองเท้าสำเร็จรูป"
                                     strPartType = "SOLE3"
                           Case Is = "การ์ด"
                                     strPartType = "GUARD"
                           Case Is = "กระดุม"
                                    strPartType = "BUTTON"
                           Case Is = "จมูก"
                                    strPartType = "NOSE"
                           Case Is = "ขา"
                                    strPartType = "LEG"
                           Case Is = "TPR โลโก้"
                                    strPartType = "TPR"

                    End Select

                    '-------------------------------------------กำหนดกลุ่มของอุปกรณ์---------------------------------------------

                    Select Case cmbGp.Text.ToString.Trim

                           Case Is = "MOLD EVA"
                                strGpType = "A"

                           Case Is = "MOLD PVC"
                                strGpType = "B"

                           Case Is = "MOLD PU"
                                strGpType = "C"

                    End Select


                     '------------------------------------------------บันทึกรูปภาพการบรรจุ---------------------------------------

                    blnRetureCopyPic = CallCopyPicture(lblPicPath1.Text.ToString.Trim, ReturnImageName(lblPicName1.Text.ToString.Trim), lblPicName1.Text.ToString.Trim)
                    If blnRetureCopyPic Then
                       lblPicPath1.Text = PthName

                    Else
                          lblPicName1.Text = ""
                          lblPicPath1.Text = ""
                          picEqp1.Image = Nothing

                    End If

                    '------------------------------------------- บันทึกรูปภาพการภายนอก ภายใน -------------------------------------------

                    blnRetureCopyPic = CallCopyPicture(lblPicPath2.Text.ToString.Trim, ReturnImageName(lblPicName2.Text.ToString.Trim), lblPicName2.Text.ToString.Trim)
                    If blnRetureCopyPic Then
                       lblPicPath2.Text = PthName

                    Else
                          lblPicName2.Text = ""
                          lblPicPath2.Text = ""
                          picEqp2.Image = Nothing

                    End If

                    '----------------------------------------- บันทึกรูปภาพการชิ้นงาน ---------------------------------------------------

                    blnRetureCopyPic = CallCopyPicture(lblPicPath3.Text.ToString.Trim, ReturnImageName(lblPicName3.Text.ToString.Trim), lblPicName3.Text.ToString.Trim)
                    If blnRetureCopyPic Then
                       lblPicPath3.Text = PthName

                    Else
                          lblPicName3.Text = ""
                          lblPicPath3.Text = ""
                          picEqp3.Image = Nothing

                    End If

                      strSqlCmd = "UPDATE  eqpmst SET eqp_name ='" & ReplaceQuote(txtEqpName.Text.ToString.Trim) & "'" _
                                            & "," & "pi ='" & ReplaceQuote(txtOrd.Text.ToString.Trim) & "'" _
                                            & "," & "shoe ='" & ReplaceQuote(txtStyle.Text.ToString.Trim) & "'" _
                                            & "," & "part ='" & strPartType & "'" _
                                            & "," & "eqp_type ='" & "" & "'" _
                                            & "," & "set_qty =" & ChangFormat(txtSet.Text.ToString.Trim) _
                                            & "," & "pic_ctain ='" & ReplaceQuote(lblPicName1.Text.ToString.Trim) & "'" _
                                            & "," & "pic_io ='" & ReplaceQuote(lblPicName2.Text.ToString.Trim) & "'" _
                                            & "," & "pic_part ='" & ReplaceQuote(lblPicName3.Text.ToString.Trim) & "'" _
                                            & "," & "remark ='" & ReplaceQuote(txtRemark.Text.ToString.Trim) & "'" _
                                            & "," & "creat_date ='" & strDateDoc & "'" _
                                            & "," & "eqp_amt =" & RetrnAmount() _
                                            & "," & "last_date ='" & strDate & "'" _
                                            & "," & "last_by ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'" _
                                            & "," & "exp_id ='" & "" & "'" _
                                            & " WHERE eqp_id ='" & txtEqpId.Text.ToString.Trim & "'"

                      Conn.Execute(strSqlCmd)


                    '------------------------------------------------  ลบข้อมูลในตาราง eqptrn  -----------------------------------------------------

                     strSqlCmd = "Delete FROM eqptrn" _
                                          & " WHERE eqp_id ='" & txtEqpId.Text.ToString.Trim & "'"

                     Conn.Execute(strSqlCmd)

                     '------------------------------------------------ บันทึกข้อมูลในตาราง eqptrn  ---------------------------------------------------

        strSqlCmd = "INSERT INTO eqptrn " _
                            & " SELECT [group] ='" & strGpType.ToString.Trim & "'" _
                            & ",eqp_id ='" & txtEqpId.Text.ToString.Trim & "'" _
                            & ",size_id,size_desc,size_qty,weight,dimns,backgup,price,men_rmk" _
                            & ",delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty" _
                            & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,mouth_long" _
                            & "  FROM tmp_eqptrn" _
                            & " WHERE user_id ='" & frmMainPro.lblLogin.Text.Trim.ToString & "'"

                      Conn.Execute(strSqlCmd)
                      Conn.CommitTrans()

                      frmMoldInj.lblCmd.Text = txtEqpId.Text.ToString.Trim       'บ่งบอกว่าบันทึกข้อมูลสำเร็จ
                      frmMoldInj.Activating()
                      Me.Close()

   Conn.Close()
   Conn = Nothing

End Sub


Private Sub txtSet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSet.KeyDown
Dim intChkPoint As Integer

    With txtSet

            Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                           If .SelectionStart = 0 Then
                           End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น 
                           cmbType.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtRemark.Focus()
                                Else

                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtRemark.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง                                                            
                      Case Is = 113 'ปุ่ม F2
                             .SelectionStart = .Text.Trim.Length

            End Select

    End With

End Sub

Private Sub txtSet_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSet.KeyPress
  If e.KeyChar = Chr(13) Then
     txtRemark.Focus()
  End If

End Sub

Private Sub txtSet_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSet.LostFocus
    With txtSet
            .Text = .Text.ToString.Trim.ToUpper
    End With

End Sub

Private Sub txtExpId_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExpId.GotFocus
  txtExpId.Focus()
End Sub


Private Sub txtExpId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExpId.KeyDown
    Dim intChkPoint As Integer

    With txtExpId

            Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                  
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                     txtEqpId.Focus()
                                Else

                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                            txtEqpId.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง
                                    txtEqpName.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select

    End With

End Sub

Private Sub txtExpId_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtExpId.KeyPress

    Select Case e.KeyChar

                 Case "0" To "9"
                            e.Handled = False

                  Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" _
                            , "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
                            e.Handled = False

                Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n" _
                            , "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
                            e.Handled = False


                 Case "-"
                            e.Handled = False

                 Case Is = Chr(13)
                            e.Handled = False
                             txtEqpId.Focus()

                  Case Chr(8), Chr(46)
                            e.Handled = False

                  Case Else
                            e.Handled = True

        End Select

End Sub

Private Sub txtExpId_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExpId.LostFocus

    With txtExpId
            .Text = .Text.ToString.Trim.ToUpper
    End With

End Sub

Private Sub txtSetQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSetQty.GotFocus

    With mskSetQty
             txtSetQty.SendToBack()
            .BringToFront()
            .Focus()

    End With

End Sub

    Private Sub mskSetQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSetQty.GotFocus

        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerge As String = ""

        With mskSetQty

            If txtSetQty.Text.ToString.Trim <> "" Then

                x = Len(txtSetQty.Text.ToString)

                For i = 1 To x

                    strTmp = Mid(txtSetQty.Text.ToString, i, 1)
                    Select Case strTmp

                        Case Is = "_"
                        Case Else

                            If InStr("0123456789.", strTmp) > 0 Then
                                strMerge = strMerge & strTmp
                            End If

                    End Select

                Next i

                Select Case strMerge.IndexOf(".")

                    Case Is = -1
                        .SelectionStart = 0
                    Case Is = 1
                        .SelectionStart = 1
                    Case Is = 2
                        .SelectionStart = 0
                    Case Is = 3
                        .SelectionStart = 0
                    Case Else
                        .SelectionStart = 0
                End Select

                .SelectedText = strMerge

            End If

            .SelectAll()

        End With

    End Sub

    Private Sub mskSetQty_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskSetQty.KeyDown

        Dim intChkPoint As Integer

        With mskSetQty

            Select Case e.KeyCode
                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย

                        If .SelectionStart = 0 Then
                           txtSizeDesc.Focus()
                        End If
                   Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                   Case Is = 39 'ปุ่มลูกศรขวา

                        If .SelectionLength = .Text.Trim.Length Then
                           txtSizeQty.Focus()
                        Else
                           intChkPoint = .Text.Trim.Length
                           If .SelectionStart = intChkPoint Then
                              txtSizeQty.Focus()
                           End If
                        End If

                      Case Is = 40 'ปุ่มลง    
                                  txtRecvDate.Focus()
                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length
            End Select

    End With

End Sub

Private Sub mskSetQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskSetQty.KeyPress

        If e.KeyChar = Chr(13) Then
            txtSizeQty.Focus()
        End If

    End Sub

Private Sub mskSetQty_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskSetQty.LostFocus

  Dim i, x As Integer
  Dim z As Double

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskSetQty

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp

                                       Case Is = ","
                                       Case Is = "+"
                                       Case Is = "_"
                                       Case Else

                                            If InStr("0123456789.", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If

                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskSetQty.Text = ""
                    z = CDbl(strMerge)
                    txtSetQty.Text = z.ToString("#,##0.0")

                Catch ex As Exception
                    txtSetQty.Text = "0.0"
                    mskSetQty.Text = ""
               End Try

mskSetQty.SendToBack()
txtSetQty.BringToFront()

End With

End Sub

Private Sub txtPrDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrDate.GotFocus

  With mskPrDate
       txtPrDate.SendToBack()
       .BringToFront()
       .Focus()
 End With

End Sub

    Private Sub mskPrDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrDate.GotFocus

        Dim i, x As Integer
        Dim strTmp As String = ""
        Dim strMerge As String = ""

        With mskPrDate

            If txtPrDate.Text <> "__/__/____" Then

                x = Len(txtPrDate.Text.ToString) 'นับจำนวนตัวอักษร

                For i = 1 To x

                    strTmp = Mid(txtPrDate.Text.ToString, i, 1)
                    Select Case strTmp
                        Case Is = "_"
                        Case Else

                            If InStr("0123456789/", strTmp) > 0 Then   'ค้นหาสตริง
                                strMerge = strMerge & strTmp
                            End If

                    End Select

                Next i

                Select Case strMerge.ToString.Length

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

                .SelectedText = strMerge

            End If

            .SelectAll()

        End With
    End Sub

    Private Sub mskPrDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskPrDate.KeyDown
Dim intChkPoint As Integer

    With mskPrDate

            Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                        txtWeight.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น
                                  txtSize.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtFcDate.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtFcDate.Focus()
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง
                              txtPrDoc.Focus()
                      Case Is = 113 'ปุ่ม F2
                               .SelectionStart = .Text.Trim.Length
            End Select

    End With

End Sub

Private Sub mskPrDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskPrDate.KeyPress
   If e.KeyChar = Chr(13) Then
      txtFcDate.Focus()
   End If
End Sub

Private Sub mskPrDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskPrDate.LostFocus

  Dim i, x As Integer
  Dim z As Date

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskPrDate
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
                    mskPrDate.Text = ""
                    strMerge = "#" & strMerge & "#"
                    z = CDate(strMerge)
                    txtPrDate.Text = z.ToString("dd/MM/yyyy")
               Catch ex As Exception
                    txtPrDate.Text = "__/__/____"
                    mskPrDate.Text = ""

               End Try

    mskPrDate.SendToBack()
    txtPrDate.BringToFront()

End With

End Sub

Private Sub txtRecvDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRecvDate.GotFocus

    With mskRecvDate
         txtRecvDate.SendToBack()
         .BringToFront()
         .Focus()
    End With

End Sub

Private Sub mskRecvDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvDate.GotFocus

  Dim i, x As Integer
  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskRecvDate

           If txtRecvDate.Text <> "__/__/____" Then

                        x = Len(txtRecvDate.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtRecvDate.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = "_"
                                       Case Else

                                            If InStr("0123456789/", strTmp) > 0 Then
                                                strMerge = strMerge & strTmp
                                            End If

                                End Select

                         Next i

                        Select Case strMerge.ToString.Length

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

                        .SelectedText = strMerge
                End If

        .SelectAll()
      End With
End Sub

Private Sub mskRecvDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskRecvDate.KeyDown

   Dim intChkPoint As Integer

       With mskRecvDate

            Select Case e.KeyCode
                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย

                        If .SelectionStart = 0 Then
                          cmbType.Focus()
                        End If

                   Case Is = 38 'ปุ่มลูกศรขึ้น
                        txtImpt.Focus()
                   Case Is = 39 'ปุ่มลูกศรขวา

                        If .SelectionLength = .Text.Trim.Length Then
                           txtInvoice.Focus()
                        Else
                             intChkPoint = .Text.Trim.Length
                             If .SelectionStart = intChkPoint Then
                                txtInvoice.Focus()
                             End If
                        End If

                  Case Is = 40 'ปุ่มลง                              
                  Case Is = 113 'ปุ่ม F2
                          .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub mskRecvDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskRecvDate.KeyPress

  If e.KeyChar = Chr(13) Then
     txtInvoice.Focus()
  End If

End Sub

Private Sub mskRecvDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskRecvDate.LostFocus

  Dim i, x As Integer
  Dim z As Date

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskRecvDate

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

                    mskRecvDate.Text = ""
                    strMerge = "#" & strMerge & "#"
                    z = CDate(strMerge)
                    txtRecvDate.Text = z.ToString("dd/MM/yyyy")


            Catch ex As Exception
                    txtRecvDate.Text = "__/__/____"
                    mskRecvDate.Text = ""

               End Try

        mskRecvDate.SendToBack()
        txtRecvDate.BringToFront()
   End With

End Sub

Private Sub txtFcDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFcDate.GotFocus

    With mskFcDate
         txtFcDate.SendToBack()
         .BringToFront()
         .Focus()
    End With

End Sub

Private Sub mskFcDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFcDate.GotFocus

  Dim i, x As Integer
  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskFcDate

        If txtFcDate.Text <> "__/__/____" Then

                        x = Len(txtFcDate.Text.ToString)

                        For i = 1 To x

                                strTmp = Mid(txtFcDate.Text.ToString, i, 1)
                                Select Case strTmp
                                       Case Is = "_"
                                       Case Else

                                            If InStr("0123456789/", strTmp) > 0 Then
                                               strMerge = strMerge & strTmp
                                            End If

                                End Select

                         Next i

                        Select Case strMerge.ToString.Length  'หาตำแหน่งที่พบเป็นครั้งแรก

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

                        .SelectedText = strMerge

                End If

        .SelectAll()

    End With

End Sub

Private Sub mskFcDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskFcDate.KeyDown
Dim intChkPoint As Integer

    With mskFcDate
            Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                   txtPrDate.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น
                                  txtSizeDesc.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                        txtRecvDate.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                        txtRecvDate.Focus()
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง           
                                txtPrice.Focus()
                      Case Is = 113 'ปุ่ม F2
                               .SelectionStart = .Text.Trim.Length
            End Select
    End With
End Sub

Private Sub mskFcDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskFcDate.KeyPress
   If e.KeyChar = Chr(13) Then
      txtPrDoc.Focus()
   End If
End Sub

Private Sub mskFcDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskFcDate.LostFocus

   Dim i, x As Integer
   Dim z As Date

   Dim strTmp As String = ""
   Dim strMerge As String = ""

       With mskFcDate

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

                    mskFcDate.Text = ""
                    strMerge = "#" & strMerge & "#"
                    z = CDate(strMerge)
                    txtFcDate.Text = z.ToString("dd/MM/yyyy")

               Catch ex As Exception
                    txtFcDate.Text = "__/__/____"
                    mskFcDate.Text = ""
               End Try

         mskFcDate.SendToBack()
         txtFcDate.BringToFront()

   End With

End Sub

Private Sub txtInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.GotFocus
  txtInvoice.SelectAll()
End Sub

Private Sub txtInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoice.KeyDown
  Dim intChkPoint As Integer
      With txtInvoice

            Select Case e.KeyCode

                   Case Is = 35 'ปุ่ม End 
                   Case Is = 36 'ปุ่ม Home
                   Case Is = 37 'ลูกศรซ้าย

                        If .SelectionStart = 0 Then
                           txtRecvDate.Focus()
                        End If

                  Case Is = 38 'ปุ่มลูกศรขึ้น    
                                  txtSuplier.Focus()
                  Case Is = 39 'ปุ่มลูกศรขวา

                          If .SelectionLength = .Text.Trim.Length Then
                             txtMouth_mold.Focus()
                          Else
                                intChkPoint = .Text.Trim.Length
                                If .SelectionStart = intChkPoint Then
                                   txtMouth_mold.Focus()
                                   End If
                                End If

                 Case Is = 40 'ปุ่มลง                                      
                 Case Is = 113 'ปุ่ม F2
                         .SelectionStart = .Text.Trim.Length

            End Select

    End With

End Sub

Private Sub txtInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice.KeyPress

Select Case Asc(e.KeyChar)

       Case 45 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
             e.Handled = False

       Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

       Case 13   'Enter = 13
              e.Handled = False
              txtMouth_mold.Focus()

      Case Else
             e.Handled = True
             MessageBox.Show("กรุณาระบุข้อมูลเป็นภาษาอังกฤษ")
  End Select

End Sub

Private Sub txtInvoice_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.LostFocus
  With txtInvoice
       .Text = .Text.ToString.Trim.ToUpper
  End With

End Sub

Private Sub txtImpt_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImpt.GotFocus
  txtImpt.SelectAll()
End Sub

Private Sub txtImpt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtImpt.KeyDown
    Dim intChkPoint As Integer

    With txtImpt

               Select Case e.KeyCode

                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                  txtPrice.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น   
                                  txtPrDoc.Focus()
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtSuplier.Focus()
                                Else

                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtSuplier.Focus()
                                    End If

                                End If

                      Case Is = 40 'ปุ่มลง
                               txtRecvDate.Focus()
                      Case Is = 113 'ปุ่ม F2
                               .SelectionStart = .Text.Trim.Length
               End Select

    End With

End Sub

Private Sub txtImpt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtImpt.KeyPress

Select Case Asc(e.KeyChar)

       Case 45 To 122 ' โค๊ดภาษาอังกฤษ์ตามจริงจะอยู่ที่ 58ถึง122 แต่ที่เอา 48มาเพราะเราต้องการตัวเลข
             e.Handled = False

       Case 8, 46 ' Backspace = 8,  Delete = 46
             e.Handled = False

       Case 13   'Enter = 13
              e.Handled = False
              txtSuplier.Focus()

      Case Else
             e.Handled = True
             MessageBox.Show("กรุณาระบุข้อมูลเป็นภาษาอังกฤษ")

  End Select

End Sub

Private Sub txtImpt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtImpt.LostFocus
    txtImpt.Text = txtImpt.Text.ToUpper
End Sub

Private Sub ClearAllData()

   lblGpName.Text = ""
   txtEqpId.Text = ""
   txtEqpName.Text = ""
   txtStyle.Text = ""
   txtOrd.Text = ""
   lblOrdQty.Text = "0"
   txtSet.Text = "0"
   lblAmt.Text = "0.00"
   txtRemark.Text = ""

   '------- Cleare picture box -------
   dgvSize.Rows.Clear()
   picEqp1.Image = Nothing
   picEqp2.Image = Nothing
   picEqp3.Image = Nothing

   '--------- Clear หน้าต่างเพิ่มข้อมูล --------

   txtSize.Text = ""
   txtSizeDesc.Text = ""
   txtSetQty.Text = "0"
   txtSizeQty.Text = "0"
   txtWeight.Text = "0.00"
   txtPrDate.Text = "__/__/____"
   txtFcDate.Text = "__/__/____"
   txtPrDoc.Text = ""
   txtPrice.Text = "0.00"
   txtWd.Text = "0.00"
   txtLg.Text = "0.00"
   txtHg.Text = "0.00"
   txtImpt.Text = ""
   txtSuplier.Text = ""
   txtRecvDate.Text = "__/__/____"
   txtInvoice.Text = "'"
   txtRmk.Text = ""
End Sub

Private Sub ReSizeSort(ByVal Eqpid As String)

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim sql As String
  Dim sqlCmd As String

  Dim strArr() As String
  Dim SearchWithinThis As String
  Dim newSize As String

  Dim prDate As String
  Dim RecvDate As String
  Dim FcDate As String

      Try

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
      End With

      sql = "SELECT * FROM tmp_eqptrn " _
                & " WHERE user_id= '" & frmMainPro.lblLogin.Text.ToString.Trim & "'" _
                & " ORDER BY size_id"

      With Rsd
           .LockType = ADODB.LockTypeEnum.adLockOptimistic
           .CursorType = ADODB.CursorLocationEnum.adUseClient
           .Open(sql, Conn, , , )

           If .RecordCount <> 0 Then

               '----------------------- ล้างข้อมูลใน tmp_eqptrn_newsize ------------------------------

                 sqlCmd = "DELETE FROM tmp_eqptrn_newsize " _
                              & "WHERE user_id= '" & frmMainPro.lblLogin.Text.ToString.Trim & "'"

                 Conn.Execute(sqlCmd)

               ' ---------- วนลูปจัดเรียง size ใหม่ --------------

               Do While Not .EOF

                  SearchWithinThis = .Fields("size_id").Value.ToString.Trim
                  If SearchWithinThis.IndexOf("-") <> -1 Then          'หาก size ต้นฉบับไม่มี size รว่ม (x-xx)
                     strArr = SearchWithinThis.Split("-")              'อ่านค่า size เก็บไว้ในตัวเเปร
                     newSize = strArr(0)
                  Else
                       newSize = .Fields("size_id").Value.ToString.Trim
                  End If

                  If .Fields("pr_date").Value.ToString.Trim <> "" Then
                      prDate = Mid(.Fields("pr_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("pr_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("pr_date").Value.ToString.Trim, 1, 2)
                      prDate = "'" & SaveChangeEngYear(prDate) & "'"

                  Else
                      prDate = "NULL"
                  End If

                  If .Fields("recv_date").Value.ToString.Trim <> "" Then
                      RecvDate = Mid(.Fields("recv_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("recv_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("recv_date").Value.ToString.Trim, 1, 2)
                      RecvDate = "'" & SaveChangeEngYear(RecvDate) & "'"
                  Else
                       RecvDate = "NULL"
                  End If

                  If .Fields("fc_date").Value.ToString.Trim <> "" Then
                     FcDate = Mid(.Fields("fc_date").Value.ToString.Trim, 7, 4) & "-" _
                                                  & Mid(.Fields("fc_date").Value.ToString.Trim, 4, 2) & "-" _
                                                  & Mid(.Fields("fc_date").Value.ToString.Trim, 1, 2)
                      FcDate = "'" & SaveChangeEngYear(FcDate) & "'"
                  Else
                       FcDate = "NULL"
                  End If

                  '----------------------- Insert ข้อมูลลงในตารางใหม่หลังเรียง size ใหม่ ----------------------

                   sqlCmd = "INSERT INTO tmp_eqptrn_newsize " _
                           & "(user_id,[group],eqp_id,size_id,size_desc,size_qty,weight,dimns,backgup " _
                           & ",price,men_rmk,delvr_sta,sent_sta,set_qty,pr_date,pr_doc,recv_date,ord_rep,ord_qty " _
                           & ",fc_date,impt_id,sup_name,lp_type,size_group,cut_id,mate_type,cut_detail,tmp_newsize) " _
                           & "VALUES( " _
                           & "'" & frmMainPro.lblLogin.Text.Trim & "'" _
                           & ",'" & .Fields("group").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("eqp_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_desc").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("size_qty").Value.ToString.Trim) _
                           & "," & ChangFormat(.Fields("weight").Value.ToString.Trim) _
                           & ",'" & .Fields("dimns").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("backgup").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("price").Value.ToString.Trim) _
                           & ",'" & .Fields("men_rmk").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("delvr_sta").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("sent_sta").Value.ToString.Trim & "'" _
                           & "," & ChangFormat(.Fields("set_qty").Value.ToString.Trim) _
                           & "," & prDate _
                           & ",'" & .Fields("pr_doc").Value.ToString.Trim & "'" _
                           & "," & RecvDate _
                           & "," & ChangFormat(.Fields("ord_rep").Value.ToString.Trim) _
                           & "," & ChangFormat(.Fields("ord_qty").Value.ToString.Trim) _
                           & "," & FcDate _
                           & ",'" & .Fields("impt_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("sup_name").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("lp_type").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("size_group").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("cut_id").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("mate_type").Value.ToString.Trim & "'" _
                           & ",'" & .Fields("cut_detail").Value.ToString.Trim & "'" _
                           & "," & newSize _
                           & ")"

                        Conn.Execute(sqlCmd)


                  .MoveNext()
               Loop

           End If

          .ActiveConnection = Nothing
          .Close()
      End With

   Conn.Close()

      Catch ex As Exception
            MsgBox(ex.Message)
      End Try
End Sub

    Private Sub LoadHistory_Fixmold()

        Dim Conn As New ADODB.Connection
        Dim Rsd As New ADODB.Recordset
        Dim strSql As String
        Dim sqlSelc As String

        Dim i As Integer = 1
        Dim totalPrice As Double = 0.0
        Dim strCode As String = frmMoldInj.dgvShoe.Rows(frmMoldInj.dgvShoe.CurrentRow.Index).Cells(0).Value.ToString.Trim

        Try

            With Conn
                If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .Open()

                '---------------- เพิ่มข้อมูลในตาราง  tmp_fixeqptrn ---------------
                strSql = "INSERT INTO tmp_fixeqptrn " _
                          & " SELECT user_id= '" & frmMainPro.lblLogin.Text.Trim & "', *" _
                          & " FROM fixeqptrn " _
                          & " WHERE eqp_id='" & strCode & "'" _
                          & " AND fix_sta <> '1'"

                .Execute(strSql)

            End With

            sqlSelc = "SELECT * FROM tmp_fixeqptrn (NOLOCK) " _
                        & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                        & " AND eqp_id='" & strCode & "'" _
                        & " ORDER BY size_id"

            With Rsd

                .LockType = ADODB.LockTypeEnum.adLockOptimistic
                .CursorType = ADODB.CursorLocationEnum.adUseClient
                .Open(sqlSelc, Conn, , , )

                If .RecordCount <> 0 Then
                    lblMold.Visible = False
                    CallMoldSize()  'โหลดรายการ size ลง combo เลือก size
                    cmbFindSize.SelectedIndex = 0
                    dgvHistory_issue.Rows.Clear()
                    Do While Not .EOF
                        dgvHistory_issue.Rows.Add(
                                                 i,
                                                 .Fields("eqp_id").Value.ToString.Trim & " / " & .Fields("size_id").Value.ToString.Trim,
                                                 .Fields("issue").Value.ToString.Trim,
                                                 .Fields("sup_name").Value.ToString.Trim,
                                                 Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10),
                                                 Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10),
                                                 Format(.Fields("amt_out").Value, "##0.0"),
                                                 Format(.Fields("price").Value, "#,##0.00"),
                                                 .Fields("fix_by").Value.ToString.Trim,
                                                 .Fields("fix_rmk").Value.ToString.Trim
                                               )

                        i += 1
                        totalPrice = totalPrice + .Fields("price").Value

                        .MoveNext()
                    Loop
                    lblMoldPrice.Text = Format(totalPrice, "#,##0.00")

                    lblsize.Visible = True
                    cmbFindSize.Visible = True

                Else
                    lblMold.Visible = True
                    lblsize.Visible = False
                    cmbFindSize.Visible = False
                    lblMoldPrice.Text = "0.00"
                End If

                .ActiveConnection = Nothing
                .Close()
            End With

            Conn.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub CallMoldSize()

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSql As String

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

     strSql = "SELECT DISTINCT size_id " _
               & " FROM tmp_fixeqptrn " _
               & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'"

      With Rsd

              .LockType = ADODB.LockTypeEnum.adLockOptimistic
              .CursorType = ADODB.CursorLocationEnum.adUseClient
              .Open(strSql, Conn, , , )

              If .RecordCount <> 0 Then
                 cmbFindSize.Items.Add("--ทุก SIZE--")
                 Do While Not .EOF

                    cmbFindSize.Items.Add(.Fields("size_id").Value.ToString.Trim)

                   .MoveNext()
                 Loop

              End If

         .ActiveConnection = Nothing
         .Close()
     End With

   Conn.Close()

End Sub

Private Sub cmbFindSize_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFindSize.SelectedValueChanged
  If cmbFindSize.Text = "--ทุก SIZE--" Then
     FindFix_mold("ALL")
  Else
       FindFix_mold(cmbFindSize.Text)
  End If
End Sub

Private Sub FindFix_mold(ByVal searchTxt As String)

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSql As String

  Dim i As Integer = 1
  Dim totalPrice As Double = 0.0

      With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()
     End With

     If searchTxt = "ALL" Then

        '---------- กรณีเลือกทุก SIZE ------------

        strSql = "SELECT * FROM tmp_fixeqptrn " _
                  & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'"

     Else

          '------------- กรณ๊ระบุ SIZE -------------

          strSql = "SELECT * FROM tmp_fixeqptrn " _
                  & " WHERE user_id = '" & frmMainPro.lblLogin.Text.Trim & "'" _
                  & " AND size_id='" & searchTxt & "'"

     End If

      With Rsd

              .LockType = ADODB.LockTypeEnum.adLockOptimistic
              .CursorType = ADODB.CursorLocationEnum.adUseClient
              .Open(strSql, Conn, , , )

              If .RecordCount <> 0 Then

                 cmbFindSize.SelectedValue = searchTxt
                 dgvHistory_issue.Rows.Clear()
                 Do While Not .EOF
                       dgvHistory_issue.Rows.Add( _
                                                i, _
                                                .Fields("eqp_id").Value.ToString.Trim & " / " & .Fields("size_id").Value.ToString.Trim, _
                                                .Fields("issue").Value.ToString.Trim, _
                                                .Fields("sup_name").Value.ToString.Trim, _
                                                Mid(.Fields("fix_date").Value.ToString.Trim, 1, 10), _
                                                Mid(.Fields("recv_date").Value.ToString.Trim, 1, 10), _
                                                Format(.Fields("amt_out").Value, "##0.0"), _
                                                Format(.Fields("price").Value, "#,##0.00"), _
                                                .Fields("fix_by").Value.ToString.Trim, _
                                                .Fields("fix_rmk").Value.ToString.Trim _
                                              )

                        i += 1
                        totalPrice = totalPrice + .Fields("price").Value

                   .MoveNext()
                 Loop
                 lblMoldPrice.Text = Format(totalPrice, "#,##0.00")

                 lblsize.Visible = True
                 cmbFindSize.Visible = True

              End If

         .ActiveConnection = Nothing
         .Close()
     End With

   Conn.Close()

End Sub

Private Sub PreMoldStatus()
 Dim sta(4) As String
 Dim i As Byte

     sta(0) = "------ โปรดเลือกสถานะ ------"
     sta(1) = "ปกติ / มีการเคลื่อนไหว"
     sta(2) = "รอใช้งาน / ไม่มีการเคลื่อนไหว"
     sta(3) = "ยกเลิกการใช้งาน"

     With cmbStatus_mold

         For i = 0 To 3
             .Items.Add(sta(i))
         Next

          .SelectedIndex = 0
  End With

End Sub

Private Sub txtMouth_mold_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMouth_mold.GotFocus
  With mskMouth_mold
       .BringToFront()
       txtMouth_mold.SendToBack()
       .Focus()
  End With
End Sub

Private Sub mskMouth_mold_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskMouth_mold.GotFocus

  Dim i, x As Integer
  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskMouth_mold

           If txtMouth_mold.Text <> "0.00" Then

                x = Len(txtMouth_mold.Text.ToString)

                For i = 1 To x

                    strTmp = Mid(txtMouth_mold.Text.ToString, i, 1)
                    Select Case strTmp

                        Case Is = "_"
                        Case Else
                            If InStr(",.0123456789", strTmp) > 0 Then
                                strMerge = strMerge & strTmp
                            End If

                    End Select

                Next i

                Select Case strMerge.IndexOf(".")

                    Case Is = 5
                        .SelectionStart = 0
                    Case Is = 3
                        .SelectionStart = 2
                    Case Is = 2
                        .SelectionStart = 3
                    Case Is = 1
                        .SelectionStart = 4
                    Case Else
                        .SelectionStart = 0
                End Select

                .SelectedText = strMerge

                End If
          .SelectAll()
       End With
End Sub

Private Sub mskMouth_mold_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mskMouth_mold.KeyDown

 Dim intChkPoint As Integer

     With mskMouth_mold

               Select Case e.KeyCode
                      Case Is = 35 'ปุ่ม End 
                      Case Is = 36 'ปุ่ม Home
                      Case Is = 37 'ลูกศรซ้าย

                               If .SelectionStart = 0 Then
                                    txtInvoice.Focus()
                               End If

                      Case Is = 38 'ปุ่มลูกศรขึ้น                                      
                      Case Is = 39 'ปุ่มลูกศรขวา

                                If .SelectionLength = .Text.Trim.Length Then
                                   txtRmk.Focus()
                                Else
                                    intChkPoint = .Text.Trim.Length
                                    If .SelectionStart = intChkPoint Then
                                       txtRmk.Focus()
                                    End If
                                End If

                      Case Is = 40 'ปุ่มลง               
                                txtRemark.Focus()

                      Case Is = 113 'ปุ่ม F2
                                .SelectionStart = .Text.Trim.Length

            End Select

    End With
End Sub

Private Sub mskMouth_mold_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles mskMouth_mold.KeyPress
  Select Case Asc(e.KeyChar)

         Case 13
              txtRmk.Focus()
         Case 46 'เครื่องหมายจุลภาค(.)                                      
                 mskMouth_mold.SelectionStart = 6
 End Select
End Sub

Private Sub mskMouth_mold_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskMouth_mold.LostFocus
  Dim i, x As Integer
  Dim z As Double

  Dim strTmp As String = ""
  Dim strMerge As String = ""

      With mskMouth_mold

                x = .Text.Length

                        For i = 1 To x

                                strTmp = Mid(.Text.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = ","
                                          Case Is = "+"
                                          Case Is = "_"
                                          Case Else

                                               If InStr(".0123456789", strTmp) > 0 Then
                                                  strMerge = strMerge & strTmp
                                               End If

                                End Select
                                strTmp = ""
                        Next i
                Try

                    mskMouth_mold.Text = ""
                    z = CDbl(strMerge)
                    txtMouth_mold.Text = z.ToString("#,##0.00")


                Catch ex As Exception
                    txtMouth_mold.Text = "0.00"
                    mskMouth_mold.Text = ""
               End Try

mskMouth_mold.SendToBack()
txtMouth_mold.BringToFront()

End With
End Sub

End Class