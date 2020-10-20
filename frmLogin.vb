Imports System.Data
Imports System.Data.OleDb.OleDbDataAdapter
Imports ADODB
Imports Microsoft.VisualBasic
Imports System.Net

Public Class frmLogin

Private Sub FrmLogin_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
  ClearUserOutofSystem()
  InputUsrUsedProgram()

  Me.Dispose()
End Sub

Private Sub frmLogin_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Chr(27) Then    '���� ASC

            If FormCount("frmMainpro") > 0 Then '�ʴ���ҿ������ѡ�١�Դ����
                Me.Hide()
                frmMainPro.Enabled = True
                frmMainPro.Focus()
            Else
                Me.Close()
            End If

        End If
End Sub

Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim Conn As New ADODB.Connection
        Dim Rsd As New ADODB.Recordset
        Dim strSqlCmdSelc As String

        'StdDateTime()

        With Conn
                  If .State Then .Close()

                        .ConnectionString = strConnAdodb
                        .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                        .ConnectionTimeout = 90
                        .CommandTimeout = 10
                        .Open()

        End With

                        strSqlCmdSelc = "SELECT * FROM usermst (NOLOCK)" _
                                                     & " WHERE sta_usr =0" _
                                                     & " ORDER BY user_id"

                        Rsd = New ADODB.Recordset

                        With Rsd

                                    .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                                    .LockType = ADODB.LockTypeEnum.adLockOptimistic
                                    .Open(strSqlCmdSelc, Conn, , , )

                                    If .RecordCount <> 0 Then
                Do While Not .EOF   'ǹ�ٻ���֧��÷Ѵ�ش���¢ͧ��Ǵ�
                    cboUser.Items.Add(.Fields("user_id").Value)
                    .MoveNext()
                Loop

                                    End If

                        End With
                        Rsd.ActiveConnection = Nothing
                        Rsd.Close()
                        Rsd = Nothing

         Conn.Close()
         Conn = Nothing

End Sub

Private Sub lklClose_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lklClose.LinkClicked

  If FormCount("frmMainpro") > 0 Then '�ʴ���ҿ������ѡ�١�Դ����
      Me.Hide()
      frmMainpro.Enabled = True
      frmMainpro.Focus()
  Else
      Me.Close()
  End If

End Sub

Private Sub txtPass_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPass.GotFocus
        With txtPass
             .Select(0, .Text.Length)
        End With
End Sub

Private Sub txtPass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPass.KeyDown
        Select Case e.KeyCode
            Case 38, 40  '�١�â��, �١��ŧ
                cboUser.Focus()
        End Select
End Sub

Private Sub txtPass_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPass.KeyPress
        If e.KeyChar = Chr(13) Then   '�� Enter
            BeforeLoginProgram()
        End If
End Sub

Private Sub cboUser_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUser.GotFocus

   With cboUser
        .Select(0, .Text.Length)
   End With

End Sub

Private Sub cboUser_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cboUser.KeyPress

    If e.KeyChar = Chr(13) Then
        txtPass.Focus()
    End If

End Sub

Private Sub cboUser_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUser.LostFocus

    With cboUser
         .Text = .Text.ToString.ToUpper.Trim    '��ʴ��繵�Ǿ�����˭�
    End With

End Sub

Private Sub cboUser_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboUser.TextChanged
Dim btyCharStart As Byte

        With cboUser
            btyCharStart = .SelectionStart()        '��� BtyCharStart = �ش�������
            .Text = .Text.ToUpper                   '��� cboUser �繵�Ǿ�����˭�
            .SelectionStart = btyCharStart          '�����Ҩ����������˹������ focus �յ��˹��á 

    End With

End Sub

Private Sub BeforeLoginProgram()

    Dim Conn As New ADODB.Connection
    Dim strUserId As String
    Dim strPassWord As String

    Dim strDept As String = ""
    Dim strPost As String = ""
    Dim strSname As String = ""
    Dim strLevel As String = ""
    Dim IsPermitt As Boolean

    Dim strCmdSQL As String
    Dim strDate As String
    Dim strTime As String
    Dim strIpAddress As String
    Dim datLogin As Date = Now()    '�Ѻ����ѹ���Ѩ�غѹ

    Dim bytStaUsr As Byte
    Dim strUserPermiss As String = ""

          strUserId = ChangeFloat(cboUser.Text.ToUpper.Trim)
          strPassWord = txtPass.Text.ToUpper.Trim

          If Len(strUserId) <> 0 Then

                If Len(strPassWord) <> 0 Then

                        IsPermitt = CheckUserName(strUserId, strPassWord, strDept, strPost, strSname, strLevel)

                        If IsPermitt Then

                            Try

                                    strIpAddress = GetIPuserLogin()       '�Ѻ��� IP Address
                                    bytStaUsr = ThiefUser(strUserId, strPassWord, strIpAddress)  '����� userlogin �١������ҹ �������

                                    If bytStaUsr = 0 Then 'User ����դ���

                                            With Conn                                           
                                                 If .State Then .Close()
                                                    .ConnectionString = strConnAdodb
                                                    .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                                    .ConnectionTimeout = 90
                                                    .Open()
                                            End With

                                            If FormCount("frmMainpro") > 0 Then '�ʴ���ҿ������ѡ�١�Դ����

                                                  '-----------------------------------------������������� ���� LogIn ����---------------------------

                                                    strCmdSQL = "UPDATE usermst SET isexist =0" _
                                                                    & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"

                                                    Conn.Execute(strCmdSQL)

                                            End If

                                            '-----------------------------------------�������ż����ҹ------------------------------------------------------

                                            StdDateTime()
                                            strDate = datLogin.Date.ToString("yyyy-MM-dd")
                                            StdDateTimeThai()

                                            strTime = datLogin.TimeOfDay.ToString.Substring(0, 8)    '�Ѵ�����͡��
                                            strCmdSQL = "UPDATE usermst SET log_date ='" & strDate & "'" _
                                                      & "," & "log_time='" & strTime & "'" _
                                                      & "," & "com_ip ='" & Mid(strIpAddress, 1, 13) & "'" _
                                                      & " WHERE user_id ='" & strUserId & "'"
                                            Conn.Execute(strCmdSQL)

                                            '-----------------------------------------Lock�����ҹ------------------------------------------------------
                                            strCmdSQL = "UPDATE usermst SET isexist =1" _
                                                      & " WHERE user_id ='" & strUserId & "'"
                                            Conn.Execute(strCmdSQL)

                                            Conn.Close()
                                            Conn = Nothing

                                            Me.Hide()
                                            With frmMainPro

                                                    .Show()
                                                    .Enabled = True

                                                    .lblLogin.Text = strUserId
                                                    .lblUsrName.Text = "User Name : " & strSname
                                                    .lblIp.Text = "IP Address : " & strIpAddress.ToString
                                                    '.lblPost.Text = strPost

                                                    '-------------------��͹���� Administrator------------------------

                                                    If strLevel = "A" Then
                                                        .mnFileSys.Visible = True
                                                        .lblIcon.Image = My.Resources.admin
                                                    Else
                                                        .mnFileSys.Visible = False
                                                        .lblIcon.Image = My.Resources.users
                                                    End If

                                            End With

                                            InputUsrUsedProgram()


                                     Else

                                         MsgBox("UserName : " & strUserId & vbNewLine _
                                                     & "�ռ������������!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "UserName In Used!")
                                        'frmMainpro.Hide()
                                         cboUser.Focus()


                                     End If


                              Catch ex As Exception


                                        MsgBox("UserName : " & strUserId & vbNewLine _
                                                     & "�ռ������������!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "UserName In Used!")
                                        'frmMainpro.Hide()
                                        cboUser.Focus()

                              End Try

                        Else

                            MsgBox("Username ���� Password" & vbNewLine _
                                         & "���١��ͧ! �ô�к�����", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Wrong Username and Password!")
                            txtPass.Text = ""
                            txtPass.Focus()
                            'cboUser.Focus()

                        End If

                Else

                    MsgBox("�ô�к� PassWord ��͹�����ҹ!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "PassWord Not Define!")
                    txtPass.Focus()

                End If

          Else

             MsgBox("�ô�к� UserName ��͹�����ҹ!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "UserName Not Define!")
             cboUser.Focus()

          End If


End Sub

    '------------------------ �ѧ���� GetIPAddress ----------------------------------------
Private Function GetIPuserLogin() As String

  Dim IPHEntry As IPHostEntry    'IP ����Ѻ�����
  Dim IPAdd() As IPAddress      '�����ê�Դ Array
  Dim localHost As String
  Dim strIpMerge As String = ""
  Dim strIp As String = ""

      localHost = Dns.GetHostName()    '�Ѻ��Ҫ�������ͧ
      'IPHEntry = Dns.GetHostByName(localHost)
      IPHEntry = Dns.GetHostEntry(localHost)    'GetHostEntry ����� ���໪ system.net
      IPAdd = IPHEntry.AddressList
      Dim i As Integer

        For i = 0 To IPAdd.GetUpperBound(0)     'Index����ش���¢ͧ IPAdd
            'Console.Write("IP Address {0}: {1} ", i, IPAdd(i).ToString)
            strIp = (IPAdd(i).ToString())
        Next
        'Console.ReadLine()
        GetIPuserLogin = strIp

End Function

Private Sub lklLogin_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lklLogin.LinkClicked
    BeforeLoginProgram()
End Sub

Private Sub frmLogin_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        cboUser.Focus()
End Sub

'-----------------------------------------�Ѻ�ӹǹ�����ҹ------------------------------------------------------
Private Sub InputUsrUsedProgram()

Dim Conn As New ADODB.Connection
Dim ConnApp As New ADODB.Connection

Dim Rsd As New ADODB.Recordset
Dim strSqlCmdSelc As String
Dim strSqlCmdApp As String
Dim intUserQty As Integer

      With Conn

             If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .CommandTimeout = 10
                .Open()

       End With

       strSqlCmdSelc = "SELECT user_id FROM usermst (NOLOCK)" _
                                        & " WHERE isexist =1"

        With Rsd

                .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                .LockType = ADODB.LockTypeEnum.adLockOptimistic
                .Open(strSqlCmdSelc, Conn, , , )

                 If .RecordCount <> 0 Then
                        intUserQty = .RecordCount
                 Else
                        intUserQty = 0
                 End If

        End With

        Rsd.ActiveConnection = Nothing
        Rsd.Close()
        Rsd = Nothing

        Conn.Close()
        Conn = Nothing

        '-----------------------------------------���ӹǹ�����ҹ------------------------------------------------------                
        With ConnApp

                  If .State Then .Close()

                     .ConnectionString = strConnAdodbApp
                     .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                     .ConnectionTimeout = 90
                     .Open()

                      strSqlCmdApp = "UPDATE appname SET usr_logon =" & intUserQty.ToString _
                                                                  & " WHERE app_id ='app12'"
                    .Execute(strSqlCmdApp)
                    .Close()

        End With
        ConnApp = Nothing


End Sub

Private Sub ClearUserOutofSystem()

Dim Conn As New ADODB.Connection
Dim strCmdSQL As String
Dim strUserId As String

        With Conn

             If .State Then .Close()
                .ConnectionString = strConnAdodb
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionTimeout = 90
                .CommandTimeout = 10
                .Open()

                '-----------------------------------------UnLock�����ҹ------------------------------------------------------

                strUserId = ChangeFloat(cboUser.Text.ToUpper.Trim)
                strCmdSQL = "UPDATE usermst SET isexist =0" _
                                & " WHERE user_id ='" & strUserId & "'"
                .Execute(strCmdSQL)
                .Close()

        End With
        Conn = Nothing

End Sub

'-------------------- �ѧ������ ���� User --------------------------------------
Function ThiefUser(ByVal strUserName As String, ByVal strPassword As String, ByVal strIpAddress As String) As Byte

Dim Conn As New ADODB.Connection
Dim Rsd As New ADODB.Recordset
Dim strSqlCmdSelc As String

              With Conn

                        If .State Then .Close()
                           .ConnectionString = strConnAdodb
                           .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                           .ConnectionTimeout = 90
                           .CommandTimeout = 10
                           .Open()

              End With

              strSqlCmdSelc = "SELECT user_id,pass,isexist,com_ip FROM usermst (NOLOCK)" _
                                    & " WHERE user_id ='" & strUserName & "'" _
                                    & " AND pass ='" & strPassword & "'"

              Rsd = New ADODB.Recordset

              With Rsd

                        .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                        .LockType = ADODB.LockTypeEnum.adLockOptimistic
                        .Open(strSqlCmdSelc, Conn, , , )

                          If .RecordCount <> 0 Then

                                If strIpAddress <> .Fields("com_ip").Value.ToString.Trim Then
                                   ThiefUser = .Fields("isexist").Value
                                Else
                                   ThiefUser = 0   '����
                                End If

                          Else
                                ThiefUser = 0
                          End If

              End With

              Rsd.ActiveConnection = Nothing
              Rsd.Close()
              Rsd = Nothing

             Conn.Close()
             Conn = Nothing

End Function

Private Function ChangeFloat(ByVal strNumber As String) As String
Dim i, x As Integer

Dim strTmp As String = ""
Dim strMerge As String = ""


        x = Len(strNumber.ToString)     '�Ѻ�ӹǹ����ѡ��

                        For i = 1 To x

                                strTmp = Mid(strNumber.ToString, i, 1)
                                Select Case strTmp
                                          Case Is = "'"
                                          Case Is = ""
                                          Case Else
                                                 strMerge = strMerge & strTmp
                                End Select
                         Next i



ChangeFloat = strMerge

End Function

    '------------------------ �ѧ���蹹Ѻ�ӹǹ����� ---------------------------------------------------------
Private Function FormCount(ByVal frmName As String) As Long
Dim frm As Form

    For Each frm In My.Application.OpenForms

                If frm Is My.Forms.frmMainpro Then
                    FormCount = FormCount + 1
                End If
    Next

End Function

End Class