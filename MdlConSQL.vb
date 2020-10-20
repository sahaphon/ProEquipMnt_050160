Imports System.IO

Module MdlConSQL

    'OLE DB Services = -2 ����������͡ѹ  ConnectionRead Error �������Դ��������ҹ� �ѹ������ = 30/4/54
    'sahaphon
    'ADDASRV03
    Public Const strConnAdodb = "Provider = sqloledb;" & _
                                                "Data Source=ADDASRV03;" & _
                                                "Initial Catalog=DBequipmnt;" & _
                                                "User ID=Sa;" & _
                                                "Password=Sa2008"

    Public Const strConnDbHr2 = "Provider = sqloledb;" & _
                                                "Data Source=ADDASRV03;" & _
                                                "Initial Catalog=DBhr2;" & _
                                                "User ID=Sa;" & _
                                                "Password=Sa2008"

    Public Const strConnAdodbApp = "Provider = sqloledb;" & _
                                                "Data Source=ADDASRV03;" & _
                                                "OLE DB Services=-2;" & _
                                                "Initial Catalog=DBappname;" & _
                                                "User ID=Sa;" & _
                                                "Password=Sa2008"

   Public Const sqlclint = " Server=ADDASRV03;Database=DBequipmnt;User Id=Sa;Password=Sa2008;"  '\\10.32.0.16\data2\WIPNEW\
   Public Const dbase = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=\\10.32.0.16\data2\WIPNEW\;Extended Properties=dBase IV"


Public Function App_path() As String
   Return System.AppDomain.CurrentDomain.BaseDirectory
End Function

Public Sub StdDateTime() '�ѹ�����͹���ҡ�

   Dim ct As New System.Globalization.CultureInfo("en-US")
       ct.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
       System.Threading.Thread.CurrentThread.CurrentCulture = ct

End Sub

Public Sub StdDateTimeThai() '�ѹ�����͹����

   Dim ct As New System.Globalization.CultureInfo("th-TH", True)
       ct.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy"
       System.Threading.Thread.CurrentThread.CurrentCulture = ct

End Sub

Public Function CheckUserName(ByVal strUsrName As String, ByVal strUsrPass As String, _
                                                                              ByRef strRntDept As String, ByRef strRntPost As String, _
                                                                             ByRef strRntSname As String, ByRef strRntLevel As String) As Boolean
  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlCmdSelc As String
  Dim IsExist As Boolean

      With Conn

           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()

      End With

                strSqlCmdSelc = "SELECT * FROM usermst (NOLOCK)" _
                                     & " WHERE user_id ='" & strUsrName & "'" _
                                     & " AND pass ='" & strUsrPass & "'" _
                                     & " AND sta_usr =0"

                With Rsd
                     .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                     .LockType = ADODB.LockTypeEnum.adLockOptimistic
                     .Open(strSqlCmdSelc, Conn, , , )

                     If .RecordCount <> 0 Then
                        strRntSname = .Fields("sname").Value.ToString.Trim
                        strRntPost = .Fields("post").Value.ToString.Trim
                        strRntDept = .Fields("dept").Value.ToString.Trim
                        strRntLevel = .Fields("act_usr").Value.ToString.Trim
                        IsExist = True
                     Else
                        strRntPost = ""
                        strRntSname = ""
                        strRntDept = ""
                        strRntLevel = ""
                        IsExist = False
                     End If

                End With

        Rsd.ActiveConnection = Nothing
        Rsd.Close()
        Rsd = Nothing

    Conn.Close()
    Conn = Nothing

    CheckUserName = IsExist

End Function

Public Function ReplaceQuote(ByVal strString As String)
   ReplaceQuote = Replace(strString, "'", "''")
End Function

Public mdiHost As MdiClient
Public backgrounds As Image() = {My.Resources.wall_prog}
Public backgroundIndex As Integer = -1

Public Sub ImageBackground()

    For Each ctl As Control In frmMainPro.Controls
        If TypeOf ctl Is MdiClient Then
           mdiHost = DirectCast(ctl, MdiClient)
           Exit For
        End If
    Next ctl

SetBackgroundImage()
End Sub

Public Sub SetBackgroundImage()

 backgroundIndex += 1

    If backgroundIndex = backgrounds.Length Then
       backgroundIndex = 0
    End If
       mdiHost.BackgroundImage = backgrounds(backgroundIndex)

End Sub

Public Sub MsnAdmin()
   MsgBox("�س������Է�����������ǹ���!!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "User LogIn Access is denied!!")
End Sub

Public Function CheckUserLevel(ByVal strUsrName As String) As String

  Dim Conn As New ADODB.Connection
  Dim Rsd As New ADODB.Recordset
  Dim strSqlCmdSelc As String

      With Conn

           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()

     End With

                strSqlCmdSelc = "SELECT act_usr FROM usermst (NOLOCK)" _
                                               & " WHERE user_id ='" & strUsrName & "'"

                With Rsd

                     .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                     .LockType = ADODB.LockTypeEnum.adLockOptimistic
                     .Open(strSqlCmdSelc, Conn, , , )

                     If .RecordCount <> 0 Then
                        CheckUserLevel = .Fields("act_usr").Value.ToString.Trim

                     Else
                          CheckUserLevel = ""

                     End If

                End With

     Rsd.ActiveConnection = Nothing
     Rsd.Close()
     Rsd = Nothing

  Conn.Close()
  Conn = Nothing

End Function

Public Function CheckUserEntry(ByVal ObjCode As String, ByVal Docfield As String) As Boolean

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset

 Dim strSqlCmdSelc As String
 Dim strSqlCmd As String

 Dim datLogin As Date = Now()
 Dim strTime As String
 Dim strDate As String

 Dim dubCounter As Double

     With Conn
          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
    End With

                strSqlCmdSelc = "SELECT user_id,file_icon,open_cnt FROM usertrn (NOLOCK) " _
                                     & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                     & " AND file_icon ='" & ObjCode & "'" _
                                     & " AND " & Docfield & "=1"                           '��Ǵ� act_xxx

                With Rsd

                     .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                     .LockType = ADODB.LockTypeEnum.adLockOptimistic
                     .Open(strSqlCmdSelc, Conn, , , )

                     If .RecordCount <> 0 Then
                        CheckUserEntry = True

                       '----------------------------------------�������š�������----------------------------------------------

                        strTime = datLogin.TimeOfDay.ToString.Substring(0, 8)
                        strDate = datLogin.Date.ToString("yyyy-MM-dd")
                        dubCounter = .Fields("open_cnt").Value + 1                      '�Ѻ�ӹǹ���������ҹ

                       '------------------------------ Update ������ŧ���ҧ usertrn -------------------------------------------  

                       strSqlCmd = "UPDATE usertrn SET last_date ='" & strDate & "'" _
                                           & "," & "last_time  ='" & strTime & "'" _
                                           & "," & "open_cnt =" & dubCounter.ToString _
                                           & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'" _
                                           & " AND file_icon ='" & ObjCode & "'"

                       Conn.Execute(strSqlCmd)

                     Else
                          CheckUserEntry = False
                     End If

                     Rsd.ActiveConnection = Nothing
                     Rsd.Close()
                End With

   Conn.Close()
   Conn = Nothing

End Function

Public Function ActualValue(ByVal dubNumber As Double) As String

   If dubNumber = Int(dubNumber) Then '0
      ActualValue = Format(dubNumber, "#,##0")
   Else

        If dubNumber * 10 = Int(dubNumber * 10) Then '1
           ActualValue = Format(dubNumber, "#,##0.0#")
        Else

                If dubNumber * 100 = Int(dubNumber * 100) Then '2
                   ActualValue = Format(dubNumber, "#,##0.0#")
                Else

                    If dubNumber * 1000 = Int(dubNumber * 1000) Then '3
                       ActualValue = Format(dubNumber, "#,##0.0###")
                    Else

                        If dubNumber * 10000 = Int(dubNumber * 10000) Then '4
                           ActualValue = Format(dubNumber, "#,##0.0###")
                        Else

                            If dubNumber * 100000 = Int(dubNumber * 100000) Then '5
                               ActualValue = Format(dubNumber, "#,##0.0###")
                            Else
                                ActualValue = Format(dubNumber, "#,##0.0###")
                            End If

                         End If

                    End If

                End If

         End If

   End If

End Function

Public Function CallUserName(ByVal strUserId As String) As String

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmdSelc As String

     With Conn
          If .State Then .Close()
             .ConnectionString = strConnAdodb
             .CursorLocation = ADODB.CursorLocationEnum.adUseClient
             .ConnectionTimeout = 90
             .Open()
     End With

                strSqlCmdSelc = "SELECT sname" _
                                     & " FROM usermst (NOLOCK) " _
                                     & " WHERE user_id ='" & strUserId & "'"

                With Rsd
                     .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
                     .LockType = ADODB.LockTypeEnum.adLockOptimistic
                     .Open(strSqlCmdSelc, Conn, , , )

                     If .RecordCount <> 0 Then
                        CallUserName = .Fields("sname").Value.ToString.Trim
                     Else
                         CallUserName = ""
                     End If

                End With

        Rsd.ActiveConnection = Nothing
        Rsd.Close()
        Rsd = Nothing

    Conn.Close()
    Conn = Nothing

End Function

Public Function CallPathSignPicture(ByVal strUserId As String) As String   '�Ҹ��������� 10.32.0.14

 Dim Conn As New ADODB.Connection
 Dim Rsd As New ADODB.Recordset
 Dim strSqlCmdSelc As String

     With Conn
           If .State Then .Close()
              .ConnectionString = strConnAdodb
              .CursorLocation = ADODB.CursorLocationEnum.adUseClient
              .ConnectionTimeout = 90
              .Open()

     End With

     strSqlCmdSelc = "SELECT pic_sign" _
                                 & " FROM usermst (NOLOCK) " _
                                 & " WHERE user_id ='" & strUserId & "'"

     With Rsd

          .CursorType = ADODB.CursorTypeEnum.adOpenKeyset
          .LockType = ADODB.LockTypeEnum.adLockOptimistic
          .Open(strSqlCmdSelc, Conn, , , )

          If .RecordCount <> 0 Then
             CallPathSignPicture = .Fields("pic_sign").Value.ToString.Trim

          Else
             CallPathSignPicture = ""
          End If

     .ActiveConnection = Nothing
     .Close()

     End With

 Conn.Close()
 Conn = Nothing

End Function

Public Function SaveChangeThaYear(ByVal strDate As String) As String 'yyyy-MM-dd ����

Dim strCon1 As String
Dim strCon2 As String

    If Val(Mid(strDate, 1, 4)) > 2500 Then               '�ѧ��ѹ Val �觤�ҡ�Ѻ�繵���Ţ

       strCon1 = Trim(Str(Val(Mid(strDate, 1, 4))))     '�Ѵ��һ�   Function VAL() �ŧ���ʵ�ԧ�� Numeric
       strCon2 = Trim(Mid(strDate, 5, 6))               '�Ѵ�����͹ ��� �ѹ���
       SaveChangeThaYear = strCon1 & strCon2            '����� yyyy/MM/dd

   Else
        strCon1 = Trim(Str(Val(Mid(strDate, 1, 4)) + 543))
        strCon2 = Trim(Mid(strDate, 5, 6))
        SaveChangeThaYear = strCon1 & strCon2

  End If

End Function

Public Function SaveChangeEngYear(ByVal strDate As String) As String 'yyyy-MM-dd

 Dim strCon1 As String
 Dim strCon2 As String

     If Val(Mid(strDate, 1, 4)) < 2500 Then               '�ѧ��ѹ Val �觤�ҡ�Ѻ�繵���Ţ
        strCon1 = Trim(Str(Val(Mid(strDate, 1, 4))))     '�Ѵ��һ� - 543   Function VAL() �ŧ���ʵ�ԧ�� Numeric
        strCon2 = Trim(Mid(strDate, 5, 6))               '�Ѵ�����͹ ��� �ѹ���
        SaveChangeEngYear = strCon1 & strCon2            '����� yyyy/MM/dd
     Else
         strCon1 = Trim(Str(Val(Mid(strDate, 1, 4)) - 543))
         strCon2 = Trim(Mid(strDate, 5, 6))
         SaveChangeEngYear = strCon1 & strCon2

     End If

End Function

    Public Function ShowChangeEngYear(ByVal strDate As String) As String 'dd/MM/yyyy

        Dim strCon1 As String
        Dim strCon2 As String

        strCon1 = Trim(Str(Val(Mid(strDate, 7, 4)) - 543))
        strCon2 = Trim(Mid(strDate, 1, 6))
        ShowChangeEngYear = strCon2 & strCon1

    End Function

    Public Function ShowChangeThaiYear(ByVal strDate As String) As String 'dd/MM/yyyy

        Dim strCon1 As String
        Dim strCon2 As String

        strCon1 = Trim(Str(Val(Mid(strDate, 7, 4)) + 543))
        strCon2 = Trim(Mid(strDate, 1, 6))
        ShowChangeThaiYear = strCon2 & strCon1

    End Function

    Public Function ChangFormat(ByVal strNumber As String) As String

        Dim i, x As Integer

        Dim strTmp As String = ""
        Dim strMerge As String = ""

        x = Len(strNumber.ToString)

        For i = 1 To x

            strTmp = Mid(strNumber.ToString, i, 1)
            Select Case strTmp
                Case Is = "+"
                Case Is = ","
                Case Is = "_"
                Case Else
                    If InStr("-.0123456789", strTmp) > 0 Then
                        strMerge = strMerge & strTmp
                    End If
            End Select
        Next i

        ChangFormat = strMerge

    End Function

    '�ѧ���ѹ��ŧ  image ��   Base64 String 
    Public Function ImageToBase64(ByVal image As Image, ByVal format As System.Drawing.Imaging.ImageFormat) As String
  Using ms As New MemoryStream()

  '�ŧ Image to byte[] 
  image.Save(ms, format)
  Dim imageBytes As Byte() = ms.ToArray()

  '�ŧ byte[] �� Base64 String 
  Dim base64String As String = Convert.ToBase64String(imageBytes)
  Return base64String

  End Using

End Function

'�ѧ���ѹ��ŧ Base64 String �� image 
Public Function Base64ToImage(ByVal base64String As String) As Image
  ' �ŧ Base64 String to byte[] 
  Dim imageBytes As Byte() = Convert.FromBase64String(base64String)
  Dim ms As New MemoryStream(imageBytes, 0, imageBytes.Length)

      ' �ŧ byte[] �� Image 
      ms.Write(imageBytes, 0, imageBytes.Length)
  Dim image1 As Image = Image.FromStream(ms, True)

Return image1

End Function

Public Sub ClearTmpTableUser(ByVal strTmpTableName As String)

   Dim Conn As New ADODB.Connection
   Dim strSqlCmd As String = ""

       With Conn
         If .State Then .Close()
            .ConnectionString = strConnAdodb
            .CursorLocation = ADODB.CursorLocationEnum.adUseClient
            .ConnectionTimeout = 90
            .Open()
      End With

      strSqlCmd = "Delete FROM " & strTmpTableName _
                             & " WHERE user_id ='" & frmMainPro.lblLogin.Text & "'"

      Conn.Execute(strSqlCmd)

   Conn.Close()
   Conn = Nothing

End Sub

Public Function ScaleImage(ByVal OldImage As Image, ByVal TargetHeight As Integer, ByVal TargetWidth As Integer) As System.Drawing.Image

  Dim NewHeight As Integer = TargetHeight
  Dim NewWidth As Integer = NewHeight / OldImage.Height * OldImage.Width

      If NewWidth > TargetWidth Then
         NewWidth = TargetWidth
         NewHeight = NewWidth / OldImage.Width * OldImage.Height
      End If

      Return New Bitmap(OldImage, NewWidth, NewHeight)
End Function

Public Function DateTimeCutString(ByVal strDate As String) As String  '�ѧ��������ѹ������� ��ͷ����ٻ�Ҿ
 Dim newDateTime As String
 Dim dt1, dt2 As String
 Dim resultDate As String = ""
 Dim resultTime As String = ""
 Dim strArr(), strArr2() As String
 Dim i As Integer

    dt1 = Mid(strDate, 1, 10) '�Ѵ����ѹ��͹��
    dt2 = Mid(strDate, 12, 8) '�Ѵ�������

     '�Ѵ / ��ѹ���
     strArr = dt1.Split("/")
     For i = 0 To strArr.Length - 1
         resultDate = resultDate & strArr(i)
     Next

    '�Ѵ : ������������
    strArr2 = dt2.Split(":")
    For i = 0 To strArr2.Length - 1
        resultTime = resultTime & strArr2(i)
    Next

     newDateTime = resultDate & resultTime
    Return newDateTime

End Function

    Public Function ReturnImageName(ByVal strPic As String) As String   '�׹��Ҫ�������ٻ��ԧ(������ѹ������)

        Dim IntLengTxt As Integer
        Dim newFielnm As String = ""
        Dim newFielnm2 As String = ""
        Dim typeNm As String = ""

        Try

            IntLengTxt = strPic.Length  '�ӹǹ����ѡ�âͧ������������

            If IntLengTxt > 18 Then  '�ó�����������ҧ���� + �����ѹ��� ���� ��ͷ������
                '������ٻ�຺������������ѹ������� �������
                newFielnm = Mid(strPic, (IntLengTxt - 18) + 1, 14)

                If IsNumeric(newFielnm) Then   '�ó��ա�������ѹ������㹪������
                    '�Ѵ���੾�Ъ������ 
                    newFielnm = Mid(strPic, 1, IntLengTxt - 19)  '�Ѵ�����á �֧ (������ - 18) 18 ��ͨӹǹ�ѹ���������ʡ�����
                    '�Ѵ���੾�й��ʡ�����
                    typeNm = (Microsoft.VisualBasic.Right(strPic, 4))   '�Ѵ����ѡ�� 4 ��� ������ҡ��ҹ���
                    newFielnm2 = newFielnm & typeNm  '���������ѧ�Ѵ�ѹ����͡
                Else
                    newFielnm2 = strPic
                End If

            Else   '�ó�����ա�������ѹ������Ҫ������
                newFielnm2 = strPic
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Return newFielnm2
    End Function

End Module

