Imports System.Net.Mail
Imports System.Net
Imports System.Web.HttpContext
Imports System.Web.Hosting
Imports System.Net.Mime
Imports System.Threading

Namespace Common.tools

    ''' <summary>
    ''' Gmail的郵件類別
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EmailMod

        Private smtpName As String      'smtp伺服器名稱
        Private smtpPort As Int16       'smtp埠號
        Private smtpAcn As String       'smtp帳號
        Private smtpPwd As String       'smtp密碼
        Private sslConnect As Boolean      '是否使用 SSL 加密連線

        ''' <summary>
        ''' 使用自訂的SMTP設定檔寄送郵件(使用預設郵寄埠號25)
        ''' </summary>
        ''' <param name="smtpName">smtp帳號</param>
        ''' <param name="smtpAcn">smtp帳號</param>
        ''' <param name="smtpPwd">smtp密碼</param>
        ''' <param name="sslConnect">是否使用 SSL 連線</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal smtpName As String, ByVal smtpAcn As String, ByVal smtpPwd As String, ByVal sslConnect As Boolean)
            Me.New(smtpName, 25, smtpAcn, smtpPwd, sslConnect)
        End Sub

        ''' <summary>
        ''' 使用自訂的SMTP設定檔寄送郵件
        ''' </summary>
        ''' <param name="smtpName">smtp帳號</param>
        ''' <param name="smtpPort">smtp埠號</param>
        ''' <param name="smtpAcn">smtp帳號</param>
        ''' <param name="smtpPwd">smtp密碼</param>
        ''' <param name="sslConnect">是否使用 SSL 連線</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal smtpName As String, ByVal smtpPort As Int16, ByVal smtpAcn As String, ByVal smtpPwd As String, ByVal sslConnect As Boolean)
            Me.smtpName = smtpName
            Me.smtpPort = smtpPort
            Me.smtpAcn = smtpAcn
            Me.smtpPwd = smtpPwd
            Me.sslConnect = sslConnect
        End Sub

        ''' <summary>
        ''' 發送純文字郵件，成功回傳1
        ''' </summary>
        ''' <param name="toPeople">收件者信箱地址，需用","隔開</param>
        ''' <param name="fromMail">發件者信箱</param>
        ''' <param name="fromName">發件者姓名</param>
        ''' <param name="subject">標題主旨</param>
        ''' <param name="content">信件內容</param>
        ''' <param name="isBcc">是否以密件副本方式寄送</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Overloads Function SendTextMail(ByVal toPeople As String, ByVal fromMail As String, ByVal fromName As String, ByVal subject As String,
                                    ByVal content As String, ByVal isBcc As Boolean) As Int16

            Dim result As Int16 = Nothing
            Dim msg As MailMessage = New MailMessage()      '郵件訊息主體

            Try
                Dim mailTo As String = toPeople.TrimStart(",").TrimEnd(",")
                If Not isBcc Then
                    msg.To.Add(mailTo)       '收件者
                Else
                    msg.Bcc.Add(mailTo)       '密件副本
                End If

                'msg.CC.Add()                        '副本

                Dim sender As MailAddress = New MailAddress(fromMail, fromName, System.Text.Encoding.UTF8)
                msg.From = sender           '寄件者
                msg.Sender = sender         '寄件者
                msg.Subject = subject       '標題主旨
                msg.SubjectEncoding = System.Text.Encoding.UTF8     '標題主旨編碼
                msg.BodyEncoding = System.Text.Encoding.UTF8        '信件內容編碼            
                msg.Priority = MailPriority.Normal      '優先等級
                msg.Body = content      '信件內容
                msg.IsBodyHtml = False             '信件內容是否為HTML
            Catch ex As Exception
                result = -1
                msg.Dispose()
            End Try

            If result <> -1 Then
                'SMTP
                Dim mySmtp As SmtpClient = New SmtpClient(Me.smtpName, Me.smtpPort)        'Smtp Server and Port

                mySmtp.EnableSsl = Me.sslConnect         '是否使用 SSL 連線 (Gmail需使用SSL)

                mySmtp.UseDefaultCredentials = IIf(Me.sslConnect, False, True)        '若開啟 SSL 驗證時需把預設驗證關閉，否則不用

                mySmtp.Credentials = New NetworkCredential(Me.smtpAcn, Me.smtpPwd)        '存取smtp的帳號密碼

                Try
                    mySmtp.Send(msg)        '發送信件

                    result = 1      '1為成功

                Catch ex As SmtpFailedRecipientsException       '捕捉Smtp的錯誤事件
                    result = -1      '-1為失敗                    
                Finally
                    msg.Dispose()
                    mySmtp.Dispose()
                End Try
            End If

            Return result
        End Function

        ''' <summary>
        ''' 發送純文字郵件，成功回傳1
        ''' </summary>
        ''' <param name="toPeople">收件者信箱地址，需用","隔開</param>
        ''' <param name="fromMail">發件者信箱</param>
        ''' <param name="fromName">發件者姓名</param>
        ''' <param name="subject">標題主旨</param>
        ''' <param name="content">信件內容</param>
        ''' <param name="isBcc">是否以密件副本方式寄送</param>
        ''' <param name="sleepTime">發送失敗系統重發間隔</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Overloads Function SendTextMail(ByVal toPeople As String, ByVal fromMail As String, ByVal fromName As String, ByVal subject As String,
                                    ByVal content As String, ByVal isBcc As Boolean, ByVal sleepTime As Int32) As Int16

            Dim result As Int16 = Nothing
            Dim msg As MailMessage = New MailMessage()      '郵件訊息主體

            Try
                Dim mailTo As String = toPeople.TrimStart(",").TrimEnd(",")
                If Not isBcc Then
                    msg.To.Add(mailTo)       '收件者
                Else
                    msg.Bcc.Add(mailTo)       '密件副本
                End If

                'msg.CC.Add()                        '副本

                Dim sender As MailAddress = New MailAddress(fromMail, fromName, System.Text.Encoding.UTF8)
                msg.From = sender           '寄件者
                msg.Sender = sender         '寄件者
                msg.Subject = subject       '標題主旨
                msg.SubjectEncoding = System.Text.Encoding.UTF8     '標題主旨編碼
                msg.BodyEncoding = System.Text.Encoding.UTF8        '信件內容編碼            
                msg.Priority = MailPriority.Normal      '優先等級
                msg.Body = content      '信件內容
                msg.IsBodyHtml = False             '信件內容是否為HTML
            Catch ex As Exception
                result = -1
                msg.Dispose()
            End Try

            If result <> -1 Then
                'SMTP
                Dim mySmtp As SmtpClient = New SmtpClient(Me.smtpName, Me.smtpPort)        'Smtp Server and Port

                mySmtp.EnableSsl = Me.sslConnect         '是否使用 SSL 連線 (Gmail需使用SSL)

                mySmtp.UseDefaultCredentials = IIf(Me.sslConnect, False, True)        '若開啟 SSL 驗證時需把預設驗證關閉，否則不用

                mySmtp.Credentials = New NetworkCredential(Me.smtpAcn, Me.smtpPwd)        '存取smtp的帳號密碼

                Try
                    mySmtp.Send(msg)        '發送信件

                    result = 1      '1為成功

                Catch ex As SmtpFailedRecipientsException       '捕捉Smtp的錯誤事件
                    result = -1      '-1為失敗
                    For i = 0 To ex.InnerExceptions.Length
                        Dim status As SmtpStatusCode = ex.InnerExceptions(i).StatusCode()       '建立SmtpStatusCode物件取得Smtp代碼
                        If status = SmtpStatusCode.MailboxBusy Or status = SmtpStatusCode.MailboxUnavailable Then       '判斷是否為Smtp忙碌或者是無法取存取的信件
                            Thread.Sleep(sleepTime)      '暫停指定時間後重發
                            mySmtp.Send(msg)        '發送信件
                        End If
                    Next
                Finally
                    msg.Dispose()
                    mySmtp.Dispose()
                End Try
            End If

            Return result
        End Function

        ''' <summary>
        ''' 發送HTML郵件，成功回傳1
        ''' </summary>
        ''' <param name="toPeople">收件者信箱地址，需用","隔開</param>
        ''' <param name="fromMail">發件者信箱</param>
        ''' <param name="fromName">發件者姓名</param>
        ''' <param name="subject">標題主旨</param>
        ''' <param name="content">信件內容</param>
        ''' <param name="isBcc">是否以密件副本方式寄送</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Overloads Function SendHtmlMail(ByVal toPeople As String, ByVal fromMail As String, ByVal fromName As String, ByVal subject As String,
                                    ByVal content As String, ByVal isBcc As Boolean) As Int16

            Dim result As Int16 = Nothing
            Dim msg As MailMessage = New MailMessage()      '郵件訊息主體

            Try
                Dim mailTo As String = toPeople.TrimStart(",").TrimEnd(",")
                If Not isBcc Then
                    msg.To.Add(mailTo)       '收件者
                Else
                    msg.Bcc.Add(mailTo)       '密件副本
                End If

                'msg.CC.Add()                        '副本

                Dim sender As MailAddress = New MailAddress(fromMail, fromName, System.Text.Encoding.UTF8)
                msg.From = sender       '寄件者
                msg.Sender = sender     '寄件者
                msg.Subject = subject       '標題主旨
                msg.SubjectEncoding = System.Text.Encoding.UTF8     '標題主旨編碼
                msg.BodyEncoding = System.Text.Encoding.UTF8        '信件內容編碼            
                msg.Priority = MailPriority.Normal      '優先等級
            Catch ex As Exception
                result = -1
                msg.Dispose()
            End Try

            If result <> -1 Then
                Dim appSets As NameValueCollection = ConfigurationManager.AppSettings()     '取得WebConfig內AppSetting的集合

                For Each key In appSets.Keys()
                    If content.Contains(key) Then                '判斷傳入的HTML內容內是否有與WebConfig相對應的Key
                        content.Replace(key, appSets(key))                  '若有便將內容相對應的Key換成WebConfig相對應的路徑
                    End If
                Next
                msg.Body = content
                msg.IsBodyHtml = True

                'SMTP
                Dim mySmtp As SmtpClient = New SmtpClient(Me.smtpName, Me.smtpPort)        'Smtp Server and Port

                mySmtp.EnableSsl = Me.sslConnect         '是否使用 SSL 連線 (Gmail需使用SSL)

                mySmtp.UseDefaultCredentials = IIf(Me.sslConnect, False, True)        '若開啟 SSL 驗證時需把預設驗證關閉，否則不用

                mySmtp.Credentials = New NetworkCredential(Me.smtpAcn, Me.smtpPwd)        '存取smtp的帳號密碼

                Try
                    mySmtp.Send(msg)        '發送信件

                    result = 1      '1為成功

                Catch ex As SmtpFailedRecipientsException       '捕捉Smtp的錯誤事件
                    result = -1      '-1為失敗                    
                Finally
                    msg.Dispose()
                    mySmtp.Dispose()
                End Try
            End If

            Return result
        End Function

        ''' <summary>
        ''' 發送HTML郵件，成功回傳1
        ''' </summary>
        ''' <param name="toPeople">收件者信箱地址，需用","隔開</param>
        ''' <param name="fromMail">發件者信箱</param>
        ''' <param name="fromName">發件者姓名</param>
        ''' <param name="subject">標題主旨</param>
        ''' <param name="content">信件內容</param>
        ''' <param name="isBcc">是否以密件副本方式寄送</param>
        ''' <param name="sleepTime">發送失敗系統重發間隔</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Overloads Function SendHtmlMail(ByVal toPeople As String, ByVal fromMail As String, ByVal fromName As String, ByVal subject As String,
                                    ByVal content As String, ByVal isBcc As Boolean, ByVal sleepTime As Int32) As Int16

            Dim result As Int16 = Nothing
            Dim msg As MailMessage = New MailMessage()      '郵件訊息主體

            Try
                Dim mailTo As String = toPeople.TrimStart(",").TrimEnd(",")
                If Not isBcc Then
                    msg.To.Add(mailTo)       '收件者
                Else
                    msg.Bcc.Add(mailTo)       '密件副本
                End If

                'msg.CC.Add()                        '副本

                Dim sender As MailAddress = New MailAddress(fromMail, fromName, System.Text.Encoding.UTF8)
                msg.From = sender       '寄件者
                msg.Sender = sender     '寄件者
                msg.Subject = subject       '標題主旨
                msg.SubjectEncoding = System.Text.Encoding.UTF8     '標題主旨編碼
                msg.BodyEncoding = System.Text.Encoding.UTF8        '信件內容編碼            
                msg.Priority = MailPriority.Normal      '優先等級
            Catch ex As Exception
                result = -1
                msg.Dispose()
            End Try

            If result <> -1 Then
                Dim appSets As NameValueCollection = ConfigurationManager.AppSettings()     '取得WebConfig內AppSetting的集合

                For Each key In appSets.Keys()
                    If content.Contains(key) Then                '判斷傳入的HTML內容內是否有與WebConfig相對應的Key
                        content.Replace(key, appSets(key))                  '若有便將內容相對應的Key換成WebConfig相對應的路徑
                    End If
                Next
                msg.Body = content
                msg.IsBodyHtml = True

                'SMTP
                Dim mySmtp As SmtpClient = New SmtpClient(Me.smtpName, Me.smtpPort)        'Smtp Server and Port

                mySmtp.EnableSsl = Me.sslConnect         '是否使用 SSL 連線 (Gmail需使用SSL)

                mySmtp.UseDefaultCredentials = IIf(Me.sslConnect, False, True)        '若開啟 SSL 驗證時需把預設驗證關閉，否則不用

                mySmtp.Credentials = New NetworkCredential(Me.smtpAcn, Me.smtpPwd)        '存取smtp的帳號密碼

                Try
                    mySmtp.Send(msg)        '發送信件

                    result = 1      '1為成功

                Catch ex As SmtpFailedRecipientsException       '捕捉Smtp的錯誤事件
                    result = -1      '-1為失敗
                    For i = 0 To ex.InnerExceptions.Length
                        Dim status As SmtpStatusCode = ex.InnerExceptions(i).StatusCode()       '建立SmtpStatusCode物件取得Smtp代碼
                        If status = SmtpStatusCode.MailboxBusy Or status = SmtpStatusCode.MailboxUnavailable Then       '判斷是否為Smtp忙碌或者是無法取存取的信件
                            Thread.Sleep(sleepTime)      '暫停指定時間後重發
                            mySmtp.Send(msg)        '發送信件
                        End If
                    Next
                Finally
                    msg.Dispose()
                    mySmtp.Dispose()
                End Try
            End If

            Return result
        End Function

        ''' <summary>
        ''' 發送HTML郵件，並將圖檔資料以附件方式送出，成功回傳1
        ''' </summary>
        ''' <param name="toPeople">收件者信箱地址，需用","隔開</param>
        ''' <param name="fromMail">發件者信箱</param>
        ''' <param name="fromName">發件者姓名</param>
        ''' <param name="subject">標題主旨</param>
        ''' <param name="content">信件內容</param>
        ''' <param name="isThread">是否處於多執行緒狀態下呼叫本函式</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Overloads Function SendAttImgMail(ByVal toPeople As String, ByVal fromMail As String, ByVal fromName As String, ByVal subject As String,
                                    ByVal content As String, ByVal isBcc As Boolean, ByVal isThread As Boolean) As Int16

            Dim result As Int16 = Nothing
            Dim msg As MailMessage = New MailMessage()      '郵件訊息主體

            Try
                Dim mailTo As String = toPeople.TrimStart(",").TrimEnd(",")
                If Not isBcc Then
                    msg.To.Add(mailTo)       '收件者
                Else
                    msg.Bcc.Add(mailTo)       '密件副本
                End If

                'msg.CC.Add()                        '副本

                Dim sender As MailAddress = New MailAddress(fromMail, fromName, System.Text.Encoding.UTF8)
                msg.From = sender       '寄件者
                msg.Sender = sender     '寄件者
                msg.Subject = subject       '標題主旨
                msg.SubjectEncoding = System.Text.Encoding.UTF8     '標題主旨編碼
                msg.BodyEncoding = System.Text.Encoding.UTF8        '信件內容編碼            
                msg.Priority = MailPriority.Normal      '優先等級
            Catch ex As Exception
                result = -1
                msg.Dispose()
            End Try

            If result <> -1 Then
                '加入HTML BODY內的圖檔路徑
                Dim alView As AlternateView = AlternateView.CreateAlternateViewFromString(content, Nothing, "text/html")        '建立電子郵件訊息格式物件
                Dim appSets As NameValueCollection = ConfigurationManager.AppSettings()     '取得WebConfig內AppSetting的集合
                Dim appPath As String = Nothing
                'appPath = IIf(isThread, HostingEnvironment.MapPath("~\"), Current.Server.MapPath("~\"))        '無法使用此IF寫法？

                If isThread Then                                '判斷是否為執行緒狀態下呼叫此函式區分取得應用程式實體根目錄的方法不同
                    appPath = HostingEnvironment.MapPath("~\")
                Else
                    appPath = Current.Server.MapPath("~\")
                End If

                For Each key In appSets.Keys()
                    If content.Contains(key) Then                           '判斷傳入的HTML內容內是否有與WebConfig相對應的Key
                        Dim path As String = appPath & appSets(key)                    '取得圖檔路徑
                        Dim imgSrc As LinkedResource = New LinkedResource(path)        '建立圖檔路徑資源檔
                        imgSrc.ContentId = key                                         '圖檔路徑在HTML內對應的Key
                        imgSrc.TransferEncoding = TransferEncoding.Base64
                        alView.LinkedResources.Add(imgSrc)                             '將圖檔路徑資源檔加入到電子郵件訊息格式物件
                        msg.AlternateViews.Add(alView)                                 '將電子郵件訊息格式物件匯入至郵件訊息主體
                    End If
                Next

                msg.IsBodyHtml = True

                'SMTP
                Dim mySmtp As SmtpClient = New SmtpClient(Me.smtpName, Me.smtpPort)        'Smtp Server and Port

                mySmtp.EnableSsl = Me.sslConnect         '是否使用 SSL 連線 (Gmail需使用SSL)

                mySmtp.UseDefaultCredentials = IIf(Me.sslConnect, False, True)        '若開啟 SSL 驗證時需把預設驗證關閉，否則不用

                mySmtp.Credentials = New NetworkCredential(Me.smtpAcn, Me.smtpPwd)        '存取smtp的帳號密碼

                Try
                    mySmtp.Send(msg)        '發送信件

                    result = 1      '1為成功

                Catch ex As SmtpFailedRecipientsException       '捕捉Smtp的錯誤事件
                    result = -1      '-1為失敗
                Finally
                    msg.Dispose()
                    mySmtp.Dispose()
                End Try
            End If

            Return result
        End Function


        ''' <summary>
        ''' 發送HTML郵件，並將圖檔資料以附件方式送出，成功回傳1
        ''' </summary>
        ''' <param name="toPeople">收件者信箱地址，需用","隔開</param>
        ''' <param name="fromMail">發件者信箱</param>
        ''' <param name="fromName">發件者姓名</param>
        ''' <param name="subject">標題主旨</param>
        ''' <param name="content">信件內容</param>
        ''' <param name="isThread">是否處於多執行緒狀態下呼叫本函式</param>
        ''' <param name="sleepTime">發送失敗系統重發間隔</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Overloads Function SendAttImgMail(ByVal toPeople As String, ByVal fromMail As String, ByVal fromName As String, ByVal subject As String,
                                    ByVal content As String, ByVal isBcc As Boolean, ByVal isThread As Boolean, ByVal sleepTime As Int32) As Int16

            Dim result As Int16 = Nothing
            Dim msg As MailMessage = New MailMessage()      '郵件訊息主體

            Try
                Dim mailTo As String = toPeople.TrimStart(",").TrimEnd(",")
                If Not isBcc Then
                    msg.To.Add(mailTo)       '收件者
                Else
                    msg.Bcc.Add(mailTo)       '密件副本
                End If

                'msg.CC.Add()                        '副本

                Dim sender As MailAddress = New MailAddress(fromMail, fromName, System.Text.Encoding.UTF8)
                msg.From = sender       '寄件者
                msg.Sender = sender     '寄件者
                msg.Subject = subject       '標題主旨
                msg.SubjectEncoding = System.Text.Encoding.UTF8     '標題主旨編碼
                msg.BodyEncoding = System.Text.Encoding.UTF8        '信件內容編碼            
                msg.Priority = MailPriority.Normal      '優先等級
            Catch ex As Exception
                result = -1
                msg.Dispose()
            End Try

            If result <> -1 Then
                '加入HTML BODY內的圖檔路徑
                Dim alView As AlternateView = AlternateView.CreateAlternateViewFromString(content, Nothing, "text/html")        '建立電子郵件訊息格式物件
                Dim appSets As NameValueCollection = ConfigurationManager.AppSettings()     '取得WebConfig內AppSetting的集合
                Dim appPath As String = Nothing
                'appPath = IIf(isThread, HostingEnvironment.MapPath("~\"), Current.Server.MapPath("~\"))        '無法使用此IF寫法？

                If isThread Then                                '判斷是否為執行緒狀態下呼叫此函式區分取得應用程式實體根目錄的方法不同
                    appPath = HostingEnvironment.MapPath("~\")
                Else
                    appPath = Current.Server.MapPath("~\")
                End If

                For Each key In appSets.Keys()
                    If content.Contains(key) Then                           '判斷傳入的HTML內容內是否有與WebConfig相對應的Key
                        Dim path As String = appPath & appSets(key)                    '取得圖檔路徑
                        Dim imgSrc As LinkedResource = New LinkedResource(path)        '建立圖檔路徑資源檔
                        imgSrc.ContentId = key                                         '圖檔路徑在HTML內對應的Key
                        imgSrc.TransferEncoding = TransferEncoding.Base64
                        alView.LinkedResources.Add(imgSrc)                             '將圖檔路徑資源檔加入到電子郵件訊息格式物件
                        msg.AlternateViews.Add(alView)                                 '將電子郵件訊息格式物件匯入至郵件訊息主體
                    End If
                Next

                msg.IsBodyHtml = True

                'SMTP
                Dim mySmtp As SmtpClient = New SmtpClient(Me.smtpName, Me.smtpPort)        'Smtp Server and Port

                mySmtp.EnableSsl = Me.sslConnect         '是否使用 SSL 連線 (Gmail需使用SSL)

                mySmtp.UseDefaultCredentials = IIf(Me.sslConnect, False, True)        '若開啟 SSL 驗證時需把預設驗證關閉，否則不用

                mySmtp.Credentials = New NetworkCredential(Me.smtpAcn, Me.smtpPwd)        '存取smtp的帳號密碼

                Try
                    mySmtp.Send(msg)        '發送信件

                    result = 1      '1為成功

                Catch ex As SmtpFailedRecipientsException       '捕捉Smtp的錯誤事件
                    result = -1      '-1為失敗
                    For i = 0 To ex.InnerExceptions.Length
                        Dim status As SmtpStatusCode = ex.InnerExceptions(i).StatusCode()       '建立SmtpStatusCode物件取得Smtp代碼
                        If status = SmtpStatusCode.MailboxBusy Or status = SmtpStatusCode.MailboxUnavailable Then       '判斷是否為Smtp忙碌或者是無法取存取的信件
                            Thread.Sleep(sleepTime)      '暫停指定時間後重發
                            mySmtp.Send(msg)        '發送信件
                        End If
                    Next
                Finally
                    msg.Dispose()
                    mySmtp.Dispose()
                End Try
            End If

            Return result
        End Function

    End Class

End Namespace
