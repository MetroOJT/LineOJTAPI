Imports Newtonsoft.Json
Imports System.Net


Partial Class API_GetContent_Index
    Inherits System.Web.UI.Page

    'LINEに送るJSONオブジェクトのためのクラス
    Public Class Requestmessage
        Public Property replyToken As String
        Public Property messages As New List(Of Object)
        Public Sub add_message(str As String)
            Dim message As Messages = New Messages
            message.type = "text"
            message.text = str
            messages.Add(message)
        End Sub
    End Class

    'Requestmessage.messagesに追加するJSONオブジェクトのクラス
    Public Class Messages
        Public Property type As String
        Public Property text As String
    End Class

    '読み込む前にTLS1.2を適用する
    Public Sub Pre_Load(sender As Object, e As EventArgs) Handles Me.PreLoad
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12
    End Sub

    '読み込み後
    Public Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

        '変数宣言
        Dim cCom As New Common
        Dim cDB As New CommonDB
        Dim Cki As New Cookie
        Dim sSQL As New StringBuilder
        Dim sRet As String = ""
        Dim len As Integer = 0
        Dim bData() As Byte
        Dim enc As Encoding = Encoding.GetEncoding("UTF-8")
        Dim sPostData As String = ""
        Dim left As Integer = 0
        Dim right As Integer = 0
        Dim sjson As String = ""
        Dim jsonObj As Object = Nothing
        Dim eventsObj As Object = Nothing
        Dim messageObj As Object = Nothing
        Dim sourceObj As Object = Nothing
        Dim Line_UserID As String = ""
        Dim Keyword As String = ""
        Dim ReplyToken As String = ""
        Dim CouponCode As String = ""
        Dim HasEventID_flg As Boolean = False

        'LINE返信用のWebRequest
        Dim req As System.Net.WebRequest =
            System.Net.WebRequest.Create("https://api.line.me/v2/bot/message/reply")

        'Requestmessageのインスタンス化
        Dim requestmessage As Requestmessage = New Requestmessage

        '公式アカウントアクセス用のトークン
        Const AccessToken As String = "p/w9bCbcf1BCehRo4gVCvfz7mG+sCLiUF3AeLm3kzrL8ugjKA0JsRYMx/WDnoqNzQhqLjbFXX9QFn/mBQr5wpC9Nrd7uDVjtBGVLDGqlIsbUh+ycI9zhl1rw/UJE6BUawPXWJZ4VMLRk/ItsQkKA3QdB04t89/1O/w1cDnyilFU="
        Try

            'Requestのデータを読み込む
            len = Request.ContentLength
            bData = Request.BinaryRead(len)
            sPostData = enc.GetString(bData)
            sPostData = HttpUtility.UrlDecode(sPostData)

            'JSONデータに変換
            left = sPostData.IndexOf("{")
            right = sPostData.LastIndexOf("}")
            sjson = sPostData.Substring(left, right - left + 1)

            'Object型に変えて必要なデータを取り出す
            jsonObj = JsonConvert.DeserializeObject(sjson)
            eventsObj = jsonObj("events")(0)
            messageObj = eventsObj("message")
            sourceObj = eventsObj("source")
            Line_UserID = sourceObj("userId").ToString()
            Keyword = messageObj("text").ToString()
            ReplyToken = eventsObj("replyToken").ToString()

            '変数をSQLで扱えるようにする
            cDB.AddWithValue("@Line_UserID", Line_UserID)
            cDB.AddWithValue("@Keyword", Keyword)
            cDB.AddWithValue("@ReplyToken", ReplyToken)
            cDB.AddWithValue("@Send", "Send")
            cDB.AddWithValue("@Recv", "Recv")
            cDB.AddWithValue("@RecvLog", sjson)

            'Recvlog追加
            sSQL.Clear()
            sSQL.Append(" INSERT INTO " & cCom.gctbl_LogMst)
            sSQL.Append(" (ReplyToken, SendRecv, Line_UserID, Status, Log, Datetime)")
            sSQL.Append(" VALUES(@ReplyToken, @Recv, @Line_UserID, 200, @RecvLog, NOW())")
            cDB.ExecuteSQL(sSQL.ToString)

            'ReplyTokenをrequestmessageに設定
            requestmessage.replyToken = ReplyToken

            'Line_UserIDが登録済みか確認
            sSQL.Clear()
            sSQL.Append(" SELECT")
            sSQL.Append(" *")
            sSQL.Append(" FROM " & cCom.gctbl_LineUserMst)
            sSQL.Append(" WHERE Line_UserID = @Line_UserID")
            cDB.SelectSQL(sSQL.ToString)
            '未登録の場合挿入
            If Not cDB.IsSelectExistRecord() Then
                sSQL.Clear()
                sSQL.Append(" INSERT INTO " & cCom.gctbl_LineUserMst)
                sSQL.Append(" (Line_UserID, Insert_Date)")
                sSQL.Append(" VALUES (@Line_UserID, NOW())")
                cDB.ExecuteSQL(sSQL.ToString)
            End If

            'クーポンモード取得
            Dim iCouponMode As Integer = 0
            sSQL.Clear()
            sSQL.Append(" SELECT")
            sSQL.Append(" CouponMode")
            sSQL.Append(" FROM " & cCom.gctbl_LineUserMst)
            sSQL.Append(" WHERE Line_UserID = @Line_UserID")
            cDB.SelectSQL(sSQL.ToString)
            If cDB.ReadDr Then
                iCouponMode = Integer.Parse(cDB.DRData("CouponMode").ToString)
            End If

            'クーポンコードボタン押下
            If (Keyword = "クーポンコード") Then
                'クーポンモードON
                iCouponMode = 1
                requestmessage.add_message("クーポンコードを入力してください")
                sSQL.Clear()
                sSQL.Append(" UPDATE " & cCom.gctbl_LineUserMst)
                sSQL.Append(" SET CouponMode = b'" & iCouponMode & "'")
                sSQL.Append(" WHERE Line_UserID = @Line_UserID")
                cDB.ExecuteSQL(sSQL.ToString)
            ElseIf (iCouponMode = 1) Then
                'クーポンモードOFF
                iCouponMode = 0
                sSQL.Clear()
                sSQL.Append(" UPDATE " & cCom.gctbl_LineUserMst)
                sSQL.Append(" SET CouponMode = b'" & iCouponMode & "'")
                sSQL.Append(" WHERE Line_UserID = @Line_UserID")
                cDB.ExecuteSQL(sSQL.ToString)

                '送られてきたキーワードが有効か確認
                sSQL.Clear()
                sSQL.Append(" SELECT")
                sSQL.Append("  EventID")
                sSQL.Append(" FROM " & cCom.gctbl_EventMst)
                sSQL.Append(" WHERE Keyword = @Keyword AND")
                sSQL.Append(" Status = 1 AND")
                sSQL.Append(" ScheduleFm <= NOW() And ScheduleTo >= NOW()")
                cDB.SelectSQL(sSQL.ToString)

                'キーワード有効
                If cDB.ReadDr Then
                    HasEventID_flg = True
                    cDB.AddWithValue("@EventID", cDB.DRData("EventID"))

                    'キーワードが使用済みか確認
                    sSQL.Clear()
                    sSQL.Append(" SELECT")
                    sSQL.Append("  Line_UserID")
                    sSQL.Append(" ,Keyword")
                    sSQL.Append(" FROM " & cCom.gctbl_UsedKeyword)
                    sSQL.Append(" INNER JOIN " & cCom.gctbl_EventMst)
                    sSQL.Append(" ON " & cCom.gctbl_UsedKeyword & ".EventID =  " & cCom.gctbl_EventMst & ".EventID")
                    sSQL.Append(" WHERE Line_UserID = @Line_UserID AND Keyword = @Keyword")
                    cDB.SelectSQL(sSQL.ToString)

                    'キーワード使用済み
                    If cDB.ReadDr Then

                        '使用済みメッセージを取得
                        sSQL.Clear()
                        sSQL.Append(" SELECT")
                        sSQL.Append("  Message")
                        sSQL.Append(" FROM " & cCom.gctbl_EventMst)
                        sSQL.Append(" INNER JOIN " & cCom.gctbl_MessageMst)
                        sSQL.Append(" ON " & cCom.gctbl_EventMst & ".EventID = " & cCom.gctbl_MessageMst & ".EventID")
                        sSQL.Append(" WHERE " & cCom.gctbl_EventMst & ".EventID = 0 AND MessageID = 2")
                        cDB.SelectSQL(sSQL.ToString)
                        If cDB.ReadDr Then
                            requestmessage.add_message(cDB.DRData("Message"))
                        End If

                        'キーワード未使用
                    Else

                        'メッセージ取得
                        sSQL.Clear()
                        sSQL.Append(" SELECT Message")
                        sSQL.Append(" FROM " & cCom.gctbl_MessageMst)
                        sSQL.Append(" WHERE EventID = @EventID")
                        sSQL.Append(" ORDER BY MessageID")
                        cDB.SelectSQL(sSQL.ToString)

                        'メッセージリストを作成
                        Dim message_Array As New ArrayList()
                        Do Until Not cDB.ReadDr
                            message_Array.Add(cDB.DRData("Message"))
                        Loop

                        'クーポンコードの生成を一回に制限する
                        Dim Generated_CouponCode As Boolean = False

                        'クーポンコードの生成とrequestmessageへの追加
                        For Each message In message_Array
                            If Not Generated_CouponCode And message.IndexOf(cCom.gcFormatCouponCode) >= 0 Then
                                Dim count As Integer = 0
                                While True
                                    count += 1

                                    'クーポンコードの生成
                                    CouponCode = cCom.CmnGenerateAlphaNumeric(10)
                                    cDB.AddWithValue("@CouponCode" & count, CouponCode)

                                    'クーポンコードが使用済みか確認
                                    sSQL.Clear()
                                    sSQL.Append(" SELECT CouponCode")
                                    sSQL.Append(" FROM " & cCom.gctbl_UsedKeyword)
                                    sSQL.Append(" WHERE CouponCode = @CouponCode" & count)
                                    cDB.SelectSQL(sSQL.ToString)

                                    '未使用の場合Exit
                                    If Not cDB.IsSelectExistRecord() Then
                                        Exit While
                                    End If
                                End While
                                Generated_CouponCode = True
                            End If

                            'クーポンコードに置き換え
                            message = message.Replace(cCom.gcFormatCouponCode, CouponCode)
                            requestmessage.add_message(message)
                        Next

                        'Used_Keywordに追加
                        cDB.AddWithValue("@CouponCode", CouponCode)
                        sSQL.Clear()
                        sSQL.Append(" INSERT INTO " & cCom.gctbl_UsedKeyword)
                        sSQL.Append(" VALUES(@Line_UserID, @EventID, @ReplyToken, 0, @CouponCode, 1)")
                        cDB.ExecuteSQL(sSQL.ToString)
                    End If

                    'キーワード無効
                Else

                    'キーワード無効メッセージを取得
                    sSQL.Clear()
                    sSQL.Append(" SELECT")
                    sSQL.Append("  Message")
                    sSQL.Append(" FROM " & cCom.gctbl_EventMst)
                    sSQL.Append(" INNER JOIN " & cCom.gctbl_MessageMst)
                    sSQL.Append(" ON " & cCom.gctbl_EventMst & ".EventID = " & cCom.gctbl_MessageMst & ".EventID")
                    sSQL.Append(" WHERE " & cCom.gctbl_EventMst & ".EventID = 0 AND MessageID = 1")
                    cDB.SelectSQL(sSQL.ToString)
                    If cDB.ReadDr Then
                        requestmessage.add_message(cDB.DRData("Message"))
                    End If
                End If
            End If

            'POST送信
            Try
                If requestmessage.messages.Count >= 1 Then
                    '仮送信ログを登録
                    sSQL.Clear()
                    sSQL.Append(" INSERT INTO " & cCom.gctbl_LogMst)
                    sSQL.Append(" (ReplyToken, SendRecv, Line_UserID, Status, Log, Datetime)")
                    sSQL.Append(" VALUES(@ReplyToken, @Send, @Line_UserID, 999, 'Log', NOW())")
                    cDB.ExecuteSQL(sSQL.ToString)
                    'POST送信するデータを作成
                    Dim postData As String = JsonConvert.SerializeObject(requestmessage)
                    req.Method = "POST"
                    req.ContentType = "application/json"
                    req.Headers.Add("Authorization", "Bearer " & AccessToken)
                    Using reqStream As New System.IO.StreamWriter(req.GetRequestStream())
                        'POST送信
                        reqStream.Write(postData)
                    End Using
                    'サーバーからの応答を受信するためのWebResponseを取得
                    Dim res As System.Net.HttpWebResponse = req.GetResponse()
                    '応答データを受信するためのStreamを取得
                    Dim resStream As System.IO.Stream = res.GetResponseStream()
                    '受信して表示
                    Dim sr As New System.IO.StreamReader(resStream, enc)
                    Dim statuscode As Integer = res.StatusCode
                    cDB.AddWithValue("@SendLog", postData)
                    cDB.AddWithValue("@Status", statuscode)

                    '送信ログを更新
                    sSQL.Clear()
                    sSQL.Append(" UPDATE " & cCom.gctbl_LogMst)
                    sSQL.Append(" SET Status = @Status, Log = @SendLog")
                    sSQL.Append(" WHERE ReplyToken = @ReplyToken")
                    sSQL.Append(" AND SendRecv = @Send")
                    cDB.ExecuteSQL(sSQL.ToString)

                    '閉じる
                    sr.Close()
                End If

                '最後のログIDを取得
                sSQL.Clear()
                sSQL.Append(" SELECT")
                sSQL.Append(" MAX(LogID) AS Last_LogID")
                sSQL.Append(" FROM " & cCom.gctbl_LogMst)
                sSQL.Append(" WHERE Line_UserID = @Line_UserID")
                sSQL.Append("   AND Status = 200")
                cDB.SelectSQL(sSQL.ToString)
                If cDB.ReadDr Then
                    cDB.AddWithValue("@Last_LogID", cDB.DRData("Last_LogID"))
                End If

                '最後のログIDを更新
                sSQL.Clear()
                sSQL.Append(" UPDATE " & cCom.gctbl_LineUserMst)
                sSQL.Append(" SET Last_LogID = @Last_LogID")
                sSQL.Append(" WHERE Line_UserID = @Line_UserID")
                cDB.ExecuteSQL(sSQL.ToString)

                'Reply_flgを1(送信済み)にする
                If HasEventID_flg Then
                    sSQL.Clear()
                    sSQL.Append(" UPDATE " & cCom.gctbl_UsedKeyword)
                    sSQL.Append(" SET Reply_flg = b'1'")
                    sSQL.Append(" WHERE Line_UserID = @Line_UserID AND EventID = @EventID")
                    cDB.ExecuteSQL(sSQL.ToString)
                End If
            Catch ex As WebException

                sRet = ex.Message
                If sRet <> "" Then
                    cCom.CmnWriteStepLog(sRet)
                End If

                'エラーの応答を受信するためのWebResponseを取得
                Dim res As System.Net.HttpWebResponse = ex.Response

                '応答データを受信するためのStreamを取得
                Dim resStream As System.IO.Stream = res.GetResponseStream()

                '受信して表示
                Dim sr As New System.IO.StreamReader(resStream, enc)
                Dim statuscode As Integer = res.StatusCode
                cDB.AddWithValue("@SendLog1", sr.ReadToEnd())
                cDB.AddWithValue("@Status", statuscode)

                '送信ログ更新
                sSQL.Clear()
                sSQL.Append(" UPDATE " & cCom.gctbl_LogMst)
                sSQL.Append(" SET Log = @SendLog1, Status = @Status")
                sSQL.Append(" WHERE ReplyToken = @ReplyToken")
                sSQL.Append(" AND SendRecv = @Send")
                cDB.ExecuteSQL(sSQL.ToString)
                sr.Close()
            End Try
        Catch ex As Exception

            sRet = ex.Message
            If sRet <> "" Then
                cCom.CmnWriteStepLog(sRet)
            End If

            cDB.AddWithValue("@SendLog2", sRet)

            '送信ログ更新
            sSQL.Clear()
            sSQL.Append(" UPDATE " & cCom.gctbl_LogMst)
            sSQL.Append(" SET Log = @SendLog2")
            sSQL.Append(" WHERE ReplyToken = @ReplyToken")
            sSQL.Append(" AND SendRecv = @Send")
            cDB.ExecuteSQL(sSQL.ToString)
        Finally
            cDB.DrClose()
            cDB.Dispose()
            If sRet <> "" Then
                cCom.CmnWriteStepLog(sRet)
            End If
        End Try
    End Sub
End Class