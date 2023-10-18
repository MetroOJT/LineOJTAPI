Imports Newtonsoft.Json
Imports System.Net


Partial Class API_GetContent_Index
    Inherits System.Web.UI.Page
    Public Class Requestmessage
        Public Property replyToken As String
        Public Property messages As New List(Of Object)
        Public Sub add_message(obj As Object)
            messages.Add(obj)
        End Sub
    End Class
    Public Class Messages
        Public Property type As String
        Public Property text As String
    End Class
    Public Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
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
        Dim HasEventID_flg As Boolean = False
        Dim req As System.Net.WebRequest =
            System.Net.WebRequest.Create("https://api.line.me/v2/bot/message/reply")
        Dim requestmessage As Requestmessage = New Requestmessage
        Const AccessToken As String = "p/w9bCbcf1BCehRo4gVCvfz7mG+sCLiUF3AeLm3kzrL8ugjKA0JsRYMx/WDnoqNzQhqLjbFXX9QFn/mBQr5wpC9Nrd7uDVjtBGVLDGqlIsbUh+ycI9zhl1rw/UJE6BUawPXWJZ4VMLRk/ItsQkKA3QdB04t89/1O/w1cDnyilFU="
        Try
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol Or SecurityProtocolType.Tls12
            len = Request.ContentLength
            bData = Request.BinaryRead(len)
            sPostData = enc.GetString(bData)
            sPostData = HttpUtility.UrlDecode(sPostData)
            left = sPostData.IndexOf("{")
            right = sPostData.LastIndexOf("}")
            sjson = sPostData.Substring(left, right - left + 1)
            cCom.CmnWriteStepLog(sjson)
            jsonObj = JsonConvert.DeserializeObject(sjson)
            eventsObj = jsonObj("events")(0)
            messageObj = eventsObj("message")
            sourceObj = eventsObj("source")
            Line_UserID = sourceObj("userId").ToString()
            Keyword = messageObj("text").ToString()
            ReplyToken = eventsObj("replyToken").ToString()

            cDB.AddWithValue("@Line_UserID", Line_UserID)
            cDB.AddWithValue("@Keyword", Keyword)
            cDB.AddWithValue("@ReplyToken", ReplyToken)

            sSQL.Clear()
            sSQL.Append(" SELECT")
            sSQL.Append("  EventID")
            sSQL.Append(" FROM " & cCom.gctbl_EventMst)
            sSQL.Append(" WHERE Keyword = @Keyword AND")
            sSQL.Append(" Status = 1 AND")
            sSQL.Append(" ScheduleFm <= NOW() And ScheduleTo >= NOW()")
            cDB.SelectSQL(sSQL.ToString)
            If cDB.ReadDr Then
                HasEventID_flg = True
                cDB.AddWithValue("@EventID", cDB.DRData("EventID"))
            End If

            requestmessage.replyToken = ReplyToken

            sSQL.Clear()
            sSQL.Append(" SELECT")
            sSQL.Append("  Line_UserID")
            sSQL.Append(" ,Keyword")
            sSQL.Append(" FROM " & cCom.gctbl_UsedKeyword)
            sSQL.Append(" INNER JOIN " & cCom.gctbl_EventMst)
            sSQL.Append(" ON " & cCom.gctbl_UsedKeyword & ".EventID =  " & cCom.gctbl_EventMst & ".EventID")
            sSQL.Append(" WHERE Line_UserID = @Line_UserID AND Keyword = @Keyword")
            cDB.SelectSQL(sSQL.ToString)

            If cDB.ReadDr Then
                sSQL.Clear()
                sSQL.Append(" SELECT")
                sSQL.Append("  Message")
                sSQL.Append(" FROM " & cCom.gctbl_EventMst)
                sSQL.Append(" INNER JOIN " & cCom.gctbl_MessageMst)
                sSQL.Append(" ON " & cCom.gctbl_EventMst & ".EventID = " & cCom.gctbl_MessageMst & ".EventID")
                sSQL.Append(" WHERE " & cCom.gctbl_EventMst & ".EventID = 0 AND MessageID = 2")
                cDB.SelectSQL(sSQL.ToString)
                If cDB.ReadDr Then
                    Dim messages As Messages = New Messages
                    messages.type = "text"
                    messages.text = cDB.DRData("Message")
                    requestmessage.add_message(messages)
                End If
            Else
                sSQL.Clear()
                sSQL.Append(" SELECT")
                sSQL.Append("  EventID")
                sSQL.Append(" FROM " & cCom.gctbl_EventMst)
                sSQL.Append(" WHERE Keyword = @Keyword AND")
                sSQL.Append(" Status = 1 AND")
                sSQL.Append(" ScheduleFm <= NOW() And ScheduleTo >= NOW()")
                cDB.SelectSQL(sSQL.ToString)
                If cDB.ReadDr Then
                    sSQL.Clear()
                    sSQL.Append(" INSERT INTO " & cCom.gctbl_UsedKeyword)
                    sSQL.Append(" VALUES(@Line_UserID, @EventID, @ReplyToken, 0)")
                    cDB.ExecuteSQL(sSQL.ToString)
                    sSQL.Clear()
                    sSQL.Append(" SELECT Message")
                    sSQL.Append(" FROM " & cCom.gctbl_MessageMst)
                    sSQL.Append(" WHERE EventID = @EventID")
                    sSQL.Append(" ORDER BY MessageID")
                    cDB.SelectSQL(sSQL.ToString)
                    Do Until Not cDB.ReadDr
                        Dim messages As Messages = New Messages
                        messages.type = "text"
                        messages.text = cDB.DRData("Message")
                        requestmessage.add_message(messages)
                    Loop
                Else
                    sSQL.Clear()
                    sSQL.Append(" SELECT")
                    sSQL.Append("  Message")
                    sSQL.Append(" FROM " & cCom.gctbl_EventMst)
                    sSQL.Append(" INNER JOIN " & cCom.gctbl_MessageMst)
                    sSQL.Append(" ON " & cCom.gctbl_EventMst & ".EventID = " & cCom.gctbl_MessageMst & ".EventID")
                    sSQL.Append(" WHERE " & cCom.gctbl_EventMst & ".EventID = 0 AND MessageID = 1")
                    cDB.SelectSQL(sSQL.ToString)
                    If cDB.ReadDr Then
                        Dim messages As Messages = New Messages
                        messages.type = "text"
                        messages.text = cDB.DRData("Message")
                        requestmessage.add_message(messages)
                    End If
                End If
            End If
            Try
                cDB.AddWithValue("@SendRecv", "send")
                sSQL.Clear()
                sSQL.Append(" INSERT INTO " & cCom.gctbl_LogMst)
                sSQL.Append(" VALUES(@ReplyToken, @Line_UserID, @SendRecv, 999, 'Log', NOW())")
                cDB.ExecuteSQL(sSQL.ToString)
                'POST送信する文字列を作成
                Dim postData As String = JsonConvert.SerializeObject(requestmessage)
                Response.Write(postData)
                'バイト型配列に変換
                Dim postDataBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(postData)
                req.Method = "POST"
                req.ContentType = "application/json"
                'POST送信するデータの長さを指定
                req.ContentLength = postDataBytes.Length
                req.Headers.Add("Authorization", "Bearer " & AccessToken)
                'データをPOST送信するためのStreamを取得
                Dim reqStream As System.IO.Stream = req.GetRequestStream()
                '送信するデータを書き込む
                reqStream.Write(postDataBytes, 0, postDataBytes.Length)
                reqStream.Close()

                'サーバーからの応答を受信するためのWebResponseを取得
                Dim res As System.Net.HttpWebResponse = req.GetResponse()
                '応答データを受信するためのStreamを取得
                Dim resStream As System.IO.Stream = res.GetResponseStream()
                '受信して表示
                Dim sr As New System.IO.StreamReader(resStream, enc)
                Dim num As Integer = res.StatusCode
                Response.Write(num)
                cDB.AddWithValue("@Log", sr.ReadToEnd())
                cDB.AddWithValue("@Status", num)
                sSQL.Clear()
                sSQL.Append(" UPDATE " & cCom.gctbl_LogMst)
                sSQL.Append(" SET Status = @Status, Log = @Log")
                sSQL.Append(" WHERE ReplyToken = @ReplyToken")
                cDB.ExecuteSQL(sSQL.ToString)
                If HasEventID_flg Then
                    sSQL.Clear()
                    sSQL.Append(" UPDATE " & cCom.gctbl_UsedKeyword)
                    sSQL.Append(" SET Reply_flg = b'1'")
                    sSQL.Append(" WHERE Line_UserID = @Line_UserID AND EventID = @EventID")
                    cDB.ExecuteSQL(sSQL.ToString)
                End If

                '閉じる
                sr.Close()
            Catch ex As WebException
                'サーバーからの応答を受信するためのWebResponseを取得
                Dim res As System.Net.HttpWebResponse = ex.Response
                '応答データを受信するためのStreamを取得
                Dim resStream As System.IO.Stream = res.GetResponseStream()
                '受信して表示
                Dim sr As New System.IO.StreamReader(resStream, enc)
                Dim num As Integer = res.StatusCode
                'sRet = ex.Message
                'cDB.AddWithValue("@Log", sjson & "|" & JsonConvert.SerializeObject(requestmessage))
                Dim dumy As String = sr.ReadToEnd()
                cDB.AddWithValue("@Log1", dumy)
                cDB.AddWithValue("@Status", num)
                sSQL.Clear()
                sSQL.Append(" UPDATE " & cCom.gctbl_LogMst)
                sSQL.Append(" SET Log = @Log1, Status = @Status")
                sSQL.Append(" WHERE ReplyToken = @ReplyToken")
                cDB.ExecuteSQL(sSQL.ToString)
            End Try
        Catch ex As Exception
            sRet = ex.Message
            Response.Write(sRet)
            'cDB.AddWithValue("@Log", sjson & "|" & JsonConvert.SerializeObject(requestmessage))
            cDB.AddWithValue("@Log2", sRet)
            sSQL.Clear()
            sSQL.Append(" UPDATE " & cCom.gctbl_LogMst)
            sSQL.Append(" SET Log = @Log2")
            sSQL.Append(" WHERE ReplyToken = @ReplyToken")
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
'reqStream.Writeに文字列だけを入れる