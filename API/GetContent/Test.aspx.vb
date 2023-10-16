Imports Newtonsoft.Json

Partial Class API_GetContent_Test
    Inherits System.Web.UI.Page
    Public Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        '文字コードを指定する
        Dim enc As System.Text.Encoding =
            System.Text.Encoding.GetEncoding("UTF-8")
        Dim req As System.Net.WebRequest =
            System.Net.WebRequest.Create("https://api.line.me/v2/bot/message/reply")
        Try
            req.Method = "POST"
            'ContentTypeを"application/x-www-form-urlencoded"にする
            req.ContentType = "application/json"
            'POST送信するデータの長さを指定
            'req.Headers.Add("Authorization", "Bearer " + AccessToken);
            'req.ContentLength = postDataBytes.Length
            'データをPOST送信するためのStreamを取得
            Dim reqStream As System.IO.Stream = req.GetRequestStream()
            '送信するデータを書き込む
            'reqStream.Write(postDataBytes, 0, postDataBytes.Length)
            reqStream.Close()

            'サーバーからの応答を受信するためのWebResponseを取得
            Dim res As System.Net.HttpWebResponse = req.GetResponse()
            '応答データを受信するためのStreamを取得
            Dim resStream As System.IO.Stream = res.GetResponseStream()
            '受信して表示
            Dim sr As New System.IO.StreamReader(resStream, enc)
            Dim num As Integer = res.StatusCode
            Response.Write(num)

            '閉じる
            sr.Close()
        Catch ex As Exception

        Finally

        End Try
    End Sub
End Class
