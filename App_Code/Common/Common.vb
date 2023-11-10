Option Explicit On
Option Strict On

Imports Microsoft.VisualBasic
Imports MySql.Data.MySqlClient  'DataBase接続用
Imports System.Collections.Generic

Public Class Common

    Inherits System.Web.UI.Page

    Public gcDate_DefaultDate As Date = #1/1/1900#
    Public gcDate_EndDate As Date = #12/31/2099#
    Public gcFormatCouponCode As String = "{CouponCode}"

#Region "Table"
    Public gctbl_UserMst As String = ""
    Public gctbl_EventMst As String = ""
    Public gctbl_MenuMst As String = ""
    Public gctbl_MessageMst As String = ""
    Public gctbl_UsedKeyword As String = ""
    Public gctbl_LogMst As String = ""
    Public gctbl_LineUserMst As String = ""


#End Region

#Region "コンストラクタ"
    ''' <summary>
    ''' Common.vbが呼び出されると初めに処理する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        CmnDataBaseNameChange()
    End Sub
#End Region

#Region "Cmn**"
    ''' <summary>
    ''' セッション切れのチェック(Login画面に戻る場合は引数にTrue(デフォルトTrue))
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CmnLoginCheck() As String
        Dim sRet As String = ""
        Dim sAttKey As String = ""
        Dim sLogin As String = "../../Login/index.aspx"
        Dim sErr As String = ""

        sAttKey = getAttKey()

        If sAttKey = "" Then
            sErr = "セッションの有効期限が切れています。<br>再ログインしてください。"
        ElseIf Not Left(sAttKey, 1) = "K" Then
            sErr = "認証が不正です。再ログインしてください。"
        End If

        If sErr <> "" Then
            HttpContext.Current.Response.Redirect(sLogin)
        End If

        Return sRet

    End Function

    ''' <summary>
    ''' Web.configの設定値を取得
    ''' </summary>
    ''' <param name="sName">パラメータ名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CmnGet_AppSetting(ByVal sName As String) As String

        Dim sRet As String = ""

        Try
            sRet = System.Configuration.ConfigurationManager.AppSettings(sName).ToString
        Catch ex As Exception

        End Try

        Return sRet

    End Function

    ''' <summary>
    ''' データベース名を取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CmnDataBaseNameChange() As Boolean

        Dim scfg_LineMsg As String = ""

        Try
            scfg_LineMsg = CmnGet_AppSetting("LineOJTDB")
            scfg_LineMsg = scfg_LineMsg & "."

            gctbl_UserMst = scfg_LineMsg & "usermst"
            gctbl_EventMst = scfg_LineMsg & "eventmst"
            gctbl_MenuMst = scfg_LineMsg & "menumst"
            gctbl_MessageMst = scfg_LineMsg & "messagemst"
            gctbl_UsedKeyword = scfg_LineMsg & "usedkeyword"
            gctbl_LogMst = scfg_LineMsg & "logmst"
            gctbl_LineUserMst = scfg_LineMsg & "lineusermst"


        Catch ex As Exception

            Return False
        End Try

        Return True

    End Function

    ''' <summary>
    ''' アプリ名を取得
    ''' </summary>
    ''' <param name="sAddrWork">URL</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnGet_AppName(Optional ByVal sAddrWork As String = "") As String

        Dim sRet As String
        Dim rdWork() As String
        Dim sAddr As String = ""

        sRet = ""

        If sAddrWork = "" Then
            sAddr = HttpContext.Current.Request.ServerVariables("URL")
        Else
            sAddr = sAddrWork
        End If

        rdWork = Split(sAddr, "/")

        '配列の最後から一つ前の文字列を取り出す様に変更
        sRet = rdWork(UBound(rdWork) - 1)

        Return sRet

    End Function

    ''' <summary>
    ''' ファイル名を取得
    ''' </summary>
    ''' <param name="sAddrWork">URL</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnGet_FileName(Optional ByVal sAddrWork As String = "") As String

        Dim sRet As String
        Dim rdWork() As String
        Dim sAddr As String = ""

        sRet = ""

        If sAddrWork = "" Then
            sAddr = HttpContext.Current.Request.ServerVariables("URL")
        Else
            sAddr = sAddrWork
        End If

        rdWork = Split(sAddr, "/")

        '配列の最後の文字列を取り出す
        sRet = rdWork(UBound(rdWork))

        rdWork = Split(sRet, ".")

        '配列の最初の文字列を取り出す
        sRet = rdWork(0)

        Return sRet

    End Function

    ''' <summary>
    ''' テンプテーブル名を取得
    ''' </summary>
    ''' <param name="sSubName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnGet_TableName(Optional ByVal sSubName As String = "") As String

        Dim sName As String = ""
        Dim sRet As String = ""
        Dim sAtt As String = ""
        Dim sAppName As String = ""


        sAtt = getAttKey()
        sAppName = CmnGet_AppName()

        sRet = ""
        sName = sAppName & "_" & sAtt

        If Not sSubName = "" Then sName = sName & "_" & sSubName

        sRet = CmnGet_AppSetting("LineOJTDB") & "." & "TMP_" & Replace(sName, "-", "")

        Return sRet

    End Function

    ''' <summary>
    ''' SqlCommandからParametersを取得し、文字列にして返します。
    ''' </summary>
    ''' <param name="cmd">SqlCommand</param>
    ''' <returns>Name=Value,Name=...に変換した文字列</returns>
    ''' <remarks></remarks>
    Public Function CmnGet_SqlParameterString(ByVal cmd As MySqlCommand) As String
        If cmd Is Nothing Then
            Return "SqlCommandがNothingです。"
        End If

        Dim parameterList As New List(Of String)
        For Each itm As MySqlParameter In cmd.Parameters
            If itm.Value Is Nothing Then
                parameterList.Add(itm.ParameterName & " = Error:ValueがNothingです。")
            Else
                parameterList.Add(itm.ParameterName & " = " & itm.Value.ToString())
            End If
        Next
        Return Join(parameterList.ToArray(), ",")
    End Function

    ''' <summary>
    ''' ログ保存
    ''' </summary>
    ''' <param name="sComment">ログ内容</param>
    ''' <param name="sAppName">アプリ名※基本未指定</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnWriteStepLog(ByVal sComment As String, Optional ByVal sAppName As String = "", Optional sDate As String = "") As String

        Dim sRet As String = ""
        Dim sName As String = ""
        Dim iFile As Integer = 0
        Dim sPath As String = ""
        Dim lLen As Long = 0
        Dim sDir As String = ""
        Dim LogPath As String = ""

        Try
            sName = CmnGet_FileName()

            If sDate <> "" Then
                sName = sName & "_" & sDate
            End If

            sAppName = CmnGet_AppName()

            LogPath = MakePath(System.Configuration.ConfigurationManager.AppSettings("SystemFolder").ToString, "Log")

            sDir = LogPath

            If Not sAppName = "" Then
                sDir = MakePath(sDir, CStr(sAppName))
            End If

            If Dir$(sDir, vbDirectory) = "" Then MkDir(sDir)

            sPath = MakePath(sDir, sName & ".log")
            If Not Dir$(sPath) = "" Then
                lLen = FileLen(sPath)
                If lLen > 0 Then
                    If lLen / 5120 > 5120 Then
                        FileCopy(sPath, MakePath(sDir, sName & ".bak"))
                        Kill(sPath)
                    End If
                End If

            End If

            iFile = FreeFile()
            If Not Dir$(sPath) = "" Then
                FileOpen(iFile, sPath, OpenMode.Append)
            Else
                FileOpen(iFile, sPath, OpenMode.Output)
            End If

            PrintLine(iFile, Format(Now(), "yyyy/MM/dd HH:mm:ss") & " " & sComment)
            FileClose(iFile)

        Catch ex As Exception
            sRet = ex.Message & "-WriteStepLog-"

        End Try

        Return sRet

    End Function

    ''' <summary>
    ''' パスを生成する（渡された２つのＳｔｒを、「Ｐａｔｈ￥ＦＮａｍｅ」の形にする
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function MakePath(ByRef PATH As String, ByRef FName As String) As String

        Dim wPath As String
        Dim wFName As String

        wPath = PATH
        If Right(wPath, 1) <> "\" Then
            wPath = wPath & "\"
        End If

        wFName = FName
        If Left(wFName, 1) = "\" Then
            wFName = Mid(wFName, 2, Len(FName) - 1)
        End If

        MakePath = wPath & wFName

    End Function

    ''' <summary>
    ''' デフォルトの開始日を返す
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnGet_DefaultDate() As Date
        Return gcDate_DefaultDate
    End Function

    ''' <summary>
    ''' デフォルトの終了日を返す
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnGet_EndDate() As Date
        Return gcDate_EndDate
    End Function

    ''' <summary>
    ''' インクルードするJavaScriptのタグを作成する
    ''' </summary>
    ''' <param name="sName">JavaScriptファイルのパス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnMakeInc_Java(ByVal sName As String) As String

        Dim sRet As New StringBuilder
        Dim sSet_Name As String = ""

        sSet_Name = sName
        If UCase(Right(sName, 3)) = ".JS" Then
            sSet_Name = sSet_Name & "?dmy=" & Format(Now, "yyMMddhhmmss")
        End If

        sRet.Clear()
        sRet.Append("<script src=""" & sSet_Name & """></script>")

        Return sRet.ToString

    End Function

    ''' <summary>
    ''' インクルードするCSSのタグを作成する
    ''' </summary>
    ''' <param name="sName">CSSファイルのパス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnMakeInc_CSS(ByVal sName As String) As String

        Dim sRet As New StringBuilder
        Dim sSet_Name As String = ""

        sSet_Name = sName
        If UCase(Right(sName, 4)) = ".CSS" Then
            sSet_Name = sSet_Name & "?dmy=" & Format(Now, "yyMMddhhmmss")
        End If

        sRet.Clear()
        sRet.Append("<link href=""" & sSet_Name & """ rel=""" & "stylesheet"""" />")

        Return sRet.ToString

    End Function

    ''' <summary>
    ''' 共通インクルードファイル（必ず最初に読み込み必要）
    ''' </summary>
    ''' <param name="AddMetaText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CmnInc(Optional ByVal AddMetaText As String = "") As String

        Dim sRet As New StringBuilder
        Dim sCSS As String = "../../Common/"
        Dim sJS As String = "../../Common/"

        sRet.Clear()
        sRet.Append("<script src='https://code.jquery.com/jquery-3.7.1.min.js' integrity='sha256-/JqT3SQfawRcv/BIHPThkBvs0OEvtFFmqPF/lYI/Cxo=' crossorigin='anonymous'></script>" & vbCr)
        sRet.Append(CmnMakeInc_CSS(sCSS & "Common.css") & vbCr)
        sRet.Append(CmnMakeInc_Java(sJS & "Common.js") & vbCr)
        If AddMetaText <> "" Then
            sRet.Append(AddMetaText & vbCr)
        End If

        Return sRet.ToString

    End Function

    ''' <summary>
    ''' 共通フッター（必ず最後に読み込み必要）
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CmnDspFooter() As String
        Dim sRet As New StringBuilder

        sRet.Clear()
        sRet.Append("</div>" & vbCr)
        sRet.Append("<br class=""clear"" />" & vbCr)
        sRet.Append("</div>" & vbCr)
        Return sRet.ToString

    End Function

    ''' <summary>
    ''' デフォルトの開始日または終了日ではない場合、指定したフォーマットに変換して返す
    ''' </summary>
    ''' <param name="dDate">日付</param>
    ''' <param name="sStyle">フォーマット</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnFormatDate(ByVal dDate As Date, ByVal sStyle As String) As String

        Dim sRet As String = ""

        If Not dDate = gcDate_DefaultDate AndAlso Not dDate = gcDate_EndDate Then
            sRet = Format(dDate, sStyle)
        End If

        Return sRet

    End Function

    ''' <summary>
    ''' デフォルトの開始日または終了日ではない場合、指定したフォーマットに変換して返す
    ''' </summary>
    ''' <param name="dDate">日付</param>
    ''' <param name="sStyle">フォーマット</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnFormatDate(ByVal dDate As String, ByVal sStyle As String) As String

        Dim sRet As String = ""

        If Not CDate(dDate) = gcDate_DefaultDate AndAlso Not CDate(dDate) = gcDate_EndDate Then
            sRet = Format(CDate(dDate), sStyle)
        End If

        Return sRet

    End Function

    ''' <summary>
    ''' シングルクォートをエスケープして返す
    ''' </summary>
    ''' <param name="sReplaceStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CmnReplaceSingleQuote(ByVal sReplaceStr As String) As String
        Return sReplaceStr.Replace("'", "''")
    End Function


    ''' <summary>
    ''' ランダムな文字列を生成する
    ''' </summary>
    ''' <param name="length">生成する文字列の長さ</param>
    ''' <returns>生成された文字列</returns>
    Public Function CmnGenerateAlphaNumeric(ByVal length As Integer) As String
        'パスワードに使用する文字
        Const randomChars As String = "0123456789abcdefghijklmnopqrstuvwxyz"
        Dim sb As New System.Text.StringBuilder(length)
        Dim r As New Random()

        For i As Integer = 0 To length - 1
            '文字の位置をランダムに選択
            Dim pos As Integer = r.Next(randomChars.Length)
            '選択された位置の文字を取得
            Dim c As Char = randomChars(pos)
            'パスワードに追加
            sb.Append(c)
        Next

        Return sb.ToString()
    End Function


#End Region

#Region "get**"
    ''' <summary>
    ''' 認証キーを返す
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getAttKey() As String
        Dim sRet As String = ""

        Try

            If Not HttpContext.Current.Request.Cookies("AttKey") Is Nothing Then
                sRet = HttpContext.Current.Request.Cookies("AttKey").Value
            End If

        Catch ex As Exception
            sRet = ""
        Finally

        End Try

        Return sRet

    End Function

    ''' <summary>
    ''' 文字列を指定のバイト数にカットする関数（漢字分断回避あり）
    ''' </summary>
    ''' <param name="strInput">文字列</param>
    ''' <param name="intLen">バイト数</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CutStrByteLen(ByVal strInput As String, ByVal intLen As Integer) As String
        Dim sjis As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        Dim tempLen As Integer = sjis.GetByteCount(strInput)
        Dim bytTemp As Byte() = sjis.GetBytes(strInput)
        Dim strTemp As String = sjis.GetString(bytTemp, 0, intLen)
        ' 末尾の漢字が分断されたらバイト数-1で切り取る (VB2005="・"、.NET2003=NullChar）
        If strTemp.EndsWith(ControlChars.NullChar) OrElse strTemp.EndsWith("・") Then
            strTemp = sjis.GetString(bytTemp, 0, intLen - 1)
        End If
        Return strTemp
    End Function
#End Region

End Class
