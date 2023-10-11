Option Explicit On
Option Strict On

Imports Microsoft.VisualBasic
Imports MySql.Data.MySqlClient  'DataBase接続用

Public Class CommonDB
    Implements IDisposable

#Region "メンバ変数"
    Private mConn As MySqlConnection = Nothing
    Private mCmd As MySqlCommand = Nothing
    Private mDr As MySqlDataReader = Nothing
    Private mTran As MySqlTransaction = Nothing
#End Region

#Region "Privateメンバ関数"
    ''' <summary>
    ''' DB接続
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConnectDb()

        Try
            If mConn Is Nothing Then
                mConn = New MySqlConnection
            Else
                If mConn.State = Data.ConnectionState.Open Then
                    Return
                End If
            End If

            If mCmd Is Nothing Then
                mCmd = New MySqlCommand
            End If

            mConn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("DBConn").ToString
            mConn.Open()

            mCmd.Connection = mConn
            mCmd.CommandTimeout = 30

        Catch ex As Exception
            Throw New Exception("DBエラー:SQLサーバの接続に失敗しました。" & ex.Message)

        End Try

    End Sub

#End Region

#Region "Publicプロパティ"
    ''' <summary>
    ''' SqlCommandオブジェクトを返す。通常は使用しないこと！
    ''' </summary>
    Public ReadOnly Property Cmd() As MySqlCommand
        Get
            If mCmd Is Nothing Then
                mCmd = New MySqlCommand
            End If
            Return mCmd
        End Get
    End Property
#End Region

#Region "Publicメンバ関数"
    ''' <summary>
    ''' SELECTのレコードが存在するか?
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsSelectExistRecord() As Boolean

        If Me.mDr Is Nothing Then
            Return False

        End If

        Try
            Return mDr.HasRows

        Catch ex As Exception
            Throw New Exception("DBエラー:" & ex.Message)
            Return False

        End Try

    End Function

    ''' <summary>
    ''' レコードを読み込む
    ''' </summary>
    ''' <remarks></remarks>
    Public Function ReadDr() As Boolean

        If Me.mDr Is Nothing Then
            Return False

        End If

        Try
            Return mDr.Read()

        Catch ex As Exception
            Throw New Exception("DBエラー:" & ex.Message)

        End Try

    End Function

    ''' <summary>
    ''' Drの指定フィールド名の値取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DRData(ByVal vsFieldNm As String) As Object

        If Me.mDr Is Nothing Then
            Return Nothing

        End If

        Try
            Return mDr(vsFieldNm)

        Catch ex As Exception
            Throw New Exception("DBエラー:" & ex.Message)
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' SELECT SQLを実行する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SelectSQL(ByVal vsSQL As String)

        Try
            '接続
            Me.ConnectDb()

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            'SQL実行
            Me.mCmd.CommandText = vsSQL
            Me.mDr = Me.mCmd.ExecuteReader

        Catch ex As Exception
            Me.mDr = Nothing
            Throw New Exception(ex.Message & "SQL実行エラー" & vsSQL)

        End Try

    End Sub

    ''' <summary>
    ''' SQLを実行して最初の行の最初の列を返す（ExecuteScalar()）
    ''' </summary>
    ''' <param name="vsSQL">戻り値を取得するSQL文</param>
    ''' <returns>Object型（型変換が必要）</returns>
    ''' <remarks></remarks>
    Public Function ScalarSQL(ByVal vsSQL As String) As Object

        Try
            '接続
            Me.ConnectDb()

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            'SQL実行
            Me.mCmd.CommandText = vsSQL
            Return Me.mCmd.ExecuteScalar()

        Catch ex As Exception
            Me.mDr = Nothing
            Throw New Exception(ex.Message & "SQL実行エラー" & vsSQL)

        End Try

    End Function

    ''' <summary>
    ''' SQLの実行(SELECT以外)
    ''' </summary>
    ''' <param name="vsSQL"></param>
    ''' <remarks></remarks>
    Public Function ExecuteSQL(ByVal vsSQL As String) As Integer

        Try
            '接続
            Me.ConnectDb()

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            'SQL実行
            Me.mCmd.CommandText = vsSQL
            Return Me.mCmd.ExecuteNonQuery()

        Catch ex As Exception
            Me.RollBackTran()   'エラーが起きたときは、ロールバックする
            Throw New Exception(ex.Message & "SQL実行エラー" & vsSQL)

        End Try

    End Function

    ''' <summary>
    ''' ExecuteScalarの実行
    ''' </summary>
    ''' <param name="vsSQL"></param>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(ByVal vsSQL As String) As Object

        Try
            '接続
            Me.ConnectDb()

            'SQL実行
            Me.mCmd.CommandText = vsSQL
            Return Me.mCmd.ExecuteScalar()

        Catch ex As Exception
            Me.RollBackTran()   'エラーが起きたときは、ロールバックする
            Throw New Exception(ex.Message & "SQL実行エラー" & vsSQL)

        End Try

    End Function

    ''' <summary>
    ''' トランザクションの開始
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BeginTran()

        Try
            '接続
            Me.ConnectDb()

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            If Not mTran Is Nothing Then
                mTran = Nothing
            End If

            mTran = Me.mConn.BeginTransaction
            Me.mCmd.Transaction = mTran

        Catch ex As Exception
            Throw New Exception(ex.Message & "トランザクション開始エラー")

        End Try

    End Sub

    ''' <summary>
    ''' トランザクションのロールバック
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RollBackTran()

        Try
            '接続
            Me.ConnectDb()

            If mTran Is Nothing Then
                Return
            End If

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            mTran.Rollback()

            mTran = Nothing
            Me.mCmd.Transaction = Nothing

        Catch ex As Exception
            mTran = Nothing
            Me.mCmd.Transaction = Nothing

            Throw New Exception(ex.Message & "ロールバックエラー")

        End Try

    End Sub

    ''' <summary>
    ''' トランザクションのコミット
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CommitTran()

        Try
            '接続
            Me.ConnectDb()

            If mTran Is Nothing Then
                Return
            End If

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            mTran.Commit()

            mTran = Nothing
            Me.mCmd.Transaction = Nothing

        Catch ex As Exception
            mTran = Nothing
            Me.mCmd.Transaction = Nothing

            Throw New Exception(ex.Message & "コミットエラー")

        End Try

    End Sub

    ''' <summary>
    ''' DB接続を開くだけ。通常は使用しないこと！
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub callConnectDB()
        ConnectDb()
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 重複する呼び出しを検出するには

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: マネージ状態を破棄します (マネージ オブジェクト)。
            End If

            ' TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下の Finalize() をオーバーライドします。
            ' TODO: 大きなフィールドを null に設定します。
            If Not IsNothing(mDr) Then
                mDr.Close()
                mDr = Nothing
            End If
            If Not IsNothing(mTran) Then
                mTran.Rollback()
                mTran.Dispose()
            End If
            If Not IsNothing(mCmd) Then
                mCmd.Dispose()
            End If
            If Not IsNothing(mConn) Then
                mConn.Close()
                mConn.Dispose()
            End If

        End If
        Me.disposedValue = True
    End Sub

    ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    ''' <summary>
    ''' テーブルがあれば削除
    ''' </summary>
    ''' <param name="sTbl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function DeleteTable(ByVal sTbl As String) As Boolean

        'テーブルがあるか？
        If CheckTable(sTbl) = True Then
            'ある時は、削除

            ExecuteSQL("DROP TABLE " & sTbl)

        End If

        Return True

    End Function

    ''' <summary>
    ''' テーブルがあるかチェック
    ''' </summary>
    ''' <param name="sTbl"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CheckTable(sTbl As String) As Boolean

        Dim sSyst As String = ""
        Dim sName As String = ""
        Dim bRet As Boolean = False
        Dim wSQL As String = ""

        '検索するシステムデータベース(sysobjects)とテーブル名を取得
        If InStrRev(sTbl, ".") > 0 Then
            sSyst = Left(sTbl, InStrRev(sTbl, ".") - 1)
            sName = Mid(sTbl, InStrRev(sTbl, ".") + 1)

            wSQL = "SELECT * FROM information_schema.tables WHERE table_schema = '" & sSyst & "' AND table_name = '" & sName & "'"
            SelectSQL(wSQL)

            If IsSelectExistRecord() Then
                bRet = ReadDr()
            End If
            DrClose()
        End If

        Return bRet

    End Function

    ''' <summary>
    ''' クエリパラメータ
    ''' </summary>
    ''' <param name="sName"></param>
    ''' <param name="sValue"></param>
    ''' <remarks></remarks>
    Public Sub AddWithValue(ByVal sName As String, sValue As String)

        Try
            '接続
            Me.ConnectDb()

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            Me.mCmd.Parameters.AddWithValue(sName, sValue)

        Catch ex As Exception
            Me.RollBackTran()   'エラーが起きたときは、ロールバックする
            Throw New Exception(ex.Message & "SQL実行エラー" & sName)

        End Try

    End Sub

    ''' <summary>
    ''' クエリパラメータ初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ParameterClear()

        Try
            '接続
            Me.ConnectDb()

            If Not Me.mDr Is Nothing Then
                If Not Me.mDr.IsClosed Then
                    Me.mDr.Close()
                    Me.mDr = Nothing
                End If
            End If

            Me.mCmd.Parameters.Clear()

        Catch ex As Exception
            Me.RollBackTran()   'エラーが起きたときは、ロールバックする
            Throw New Exception(ex.Message & "SQL実行エラー")

        End Try

    End Sub

    ''' <summary>
    ''' Drが開いている場合、Closeするだけ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DrClose()
        If Not Me.mDr Is Nothing Then
            If Not Me.mDr.IsClosed Then
                Me.mDr.Close()
                Me.mDr = Nothing
            End If
        End If
    End Sub

End Class
