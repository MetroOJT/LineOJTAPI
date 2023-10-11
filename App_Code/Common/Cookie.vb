Option Explicit On
Option Strict On

Imports System.Web

Public Class Cookie

    ''' <summary>
    ''' Cookieに保持
    ''' </summary>
    ''' <param name="sname">名称</param>
    ''' <param name="svalue">値</param>
    ''' <param name="iexpiredate">有効期限（日）</param>
    ''' <remarks></remarks>
    Public Sub Set_Cookies(ByVal sname As String, ByVal svalue As String, ByVal iexpiredate As Integer)
        HttpContext.Current.Response.Cookies(sname).Value = HttpUtility.UrlEncode(svalue)
        HttpContext.Current.Response.Cookies(sname).Expires = DateTime.Now.AddDays(iexpiredate)
    End Sub

    ''' <summary>
    ''' Cookieの値を取得
    ''' </summary>
    ''' <param name="sname">名称</param>
    ''' <returns>値</returns>
    ''' <remarks></remarks>
    Public Function Get_Cookies(ByVal sname As String) As String
        Dim sret As String = Nothing
        If HttpContext.Current.Request.Cookies(sname) IsNot Nothing Then
            sret = HttpUtility.UrlDecode(HttpContext.Current.Request.Cookies(sname).Value)
        End If
        Return sret
    End Function

    ''' <summary>
    ''' Cookieの値を削除
    ''' </summary>
    ''' <param name="sname">名称</param>
    ''' <remarks></remarks>
    Public Sub Release_Cookies(ByVal sname As String)
        HttpContext.Current.Response.Cookies(sname).Value = Nothing
        HttpContext.Current.Response.Cookies(sname).Expires = DateTime.Now.AddDays(-10)
    End Sub

    ''' <summary>
    ''' 保持している全Cookieを削除
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Init_AllCookies()

        Dim MyCookieColl As HttpCookieCollection
        Dim MyCookie As HttpCookie
        Dim iLoop As Integer = 0

        MyCookieColl = HttpContext.Current.Request.Cookies
        Dim arr1 As String() = MyCookieColl.AllKeys

        For iLoop = 0 To arr1.Length - 1
            MyCookie = MyCookieColl(arr1(iLoop))
            Release_Cookies(MyCookie.Name)
        Next

    End Sub

End Class
