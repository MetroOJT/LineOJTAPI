<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Test.aspx.vb" Inherits="API_GetContent_Test" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server" method="post" action="https://www.m-netservice.jp/LineOJTAPI/API/GetContent/Index.aspx">
    <div>
    <textarea id="test" runat="server"></textarea>
        <input type="submit" value="送信" />
    </div>
    </form>
</body>
</html>
