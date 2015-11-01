<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Demo.aspx.vb" Inherits="WebApp.Demo" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Button ID="Button1" runat="server" Text="Button" />
            <asp:CheckBox ID="CheckBox1" runat="server" />
            <asp:DropDownList ID="DropDownList1" runat="server"></asp:DropDownList>
            <asp:HiddenField ID="HiddenField1" runat="server" />
            <asp:LinkButton ID="LinkButton1" runat="server">LinkButton</asp:LinkButton>
            <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
            <asp:CheckBoxList ID="CheckBoxList1" runat="server"></asp:CheckBoxList>
            <asp:Image ID="Image1" runat="server" />
            <asp:ListBox ID="ListBox1" runat="server"></asp:ListBox>
            <asp:RadioButton ID="RadioButton1" runat="server" GroupName="rdo" />
            <asp:RadioButton ID="RadioButton2" runat="server" GroupName="rdo" />
            <asp:RadioButtonList ID="RadioButtonList1" runat="server"></asp:RadioButtonList>
        </div>
    </form>
</body>
</html>
