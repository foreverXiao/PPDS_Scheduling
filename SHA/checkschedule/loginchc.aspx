<%@ Page Language="VB" AutoEventWireup="false" CodeFile="loginchc.aspx.vb" Inherits="checkschedule_login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:LinkButton ID="logout" runat="server" align="left" >Log out</asp:LinkButton>
    <div align = "center" runat="server"  >
                <table cellpadding="1" cellspacing="0" style="border-collapse:collapse;">
                    <tr>
                        <td>
                            <table cellpadding="0">
                                <tr>
                                    <td align="center" colspan="2">
                                        Log In</td>
                                </tr>
                                <tr>
                                    <td align="right">
                                        <asp:Label ID="UserNameLabel" runat="server" AssociatedControlID="username1">User Name:</asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox ID="username1" runat="server"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="UserNameRequired" runat="server" 
                                            ControlToValidate="username1" ErrorMessage="User Name is required." 
                                            ToolTip="User Name is required." ValidationGroup="Login1">*</asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="right">
                                        <asp:Label ID="PasswordLabel" runat="server" AssociatedControlID="password1">Password:</asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox ID="password1" runat="server" TextMode="Password"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="PasswordRequired" runat="server" 
                                            ControlToValidate="password1" ErrorMessage="Password is required." 
                                            ToolTip="Password is required." ValidationGroup="Login1">*</asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="center" colspan="2" style="color:Red;">
                                        &nbsp;</td>
                                </tr>
                                <tr>
                                    <td align="right" colspan="2">
                                        <asp:Button ID="LoginButton" runat="server" CommandName="Login" 
                                            onclick="LoginButton_Click" Text="Log In" ValidationGroup="Login1" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
    <asp:Label runat ="server" ID="st" Text="" /><br />
    </div>
    </form>
</body>
</html>
