<%@ Page Title="login by SSO" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="loginBySSO.aspx.vb" Inherits="usermanage_loginBySSO" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
        .style1
        {
            height: 28px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
            <div runat="server" align="center" > <br />
            <table cellpadding="1" cellspacing="0" style="border-collapse:collapse;">
                <tr>
                    <td>
                        <table cellpadding="0">
                            <tr>
                                <td align="center" colspan="2">
                                    Please use your SSO to log in</td>
                            </tr>
                            <tr>
                                <td>
                                    &nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="2" style="text-align:right">
                                    Your SSO:&nbsp;&nbsp;<asp:TextBox ID="yoursso" runat="server"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2">              
                                    SSO password:<asp:TextBox ID="ssopswrd" runat="server" TextMode="Password"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td align="right" colspan="2" class="style1">
                                    <asp:Button ID="LoginButton" runat="server" CommandName="Login" 
                                        onclick="LoginButton_Click" Text="Log In" ValidationGroup="Login1" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
    <asp:Label ID="st" runat="server" Text=""></asp:Label><br /></div>
</asp:Content>

