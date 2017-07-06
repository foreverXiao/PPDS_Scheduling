<%@ Page Title="log out" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="logout.aspx.vb" Inherits="dragDrop_logout" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
<br />
    <asp:Literal ID="Literal1" runat="server" Text="User: "></asp:Literal>
      ===> <asp:Button ID="Button1" runat="server" Text="Log out" />
    <br /><br />
</asp:Content>

