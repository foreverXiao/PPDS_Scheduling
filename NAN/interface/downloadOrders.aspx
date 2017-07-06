<%@ Page Title="Download orders" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="downloadOrders.aspx.vb" Inherits="interface_downloadOrders" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
<asp:LinkButton ID="listEDI" runat="server" style="position:absolute;right:150px;">List EDI files</asp:LinkButton>
<br />
<asp:Button ID="IMPorders" runat="server" Text="Import new and revision orders" OnClientClick="if (confirm('Are you sure you want to download new or revision orders?') == true ) { if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};} else {return false;};" />
    <br />
    <asp:Label ID="Label1" runat="server" Text="" ></asp:Label>
    <br />
    <br />
    </asp:Content>

