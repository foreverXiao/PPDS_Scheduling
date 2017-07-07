<%@ Page   Title="update ex-plant date" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="explantDateUpdate.aspx.vb" Inherits="interface_explantDateUpdate" Theme="Monochrome" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="act" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
<asp:LinkButton ID="listEDI" runat="server" style="position:absolute;right:150px;">List EDI files</asp:LinkButton>
<asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy>
    <br />
    <asp:Button ID="explantToOPM" runat="server" Text="Upload ex-plant to OPM" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" />
    <br /><br />
    <asp:Label ID="message" runat="server" Text=""></asp:Label><br /><br />
    <asp:UpdatePanel ID="UpdatePanel1" runat="server"><ContentTemplate><asp:Label ID="Label1" runat="server"
        Text="Label">The earliest ex-plant date you want to upload:(MM/DD/YYYY) </asp:Label><br />
    <asp:TextBox ID="txtStartPoint" runat="server" ClientIDMode="Static" Width="96px"></asp:TextBox>
        <asp:Image ID="Img1" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <act:CalendarExtender ID="CE1" runat="server" 
        TargetControlID="txtStartPoint" PopupButtonID="Img1">
    </act:CalendarExtender>&nbsp;<br />
        <br /> <br />
 </ContentTemplate></asp:UpdatePanel>
 
</asp:Content>

