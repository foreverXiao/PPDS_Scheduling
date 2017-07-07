<%@ Page Title="Download batch status" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="downloadBatchstatus.aspx.vb" Inherits="interface_downloadBatchstatus" Theme="Monochrome" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="act" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
<asp:LinkButton ID="listEDI" runat="server" style="position:absolute;right:150px;">List EDI files</asp:LinkButton>
<asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy>
    <br />
    <asp:Button ID="fromQAandFTP" runat="server" Text="Batch status from QA eColor and OPM ftp server" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" /><br />
    <asp:UpdatePanel ID="UpdatePanel1" runat="server"><ContentTemplate><asp:Label ID="Label1" runat="server"
        Text="Label">Batch status got from QA eColor starting point:(MM/DD/YYYY 24h) </asp:Label><br />
    <asp:TextBox ID="txtStartPoint" runat="server" ClientIDMode="Static" Width="96px"></asp:TextBox>
        <asp:Image ID="Img1" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <act:CalendarExtender ID="CE1" runat="server" 
        TargetControlID="txtStartPoint" PopupButtonID="Img1">
    </act:CalendarExtender>&nbsp; Hour:
        <asp:DropDownList ID="ddlHour1" runat="server">
        </asp:DropDownList> Minute:
        <asp:DropDownList ID="ddlMinute1" runat="server">
        </asp:DropDownList><br /><br /><br />
 </ContentTemplate></asp:UpdatePanel>
 <asp:Label ID="Label11" runat="server"></asp:Label><br /><br />
</asp:Content>

