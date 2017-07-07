<%@ Page  aspcompat="true"  Title="orders for RM usage" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ordersForRM.aspx.vb" Inherits="SCMrelated_ordersForRM" Theme="Monochrome" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
    <asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy>
    <br />
    <asp:Label ID="StatusLabel" runat="server" Text=""></asp:Label>
    <asp:UpdatePanel ID="UP1" runat="server" ClientIDMode="Static"><ContentTemplate>
    <b>Start time or finish time&nbsp;&nbsp;between</b>&nbsp;&nbsp;&nbsp;<asp:TextBox ID="earlierTime" runat="server"  Width="72px"></asp:TextBox><asp:Image ID="Img1" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <asp:CalendarExtender ID="CE1" runat="server" TargetControlID="earlierTime" PopupButtonID="Img1">
        </asp:CalendarExtender> Hour:
        <asp:DropDownList ID="ddlHour1" runat="server">
        </asp:DropDownList> Minute:
        <asp:DropDownList ID="ddlMinute1" runat="server">
        </asp:DropDownList> &nbsp;&nbsp;<b>and</b>&nbsp;&nbsp; <asp:TextBox ID="laterTime" runat="server" Width="72px"></asp:TextBox><asp:Image ID="Img2" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <asp:CalendarExtender ID="CE2" runat="server" 
        TargetControlID="laterTime" PopupButtonID="Img2">
    </asp:CalendarExtender> Hour:
        <asp:DropDownList ID="ddlHour2" runat="server">
        </asp:DropDownList> Minute:
        <asp:DropDownList ID="ddlMinute2" runat="server">
        </asp:DropDownList><br />
         <hr style="color:white;height:1px" />
        <asp:Button ID="prdctnOrdrs" runat="server" 
            Text="Prepare production orders list" 
            ToolTip="sort out all the orders to be created batch no. based on the selected production lines and timing" OnClientClick="if (this.value.indexOf('wait') > 0 ){this.disabled=true;}else{this.value='Please wait for a moment......';};"  />
         <asp:Label ID="Label1" runat="server" ></asp:Label>
        </ContentTemplate></asp:UpdatePanel>
    <hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP2" runat="server" ClientIDMode="Static"><ContentTemplate>
    <asp:Label ID="Message"
        ForeColor="Blue"          
        runat="server"/><br />
        <asp:GridView ID="GV1" runat="server" AllowPaging="True">
        </asp:GridView>
    </ContentTemplate><Triggers>
    </Triggers></asp:UpdatePanel>
    <hr style="color:white;height:1px" />
        <asp:Button ID="Download1" runat="server" ClientIDMode="Static" 
            style="left:0px;" Text="Download current selection" 
            ViewStateMode="Disabled" /><br /><br />
       
</asp:Content>

