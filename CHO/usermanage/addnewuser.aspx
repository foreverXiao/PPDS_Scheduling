<%@ Page Title="add new user" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="addnewuser.aspx.vb" Inherits="usermanage_addnewuser" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
<asp:HyperLink ID="chngPsswrd"
        runat="server" NavigateUrl="changepassword.aspx">Change password</asp:HyperLink>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:HyperLink ID="adduser"
        runat="server" NavigateUrl="addnewuser.aspx">Add new user</asp:HyperLink>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:HyperLink ID="dltUser"
        runat="server" NavigateUrl="deleteuser.aspx">Delete user</asp:HyperLink><h3>Add new user</h3>
    <asp:Label ID="Label1" runat="server" Text="User Name:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"></asp:Label>
    <asp:TextBox ID="tbUsNm" runat="server" MaxLength="20"></asp:TextBox>
    <asp:RequiredFieldValidator ID="RFV1" runat="server" 
        ErrorMessage="Required!" ControlToValidate="tbUsNm"></asp:RequiredFieldValidator>
    <asp:Label ID="Label5" runat="server" Text="User's description:&nbsp;"></asp:Label>
    <asp:TextBox ID="tbUsDscrptn" runat="server" MaxLength="20" Width="225px"></asp:TextBox><br />
    <asp:Label ID="Label2" runat="server" Text="User's role:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"></asp:Label><asp:DropDownList
        ID="roleDDL1" runat="server">
    </asp:DropDownList><br /><br />
    <asp:Label ID="Label3" runat="server" Text="Initial password:&nbsp;&nbsp;&nbsp;&nbsp;"></asp:Label>
    <asp:TextBox ID="tbNewpswrd" runat="server" TextMode="Password" MaxLength="20"></asp:TextBox><asp:RequiredFieldValidator
        ID="RFV3" runat="server" ErrorMessage="Required!" ControlToValidate="tbNewpswrd"></asp:RequiredFieldValidator><br />
    <asp:Label ID="Label4" runat="server" Text="confirm password:"></asp:Label>
    <asp:TextBox ID="tbNewpswrdAgn" runat="server" TextMode="Password" 
        MaxLength="20"></asp:TextBox><asp:CompareValidator
        ID="CV1" runat="server" ErrorMessage="Passwords do not match!" ControlToValidate="tbNewpswrdAgn" ControlToCompare="tbNewpswrd"></asp:CompareValidator>
    <br /><asp:CustomValidator ID="CstmV1" runat="server" ErrorMessage="Rules are violated." OnServerValidate="ServerValidation" >wrong input</asp:CustomValidator>
    <br /><asp:Button ID="btSubmit" runat="server" Text="Submit" />
    <asp:Label ID="lbStatus" runat="server"></asp:Label>
</asp:Content>

