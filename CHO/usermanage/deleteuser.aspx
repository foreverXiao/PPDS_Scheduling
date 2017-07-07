<%@ Page Title="delete user" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="deleteuser.aspx.vb" Inherits="usermanage_deleteuser" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
<asp:HyperLink ID="chngPsswrd"
        runat="server" NavigateUrl="changepassword.aspx">Change password</asp:HyperLink>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:HyperLink ID="adduser"
        runat="server" NavigateUrl="addnewuser.aspx">Add new user</asp:HyperLink>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:HyperLink ID="dltUser"
        runat="server" NavigateUrl="deleteuser.aspx">Delete user</asp:HyperLink><h3>
        Delete user</h3>
    <asp:GridView ID="GV1" runat="server" AllowPaging="True" AllowSorting="True" 
        AutoGenerateColumns="False" CellPadding="4" DataKeyNames="user_name" 
        DataSourceID="SDS1" ForeColor="#333333" GridLines="None" 
        PageSize="20">
        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        <Columns>
            <asp:CommandField ShowDeleteButton="True" />
            <asp:BoundField DataField="user_name" HeaderText="user_name" ReadOnly="True" 
                SortExpression="user_name" />
             <asp:TemplateField HeaderText="rightlevel" SortExpression="rightlevel">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("rightlevel") %>'></asp:TextBox>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label1" runat="server" Text='<%# formatData(Eval("rightlevel")) %>'></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:BoundField DataField="user_description" HeaderText="user_description" 
                SortExpression="user_description" />
        </Columns>
        <EditRowStyle BackColor="#999999" />
        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
        <SortedAscendingCellStyle BackColor="#E9E7E2" />
        <SortedAscendingHeaderStyle BackColor="#506C8C" />
        <SortedDescendingCellStyle BackColor="#FFFDF8" />
        <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
    </asp:GridView>
    <br />
    <br />
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ProviderName="System.Data.SqlClient"
        ConnectionString="Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\inetpub\wwwroot\test\App_Data\param.accdb;" 
        DeleteCommand="DELETE FROM [Esch_CQ_tbl_userrole] WHERE ([user_name] = @user_name)"   
        
        SelectCommand="SELECT [user_name], [rightlevel], [user_description] FROM [Esch_CQ_tbl_userrole] ORDER BY [rightlevel], [user_name]" 
>
        <DeleteParameters>
            <asp:Parameter Name="user_name" Type="String" />
        </DeleteParameters>

    </asp:SqlDataSource>
</asp:Content>

