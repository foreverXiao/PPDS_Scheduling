<%@ Page  aspcompat="true"  Title="Plan timing and horizon" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="planparam.aspx.vb" Inherits="plansetting_planparam" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
    <asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy><br />
      <asp:FileUpload ID="FileUpload1" runat="server" 
    style="margin-bottom: 0px"  ClientIDMode="Static" Width="196px" />&nbsp;&nbsp;
    <asp:Button runat="server" 
        id="UpldUpdate" text="Update" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False" 
        Enabled="False"  />
    <asp:Button runat="server" 
        id="UpldDel" text="Delete" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False" 
        Enabled="False"   />
    <asp:Button runat="server" 
        id="UpldInsrt" text="Insert" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False"   />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="overwrite" 
        style="left:80px;" runat="server" Text="Upload and overwrite" 
        OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" 
        Enabled="False"  />
    <asp:Button ID="Download1" runat="server" Text="Download current selection" 
        ViewStateMode="Disabled" style="left:100px;" ClientIDMode="Static" /><asp:HiddenField ID="hiddenBT" runat="server" /><hr style="color:white;height:1px" />
      <asp:Label ID="StatusLabel" runat="server" Text=""></asp:Label><hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP1" runat="server" ClientIDMode="Static"><ContentTemplate>
    <asp:DropDownList ID="DDL1" runat="server" ClientIDMode="Static" 
        AutoPostBack="True" ViewStateMode="Enabled">
    </asp:DropDownList><asp:DropDownList ID="DDL2" runat="server" ClientIDMode="Static">
    </asp:DropDownList>
    <asp:TextBox ID="filtercdtn1" runat="server" ClientIDMode="Static" Width="90px" 
            AutoPostBack="False"  ></asp:TextBox>
    <asp:Button ID="Filter1" runat="server" Text="Filter" ClientIDMode="Static"  />
        <asp:Button ID="clrfltr1"
        runat="server" Text="Clear Filter" ClientIDMode="Static" />
    </ContentTemplate><triggers><asp:AsyncPostBackTrigger ControlID="DDL1" EventName="TextChanged" /></triggers></asp:UpdatePanel>
    <hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP2" runat="server" ClientIDMode="Static"><ContentTemplate>
    <asp:Label ID="Message"
        ForeColor="Red"          
        runat="server"/><br />
      <asp:DataPager ID="DP1" runat="server" ClientIDMode="Static" 
          PagedControlID="LV1" ViewStateMode="Enabled">
          <Fields>
              <asp:NextPreviousPagerField ButtonType="Button" ShowFirstPageButton="True" 
                  ShowNextPageButton="False" ShowPreviousPageButton="False" 
                  FirstPageText="|&lt;" LastPageText="&gt;|" />
              <asp:NumericPagerField NextPageText="&gt;&gt;" 
                  PreviousPageText="&lt;&lt;" />
              <asp:NextPreviousPagerField ButtonType="Button" ShowLastPageButton="True" 
                  ShowNextPageButton="False" ShowPreviousPageButton="False" 
                  LastPageText="&gt;|" />
          </Fields>
      </asp:DataPager>
      <asp:ListView ID="LV1" runat="server" ClientIDMode="Static" 
          DataSourceID="SDS1" DataKeyNames="category,paramname" 
            ViewStateMode="Enabled">
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC; ">
                  <td>
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="categoryLabel" runat="server" 
                          Text='<%# Eval("category") %>' />
                  </td>
                  <td>
                      <asp:Label ID="paramnameLabel" runat="server" 
                          Text='<%# Eval("paramname") %>' />
                  </td>
                  <td>
                      <asp:Label ID="paramvalueLabel" runat="server" 
                          Text='<%# Eval("paramvalue") %>' />
                  </td>
              </tr>
          </AlternatingItemTemplate>
          <EditItemTemplate>
              <tr style="background-color:#008A8C; color: #FFFFFF;">
                  <td>
                      <asp:Button ID="UpdateButton" runat="server" CommandName="Update" 
                          Text="Update" />
                      <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" 
                          Text="Cancel" />
                  </td>
                  <td>
                      <asp:Label ID="categoryLabel1" runat="server" 
                          Text='<%# Eval("category") %>' />
                  </td>
                  <td>
                       <asp:Label ID="paramnameLabel1" runat="server" 
                           Text='<%# Eval("paramname") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="paramvalueTextBox" runat="server" 
                          Text='<%# Bind("paramvalue") %>' />
                  </td>
              </tr>
          </EditItemTemplate>
          <EmptyDataTemplate>
              <table runat="server" 
                  
                  style="background-color:#FFFFFF; border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;">
                  <tr>
                      <td>
                          No data was returned.</td>
                  </tr>
              </table>
          </EmptyDataTemplate>
          <InsertItemTemplate>
              <tr style="">
                  <td>
                      <asp:Button ID="InsertButton" runat="server" CommandName="Insert" 
                          Text="Insert" />
                      <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" 
                          Text="Clear" />
                  </td>
                  <td>
                      <asp:TextBox ID="categoryTextBox" runat="server" 
                          Text='<%# Bind("category") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="paramnameTextBox" runat="server" 
                          Text='<%# Bind("paramname") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="paramvalueTextBox" runat="server" 
                          Text='<%# Bind("paramvalue") %>' />
                  </td>
              </tr>
          </InsertItemTemplate>
          <ItemTemplate>
              <tr style="background-color:#DCDCDC; color: #000000;">
                  <td>
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="categoryLabel" runat="server" 
                          Text='<%# Eval("category") %>' />
                  </td>
                  <td>
                      <asp:Label ID="paramnameLabel" runat="server" 
                          Text='<%# Eval("paramname") %>' />
                  </td>
                  <td>
                      <asp:Label ID="paramvalueLabel" runat="server" 
                          Text='<%# Eval("paramvalue") %>' />
                  </td>
              </tr>
          </ItemTemplate>
          <LayoutTemplate>
              <table runat="server">
                  <tr runat="server">
                      <td runat="server">
                          <table ID="itemPlaceholderContainer" runat="server" border="1" 
                              style="background-color: #FFFFFF;border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;font-family: Verdana, Arial, Helvetica, sans-serif;">
                              <tr runat="server" style="background-color:#DCDCDC; color: #000000;">
                                  <th runat="server">
                                  </th>
                                  <th runat="server">
                                      category</th>
                                  <th runat="server">
                                      paramname</th>
                                  <th runat="server">
                                      paramvalue</th>
                              </tr>
                              <tr runat="server" ID="itemPlaceholder">
                              </tr>
                          </table>
                      </td>
                  </tr>
                  <tr runat="server">
                      <td runat="server" 
                          
                          
                          style="text-align: center;background-color: #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif;color: #000000;">
                      </td>
                  </tr>
              </table>
          </LayoutTemplate>
            <SelectedItemTemplate>
                <tr style="background-color:#008A8C; font-weight: bold;color: #FFFFFF;">
                    <td>
                        <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                    </td>
                    <td>
                        <asp:Label ID="categoryLabel" runat="server" 
                            Text='<%# Eval("category") %>' />
                    </td>
                    <td>
                        <asp:Label ID="paramnameLabel" runat="server" 
                            Text='<%# Eval("paramname") %>' />
                    </td>
                    <td>
                        <asp:Label ID="paramvalueLabel" runat="server" 
                            Text='<%# Eval("paramvalue") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\inetpub\wwwroot\test\App_Data\param.mdb;Persist Security Info=True" 
        ProviderName="System.Data.SqlClient"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT [category], [paramname], [paramvalue] FROM [Esch_Na_tbl_plnprmtr] ORDER BY [category], [paramname] DESC" 
    DeleteCommand="DELETE FROM [Esch_Na_tbl_plnprmtr] WHERE ([category] = @original_category) AND ([paramname] = @original_paramname)" 
                                  
            UpdateCommand="UPDATE [Esch_Na_tbl_plnprmtr] SET [paramvalue] = @paramvalue WHERE ([category] = @original_category) AND ([paramname] = @original_paramname) " 
                                           
            
            
            
            InsertCommand="INSERT INTO [Esch_Na_tbl_plnprmtr] ([category], [paramname], [paramvalue]) VALUES (@category, @paramname, @paramvalue)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_category" Type="String" />
            <asp:Parameter Name="original_paramname" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="category" Type="String" />
            <asp:Parameter Name="paramname" Type="String" />
            <asp:Parameter Name="paramvalue" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="paramvalue" Type="String" />
            <asp:Parameter Name="original_category" Type="String" />
            <asp:Parameter Name="original_paramname" Type="String" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

