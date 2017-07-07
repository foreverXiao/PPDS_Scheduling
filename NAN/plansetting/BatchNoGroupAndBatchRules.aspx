<%@ Page  aspcompat="true"  Title="Batch NO group and batch rules" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="BatchNoGroupAndBatchRules.aspx.vb" Inherits="plansetting_BatchNoGroupAndBatchRules" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
    <asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy><br />
      <asp:FileUpload ID="FileUpload1" runat="server" 
    style="margin-bottom: 0px"  ClientIDMode="Static" Width="196px" />&nbsp;&nbsp;
    <asp:Button runat="server" 
        id="UpldUpdate" text="Update" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False"  />
    <asp:Button runat="server" 
        id="UpldDel" text="Delete" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False" 
        Enabled="False"   />
    <asp:Button runat="server" 
        id="UpldInsrt" text="Insert" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False"  />
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
          PagedControlID="LV1" ViewStateMode="Enabled" PageSize="40">
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
          DataSourceID="SDS1" DataKeyNames="int_Batch_NO_group" 
            ViewStateMode="Enabled" >
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC; ">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="int_Batch_NO_groupLabel" runat="server" 
                          Text='<%# Eval("int_Batch_NO_group") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_minimum_noLabel" runat="server" 
                          Text='<%# Eval("txt_minimum_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_current_noLabel" runat="server" 
                          Text='<%# Eval("txt_current_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_last_noLabel" runat="server" 
                          Text='<%# Eval("txt_last_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_maximum_noLabel" runat="server" 
                          Text='<%# Eval("txt_maximum_no") %>' />
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
                      <asp:Label ID="int_Batch_NO_groupLabel1" runat="server" 
                          Text='<%# Eval("int_Batch_NO_group") %>' />
                  </td>
                  <td>
                       <asp:TextBox ID="txt_minimum_noTextBox" runat="server" 
                           Text='<%# Bind("txt_minimum_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_current_noTextBox" runat="server" 
                          Text='<%# Bind("txt_current_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_last_noTextBox" runat="server" 
                          Text='<%# Bind("txt_last_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_maximum_noTextBox" runat="server" 
                          Text='<%# Bind("txt_maximum_no") %>' />
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
                      <asp:TextBox ID="int_Batch_NO_groupTextBox" runat="server" 
                          Text='<%# Bind("int_Batch_NO_group") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_minimum_noTextBox" runat="server" 
                          Text='<%# Bind("txt_minimum_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_current_noTextBox" runat="server" 
                          Text='<%# Bind("txt_current_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_last_noTextBox" runat="server" 
                          Text='<%# Bind("txt_last_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_maximum_noTextBox" runat="server" 
                          Text='<%# Bind("txt_maximum_no") %>' />
                  </td>
              </tr>
          </InsertItemTemplate>
          <ItemTemplate>
              <tr style="background-color:#DCDCDC; color: #000000;">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="int_Batch_NO_groupLabel" runat="server" 
                          Text='<%# Eval("int_Batch_NO_group") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_minimum_noLabel" runat="server" 
                          Text='<%# Eval("txt_minimum_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_current_noLabel" runat="server" 
                          Text='<%# Eval("txt_current_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_last_noLabel" runat="server" 
                          Text='<%# Eval("txt_last_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_maximum_noLabel" runat="server" 
                          Text='<%# Eval("txt_maximum_no") %>' />
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
                                  <asp:Button ID = "btnNew" runat="server" Text="New" CommandName = "new" />
                                  </th>
                                  <th runat="server">
                                      int_Batch_NO_group</th>
                                  <th runat="server">
                                      txt_minimum_no</th>
                                  <th runat="server">
                                      txt_current_no</th>
                                  <th runat="server">
                                      txt_last_no</th>
                                  <th runat="server">
                                      txt_maximum_no</th>
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
                        <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                            Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                        <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                    </td>
                    <td>
                        <asp:Label ID="int_Batch_NO_groupLabel" runat="server" 
                            Text='<%# Eval("int_Batch_NO_group") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_minimum_noLabel" runat="server" 
                            Text='<%# Eval("txt_minimum_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_current_noLabel" runat="server" 
                            Text='<%# Eval("txt_current_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_last_noLabel" runat="server" 
                            Text='<%# Eval("txt_last_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_maximum_noLabel" runat="server" 
                            Text='<%# Eval("txt_maximum_no") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\IIS\Test\App_Data\param.mdb" 
        ProviderName="System.Data.SqlClient"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT * FROM [Esch_Na_tbl_batch_no_group_and_batch_rules] ORDER BY [int_Batch_NO_group]" 
    DeleteCommand="DELETE FROM [Esch_Na_tbl_batch_no_group_and_batch_rules] WHERE (([int_Batch_NO_group] = ?))" 
                                  
            UpdateCommand="UPDATE [Esch_Na_tbl_batch_no_group_and_batch_rules] SET [txt_minimum_no] = ?, [txt_current_no] = ?, [txt_last_no] = ?, [txt_maximum_no] = ? WHERE (([int_Batch_NO_group] = ?))" 
                                           
            
            
            
            
            
            
            
            InsertCommand="INSERT INTO [Esch_Na_tbl_batch_no_group_and_batch_rules] ([int_Batch_NO_group], [txt_minimum_no], [txt_current_no], [txt_last_no], [txt_maximum_no]) VALUES (?, ?, ?, ?, ?)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_int_Batch_NO_group" Type="Int16" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="int_Batch_NO_group" Type="Int16" />
            <asp:Parameter Name="txt_minimum_no" Type="String" />
            <asp:Parameter Name="txt_current_no" Type="String" />
            <asp:Parameter Name="txt_last_no" Type="String" />
            <asp:Parameter Name="txt_maximum_no" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="txt_minimum_no" Type="String" />
            <asp:Parameter Name="txt_current_no" Type="String" />
            <asp:Parameter Name="txt_last_no" Type="String" />
            <asp:Parameter Name="txt_maximum_no" Type="String" />
            <asp:Parameter Name="original_int_Batch_NO_group" Type="Int16" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

