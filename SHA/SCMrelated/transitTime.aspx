<%@ Page  aspcompat="true"  Title="transit time" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="transitTime.aspx.vb" Inherits="SCMrelated_transitTime" Theme="Monochrome" %>

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
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False"  />
    <asp:Button runat="server" 
        id="UpldInsrt" text="Insert" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False"  />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="overwrite" style="left:80px;" runat="server" Text="Upload and overwrite" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};"  />
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
          DataSourceID="SDS1" DataKeyNames="txt_currency,txt_destination,txt_ship_method" 
            ViewStateMode="Enabled" >
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC; ">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="txt_currencyLabel" runat="server" 
                          Text='<%# Eval("txt_currency") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_toLabel" runat="server" 
                          Text='<%# Eval("txt_ship_to") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_destinationLabel" runat="server" 
                          Text='<%# Eval("txt_destination") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_methodLabel" runat="server" 
                          Text='<%# Eval("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_transitLabel" runat="server" 
                          Text='<%# Eval("flt_transit") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actualLabel" runat="server" 
                          Text='<%# Eval("flt_actual") %>' />
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
                      <asp:Label ID="txt_currencyLabel1" runat="server" 
                          Text='<%# Eval("txt_currency") %>' />
                  </td>
                  <td>
                       <asp:TextBox ID="txt_ship_toTextBox" runat="server" 
                           Text='<%# Bind("txt_ship_to") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_destinationLabel1" runat="server" 
                          Text='<%# Eval("txt_destination") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_methodLabel1" runat="server" 
                          Text='<%# Eval("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_transitTextBox" runat="server" 
                          Text='<%# Bind("flt_transit") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actualTextBox" runat="server" 
                          Text='<%# Bind("flt_actual") %>' />
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
                      <asp:TextBox ID="txt_currencyTextBox" runat="server" 
                          Text='<%# Bind("txt_currency") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_toTextBox" runat="server" 
                          Text='<%# Bind("txt_ship_to") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_destinationTextBox" runat="server" 
                          Text='<%# Bind("txt_destination") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_methodTextBox" runat="server" 
                          Text='<%# Bind("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_transitTextBox" runat="server" 
                          Text='<%# Bind("flt_transit") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actualTextBox" runat="server" 
                          Text='<%# Bind("flt_actual") %>' />
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
                      <asp:Label ID="txt_currencyLabel" runat="server" 
                          Text='<%# Eval("txt_currency") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_toLabel" runat="server" 
                          Text='<%# Eval("txt_ship_to") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_destinationLabel" runat="server" 
                          Text='<%# Eval("txt_destination") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_methodLabel" runat="server" 
                          Text='<%# Eval("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_transitLabel" runat="server" 
                          Text='<%# Eval("flt_transit") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actualLabel" runat="server" 
                          Text='<%# Eval("flt_actual") %>' />
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
                                  <th  runat="server">
                                  <asp:Button ID = "btnNew" runat="server" Text="New" CommandName = "new" />
                                  </th>
                                  <th runat="server">
                                      txt_currency</th>
                                  <th runat="server">
                                      txt_ship_to</th>
                                  <th runat="server">
                                      txt_destination</th>
                                  <th runat="server">
                                      txt_ship_method</th>
                                  <th runat="server">
                                      flt_transit</th>
                                  <th runat="server">
                                      flt_actual</th>
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
                        <asp:Label ID="txt_currencyLabel" runat="server" 
                            Text='<%# Eval("txt_currency") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ship_toLabel" runat="server" 
                            Text='<%# Eval("txt_ship_to") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_destinationLabel" runat="server" 
                            Text='<%# Eval("txt_destination") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ship_methodLabel" runat="server" 
                            Text='<%# Eval("txt_ship_method") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_transitLabel" runat="server" 
                            Text='<%# Eval("flt_transit") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_actualLabel" runat="server" 
                            Text='<%# Eval("flt_actual") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\IIS\Test\App_Data\db_Resin.mdb" 
        ProviderName="System.Data.SqlClient"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT [txt_currency], [txt_ship_to], [txt_destination], [txt_ship_method], [flt_transit], [flt_actual] FROM [Esch_Sh_tbl_transit]" 
    DeleteCommand="DELETE FROM [Esch_Sh_tbl_transit]  WHERE ([txt_currency] = ?) AND ([txt_destination] = ?) AND ([txt_ship_method] = ?)" 
                                  
            UpdateCommand="UPDATE [Esch_Sh_tbl_transit] SET [txt_ship_to] = ?, [flt_transit] = ?, [flt_actual] = ?  WHERE ([txt_currency] = ?) AND ([txt_destination] = ?) AND ([txt_ship_method] = ?)" 
                                           
            InsertCommand="INSERT INTO [Esch_Sh_tbl_transit] ([txt_currency], [txt_ship_to], [txt_destination], [txt_ship_method], [flt_transit], [flt_actual]) VALUES (?, ?, ?, ?, ?, ?)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_txt_currency" Type="String" />
            <asp:Parameter Name="original_txt_destination" Type="String" />
            <asp:Parameter Name="original_txt_ship_method" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="txt_currency" Type="String" />
            <asp:Parameter Name="txt_ship_to" Type="String" />
            <asp:Parameter Name="txt_destination" Type="String" />
            <asp:Parameter Name="txt_ship_method" Type="String" />
            <asp:Parameter Name="flt_transit" Type="Int16" />
            <asp:Parameter Name="flt_actual" Type="Int16" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="txt_ship_to" Type="String" />
            <asp:Parameter Name="flt_transit" Type="Int16" />
            <asp:Parameter Name="flt_actual" Type="Int16" />
            <asp:Parameter Name="original_txt_currency" Type="String" />
            <asp:Parameter Name="original_txt_destination" Type="String" />
            <asp:Parameter Name="original_txt_ship_method" Type="String" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

