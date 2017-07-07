<%@ Page  aspcompat="true"  Title="MTI" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="MTI.aspx.vb" Inherits="dragDrop_normalOP_MTI" Theme="Monochrome" %>

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
        ViewStateMode="Disabled" style="left:100px;" ClientIDMode="Static" /><asp:HiddenField ID="hiddenBT"  runat="server"  /><hr style="color:white;height:1px" />
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
        <asp:Button ID="bckToDB"
        runat="server" Text="Add MTI orders to order table" ClientIDMode="Static" />
        &nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="btOM"
        runat="server" Text="Home" ClientIDMode="Static" />

    </ContentTemplate><triggers><asp:AsyncPostBackTrigger ControlID="DDL1" EventName="TextChanged" /></triggers></asp:UpdatePanel>
    <hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP2" runat="server" ClientIDMode="Static"><ContentTemplate>
    <asp:Label ID="Message"
        ForeColor="Red"          
        runat="server"/><br />
      <asp:DataPager ID="DP1" runat="server" ClientIDMode="Static" 
          PagedControlID="LV1" ViewStateMode="Enabled" PageSize="20">
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
        <!---'add three items in below ListView ---Gary Xu 20150108 --->
      <asp:ListView ID="LV1" runat="server" ClientIDMode="Static" 
          DataSourceID="SDS1" DataKeyNames="sequence" 
            ViewStateMode="Enabled">          
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC; ">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete"  OnClientClick="return confirm('Are you sure you want to delete this item?');"  />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:CheckBox ID="decideToAddCheckBox" runat="server" 
                          Checked='<%# Eval("decideToAdd") %>' Enabled="false" />
                  </td>
                  <td>
                      <asp:Label ID="sequenceLabel" runat="server" 
                          Text='<%# Eval("sequence") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_noLabel" runat="server" 
                          Text='<%# Eval("txt_order_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_order_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_local_soLabel" runat="server" 
                          Text='<%# Eval("txt_local_so") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_currency" runat="server" 
                          Text='<%# Eval("txt_currency") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_package_code" runat="server" 
                          Text='<%# Eval("txt_package_code")%>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_method" runat="server" 
                          Text='<%# Eval("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:Label ID="itemLabel" runat="server" 
                          Text='<%# Eval("item") %>' />
                  </td>
                  <td>
                      <asp:Label ID="quantityLabel" runat="server" Text='<%# Eval("quantity") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lineLabel" runat="server" Text='<%# Eval("line") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_gl_classLabel" runat="server" 
                          Text='<%# Eval("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:Label ID="startDateLabel" runat="server" Text='<%# Eval("startDate") %>' />
                  </td>
                  <td>
                      <asp:Label ID="remarkLabel" runat="server" Text='<%# Eval("remark") %>' />
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
                      <asp:CheckBox ID="decideToAddCheckBox" runat="server" 
                          Checked='<%# Bind("decideToAdd") %>' />
                  </td>
                  <td>
                       <asp:Label ID="sequenceLabel1" runat="server" Text='<%# Eval("sequence") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_noTextBox" runat="server" 
                          Text='<%# Bind("txt_order_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_line_noTextBox" runat="server" 
                          Text='<%# Bind("txt_order_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_local_soTextBox" runat="server" 
                          Text='<%# Bind("txt_local_so") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_currency" runat="server" 
                          Text='<%# Bind("txt_currency") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_package_code" runat="server" 
                          Text='<%# Bind("txt_package_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_method" runat="server" 
                          Text='<%# Bind("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="itemTextBox" runat="server" 
                          Text='<%# Bind("item") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="quantityTextBox" runat="server" 
                          Text='<%# Bind("quantity") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lineTextBox" runat="server" Text='<%# Bind("line") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_gl_classTextBox" runat="server" 
                          Text='<%# Bind("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="startDateTextBox" runat="server" 
                          Text='<%# Bind("startDate") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="remarkTextBox" runat="server" Text='<%# Bind("remark") %>' />
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
                      <asp:CheckBox ID="decideToAddCheckBox" runat="server" 
                          Checked='<%# Bind("decideToAdd") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="sequenceTextBox" runat="server" 
                          Text='<%# Bind("sequence") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_noTextBox" runat="server" 
                          Text='<%# Bind("txt_order_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_line_noTextBox" runat="server" 
                          Text='<%# Bind("txt_order_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_local_soTextBox" runat="server" 
                          Text='<%# Bind("txt_local_so") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_currency" runat="server" 
                          Text='<%# Bind("txt_currency") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_package_code" runat="server" 
                          Text='<%# Bind("txt_package_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_method" runat="server" 
                          Text='<%# Bind("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="itemTextBox" runat="server" 
                          Text='<%# Bind("item") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="quantityTextBox" runat="server" 
                          Text='<%# Bind("quantity") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lineTextBox" runat="server" Text='<%# Bind("line") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_gl_classTextBox" runat="server" 
                          Text='<%# Bind("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="startDateTextBox" runat="server" 
                          Text='<%# Bind("startDate") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="remarkTextBox" runat="server" Text='<%# Bind("remark") %>' />
                  </td>
              </tr>
          </InsertItemTemplate>
          <ItemTemplate>
              <tr style="background-color:#DCDCDC; color: #000000;">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete"  OnClientClick="return confirm('Are you sure you want to delete this item?');"  />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:CheckBox ID="decideToAddCheckBox" runat="server" 
                          Checked='<%# Eval("decideToAdd") %>' Enabled="false" />
                  </td>
                  <td>
                      <asp:Label ID="sequenceLabel" runat="server" 
                          Text='<%# Eval("sequence") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_noLabel" runat="server" 
                          Text='<%# Eval("txt_order_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_order_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_local_soLabel" runat="server" 
                          Text='<%# Eval("txt_local_so") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_currency" runat="server" 
                          Text='<%# Eval("txt_currency") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_package_code" runat="server" 
                          Text='<%# Eval("txt_package_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_method" runat="server" 
                          Text='<%# Eval("txt_ship_method") %>' />
                  </td>
                  <td>
                      <asp:Label ID="itemLabel" runat="server" 
                          Text='<%# Eval("item") %>' />
                  </td>
                  <td>
                      <asp:Label ID="quantityLabel" runat="server" Text='<%# Eval("quantity") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lineLabel" runat="server" Text='<%# Eval("line") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_gl_classLabel" runat="server" 
                          Text='<%# Eval("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:Label ID="startDateLabel" runat="server" Text='<%# Eval("startDate") %>' />
                  </td>
                  <td>
                      <asp:Label ID="remarkLabel" runat="server" Text='<%# Eval("remark") %>' />
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
                                      decideToAdd</th>
                                  <th runat="server">
                                      sequence</th>
                                  <th runat="server">
                                      txt_order_no</th>
                                  <th runat="server">
                                      txt_order_line_no</th>
                                  <th runat="server">
                                      txt_local_so</th>
                                  <th runat="server">
                                      txt_currency</th>
                                  <th runat="server">
                                      txt_package_code</th>
                                  <th runat="server">
                                      txt_ship_method</th>
                                  <th runat="server">
                                      item</th>
                                  <th runat="server">
                                      quantity</th>
                                  <th runat="server">
                                      line</th>
                                  <th runat="server">
                                      txt_gl_class</th>
                                  <th runat="server">
                                      startDate</th>
                                  <th runat="server">
                                      remark</th>
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
                            Text="Delete"  OnClientClick="return confirm('Are you sure you want to delete this item?');"  />
                        <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                    </td>
                    <td>
                        <asp:CheckBox ID="decideToAddCheckBox" runat="server" 
                            Checked='<%# Eval("decideToAdd") %>' Enabled="false" />
                    </td>
                    <td>
                        <asp:Label ID="sequenceLabel" runat="server" 
                            Text='<%# Eval("sequence") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_order_noLabel" runat="server" 
                            Text='<%# Eval("txt_order_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_order_line_noLabel" runat="server" 
                            Text='<%# Eval("txt_order_line_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_local_soLabel" runat="server" 
                            Text='<%# Eval("txt_local_so") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_currency" runat="server" 
                            Text='<%# Eval("txt_currency") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_package_code" runat="server" 
                            Text='<%# Eval("txt_package_code") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ship_method" runat="server" 
                            Text='<%# Eval("txt_ship_method") %>' />
                    </td>
                    <td>
                        <asp:Label ID="itemLabel" runat="server" 
                            Text='<%# Eval("item") %>' />
                    </td>

                    <td>
                        <asp:Label ID="quantityLabel" runat="server" Text='<%# Eval("quantity") %>' />
                    </td>
                    <td>
                        <asp:Label ID="lineLabel" runat="server" Text='<%# Eval("line") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_gl_classLabel" runat="server" 
                            Text='<%# Eval("txt_gl_class") %>' />
                    </td>
                    <td>
                        <asp:Label ID="startDateLabel" runat="server" Text='<%# Eval("startDate") %>' />
                    </td>
                    <td>
                        <asp:Label ID="remarkLabel" runat="server" Text='<%# Eval("remark") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>

        <!---'Change command in SDS DataSource, datasource change to map path ---Gary Xu 20150108 --->
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.ACE.OLEDB.12.0;Data Source='~/App_Data\db_Resin.accdb'" 
        ProviderName="System.Data.OleDb"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT * FROM [Esch_Sh_tbl_MTI_add] ORDER BY [sequence]" 
    DeleteCommand="DELETE FROM [Esch_Sh_tbl_MTI_add] WHERE [sequence] = ?" 
                                  
            UpdateCommand="UPDATE [Esch_Sh_tbl_MTI_add] SET [decideToAdd] = ?, [txt_order_no] = ?, [txt_order_line_no] = ?, [txt_local_so] = ?, [txt_currency] = ?, [txt_package_code] = ?, [txt_ship_method] = ?, [item] = ?, [quantity] = ?, [line] = ?, [txt_gl_class] = ?, [startDate] = ?, [remark] = ? WHERE [sequence] = ?" 
            
            
            
            
            
            
            
            
            
            
            
            
            
            InsertCommand="INSERT INTO [Esch_Sh_tbl_MTI_add] ([decideToAdd], [sequence], [txt_order_no], [txt_order_line_no], [txt_local_so], [txt_currency], [txt_package_code], [txt_ship_method],  [item], [quantity], [line], [txt_gl_class], [startDate], [remark]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_sequence" Type="Int16" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="decideToAdd" Type="Boolean" />
            <asp:Parameter Name="sequence" Type="Int16" />
            <asp:Parameter Name="txt_order_no" Type="String" />
            <asp:Parameter Name="txt_order_line_no" Type="String" />
            <asp:Parameter Name="txt_local_so" Type="String" />
            <asp:Parameter Name="txt_currency" Type="String" />
            <asp:Parameter Name="txt_package_code" Type="String" />
            <asp:Parameter Name="txt_ship_method" Type="String" />
            <asp:Parameter Name="item" Type="String" />
            <asp:Parameter Name="quantity" Type="Int32" />
            <asp:Parameter Name="line" Type="String" />
            <asp:Parameter Name="txt_gl_class" Type="String" />
            <asp:Parameter Name="startDate" Type="DateTime" />
            <asp:Parameter Name="remark" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="decideToAdd" Type="Boolean" />
            <asp:Parameter Name="txt_order_no" Type="String" />
            <asp:Parameter Name="txt_order_line_no" Type="String" />
            <asp:Parameter Name="txt_local_so" Type="String" />
            <asp:Parameter Name="txt_currency" Type="String" />
            <asp:Parameter Name="txt_package_code" Type="String" />
            <asp:Parameter Name="txt_ship_method" Type="String" />
            <asp:Parameter Name="item" Type="String" />
            <asp:Parameter Name="quantity" Type="Int32" />
            <asp:Parameter Name="line" Type="String" />
            <asp:Parameter Name="txt_gl_class" Type="String" />
            <asp:Parameter Name="startDate" Type="DateTime" />
            <asp:Parameter Name="remark" Type="String" />
            <asp:Parameter Name="original_sequence" Type="Int16" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

