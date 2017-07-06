<%@ Page  aspcompat="true"  Title="Group Lines Per Batch No" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="PrdctnLinesAndBatchNoGroup.aspx.vb" Inherits="plansetting_PrdctnLinesAndBatchNoGroup" Theme="Monochrome" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

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
      <asp:ListView ID="LV1" runat="server" ClientIDMode="Static" 
          DataSourceID="SDS1" DataKeyNames="int_Batch_NO_group,sequencePerBatchNoGroup" 
            ViewStateMode="Enabled" >
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC; ">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="int_Batch_NO_groupLabel" runat="server" 
                          Text='<%# Eval("int_Batch_NO_group") %>' />
                  </td>
                  <td>
                      <asp:Label ID="sequencePerBatchNoGroupLabel" runat="server" 
                          Text='<%# Eval("sequencePerBatchNoGroup") %>' />
                  </td>
                  <td>
                      <asp:Label ID="columnNameLabel" runat="server" 
                          Text='<%# Eval("columnName") %>' />
                  </td>
                  <td>
                      <asp:Label ID="relationOperatorLabel" runat="server" 
                          Text='<%# Eval("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:Label ID="conditionValueLabel" runat="server" 
                          Text='<%# Eval("conditionValue") %>' />
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
                       <asp:Label ID="sequencePerBatchNoGroupLabel1" runat="server" 
                           Text='<%# Eval("sequencePerBatchNoGroup") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="columnNameTextBox" runat="server" 
                          Text='<%# Bind("columnName") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="relationOperatorTextBox" runat="server" 
                          Text='<%# Bind("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="conditionValueTextBox" runat="server" 
                          Text='<%# Bind("conditionValue") %>' />
                  </td>
              </tr>
          </EditItemTemplate>
          <EmptyDataTemplate>
              <table id="Table1" runat="server" 
                  
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
                      <asp:TextBox ID="sequencePerBatchNoGroupTextBox" runat="server" 
                          Text='<%# Bind("sequencePerBatchNoGroup") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="columnNameTextBox" runat="server" 
                          Text='<%# Bind("columnName") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="relationOperatorTextBox" runat="server" 
                          Text='<%# Bind("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="conditionValueTextBox" runat="server" 
                          Text='<%# Bind("conditionValue") %>' />
                  </td>
              </tr>
          </InsertItemTemplate>
          <ItemTemplate>
              <tr style="background-color:#DCDCDC; color: #000000;">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="int_Batch_NO_groupLabel" runat="server" 
                          Text='<%# Eval("int_Batch_NO_group") %>' />
                  </td>
                  <td>
                      <asp:Label ID="sequencePerBatchNoGroupLabel" runat="server" 
                          Text='<%# Eval("sequencePerBatchNoGroup") %>' />
                  </td>
                  <td>
                      <asp:Label ID="columnNameLabel" runat="server" 
                          Text='<%# Eval("columnName") %>' />
                  </td>
                  <td>
                      <asp:Label ID="relationOperatorLabel" runat="server" 
                          Text='<%# Eval("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:Label ID="conditionValueLabel" runat="server" 
                          Text='<%# Eval("conditionValue") %>' />
                  </td>
              </tr>
          </ItemTemplate>
          <LayoutTemplate>
              <table id="Table2" runat="server">
                  <tr id="Tr1" runat="server">
                      <td id="Td1" runat="server">
                          <table ID="itemPlaceholderContainer" runat="server" border="1" 
                              style="background-color: #FFFFFF;border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;font-family: Verdana, Arial, Helvetica, sans-serif;">
                              <tr id="Tr2" runat="server" style="background-color:#DCDCDC; color: #000000;">
                                  <th id="Th1" runat="server">
                                      <asp:Button ID = "btnNew" runat="server" Text="New" CommandName = "new" />
                                  </th>
                                  <th id="Th2" runat="server">
                                      int_Batch_NO_group</th>
                                  <th id="Th3" runat="server">
                                      seq Per Group</th>
                                  <th id="Th4" runat="server">
                                      columnName</th>
                                  <th id="Th5" runat="server">
                                      relationOperator</th>
                                  <th id="Th6" runat="server">
                                      conditionValue</th>
                              </tr>
                              <tr runat="server" ID="itemPlaceholder">
                              </tr>
                          </table>
                      </td>
                  </tr>
                  <tr id="Tr3" runat="server">
                      <td id="Td2" runat="server" 
                          
                          
                          style="text-align: center;background-color: #CCCCCC; font-family: Verdana, Arial, Helvetica, sans-serif;color: #000000;">
                      </td>
                  </tr>
              </table>
          </LayoutTemplate>
            <SelectedItemTemplate>
                <tr style="background-color:#008A8C; font-weight: bold;color: #FFFFFF;">
                    <td>
                        <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                        <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                    </td>
                    <td>
                        <asp:Label ID="int_Batch_NO_groupLabel" runat="server" 
                            Text='<%# Eval("int_Batch_NO_group") %>' />
                    </td>
                    <td>
                        <asp:Label ID="sequencePerBatchNoGroupLabel" runat="server" 
                            Text='<%# Eval("sequencePerBatchNoGroup") %>' />
                    </td>
                    <td>
                        <asp:Label ID="columnNameLabel" runat="server" 
                            Text='<%# Eval("columnName") %>' />
                    </td>
                    <td>
                        <asp:Label ID="relationOperatorLabel" runat="server" 
                            Text='<%# Eval("relationOperator") %>' />
                    </td>
                    <td>
                        <asp:Label ID="conditionValueLabel" runat="server" 
                            Text='<%# Eval("conditionValue") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\param.accdb;Persist Security Info=True" 
        ProviderName="System.Data.OleDb"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT * FROM [Esch_Na_tbl_production_lines_and_batch_no_group]" 
    DeleteCommand="DELETE FROM [Esch_Na_tbl_production_lines_and_batch_no_group] WHERE ([int_Batch_NO_group] = ?) AND ([sequencePerBatchNoGroup] = ?)" 
                                  
            UpdateCommand="UPDATE [Esch_Na_tbl_production_lines_and_batch_no_group] SET [columnName] = ?, [relationOperator] = ?, [conditionValue] = ? WHERE ([int_Batch_NO_group] = ?) AND ([sequencePerBatchNoGroup] = ?)" 
                                                                          
            
            
            
            InsertCommand="INSERT INTO [Esch_Na_tbl_production_lines_and_batch_no_group] ([int_Batch_NO_group], [sequencePerBatchNoGroup], [columnName], [relationOperator], [conditionValue]) VALUES (?, ?, ?, ?, ?)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_int_Batch_NO_group" Type="Int16" />
            <asp:Parameter Name="original_sequencePerBatchNoGroup" Type="Int16" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="int_Batch_NO_group" Type="Int16" />
            <asp:Parameter Name="sequencePerBatchNoGroup" Type="Int16" />
            <asp:Parameter Name="columnName" Type="String" />
            <asp:Parameter Name="relationOperator" Type="String" />
            <asp:Parameter Name="conditionValue" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="columnName" Type="String" />
            <asp:Parameter Name="relationOperator" Type="String" />
            <asp:Parameter Name="conditionValue" Type="String" />
            <asp:Parameter Name="original_int_Batch_NO_group" Type="Int16" />
            <asp:Parameter Name="original_sequencePerBatchNoGroup" Type="Int16" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

