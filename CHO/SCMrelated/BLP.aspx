<%@ Page  aspcompat="true"  Title="BLP" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="BLP.aspx.vb" Inherits="SCMrelated_BLP" Theme="Monochrome" %>

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
          DataSourceID="SDS1" DataKeyNames="groupName,seqPerGroup" 
            ViewStateMode="Enabled" >
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC;">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="groupNameLabel" runat="server" 
                          Text='<%# Eval("groupName") %>' />
                  </td>
                  <td>
                      <asp:Label ID="seqPerGroupLabel" runat="server" 
                          Text='<%# Eval("seqPerGroup") %>' />
                  </td>
                  <td>
                      <asp:Label ID="headerNameLabel" runat="server" 
                          Text='<%# Eval("headerName") %>' />
                  </td>
                  <td>
                      <asp:Label ID="relationOperatorLabel" runat="server" 
                          Text='<%# Eval("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:Label ID="conditionValueLabel" runat="server" 
                          Text='<%# Eval("conditionValue") %>' />
                  </td>
                  <td>
                      <asp:Label ID="changeToCurrencyLabel" runat="server" 
                          Text='<%# Eval("changeToCurrency")%>' />
                  </td>
              </tr>
          </AlternatingItemTemplate>
          <EditItemTemplate>
              <tr style="background-color:#008A8C;color: #FFFFFF;">
                  <td>
                      <asp:Button ID="UpdateButton" runat="server" CommandName="Update" 
                          Text="Update" />
                      <asp:Button ID="CancelButton" runat="server" CommandName="Cancel" 
                          Text="Cancel" />
                  </td>
                  <td>
                      <asp:Label ID="groupNameLabel1" runat="server" 
                          Text='<%# Eval("groupName") %>' />
                  </td>
                  <td>
                       <asp:Label ID="seqPerGroupLabel1" runat="server" 
                           Text='<%# Eval("seqPerGroup") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="headerNameTextBox" runat="server" 
                          Text='<%# Bind("headerName") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="relationOperatorTextBox" runat="server" 
                          Text='<%# Bind("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="conditionValueTextBox" runat="server" 
                          Text='<%# Bind("conditionValue") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="changeToCurrencyTextBox" runat="server" 
                          Text='<%# Bind("changeToCurrency")%>' />
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
                      <asp:TextBox ID="groupNameTextBox" runat="server" 
                          Text='<%# Bind("groupName") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="seqPerGroupTextBox" runat="server" 
                          Text='<%# Bind("seqPerGroup") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="headerNameTextBox" runat="server" 
                          Text='<%# Bind("headerName") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="relationOperatorTextBox" runat="server" 
                          Text='<%# Bind("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="conditionValueTextBox" runat="server" 
                          Text='<%# Bind("conditionValue") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="changeToCurrencyTextBox" runat="server" 
                          Text='<%# Bind("changeToCurrency")%>' />
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
                      <asp:Label ID="groupNameLabel" runat="server" Text='<%# Eval("groupName") %>' />
                  </td>
                  <td>
                      <asp:Label ID="seqPerGroupLabel" runat="server" 
                          Text='<%# Eval("seqPerGroup") %>' />
                  </td>
                  <td>
                      <asp:Label ID="headerNameLabel" runat="server" 
                          Text='<%# Eval("headerName") %>' />
                  </td>
                  <td>
                      <asp:Label ID="relationOperatorLabel" runat="server" 
                          Text='<%# Eval("relationOperator") %>' />
                  </td>
                  <td>
                      <asp:Label ID="conditionValueLabel" runat="server" 
                          Text='<%# Eval("conditionValue") %>' />
                  </td>
                  <td>
                      <asp:Label ID="changeToCurrencyLabel" runat="server" 
                          Text='<%# Eval("changeToCurrency")%>' />
                  </td>
              </tr>
          </ItemTemplate>
          <LayoutTemplate>
              <table runat="server">
                  <tr runat="server">
                      <td runat="server">
                          <table ID="itemPlaceholderContainer" runat="server" border="1" 
                              style="background-color: #FFFFFF;border-collapse: collapse;border-color: #999999;border-style:none;border-width:1px;font-family: Verdana, Arial, Helvetica, sans-serif;">
                              <tr runat="server" style="background-color:#DCDCDC;color: #000000;">
                                  <th runat="server">
                                  <asp:Button ID = "btnNew" runat="server" Text="New" CommandName = "new" />
                                  </th>
                                  <th runat="server">
                                      groupName</th>
                                  <th runat="server">
                                      seqPerGroup</th>
                                  <th runat="server">
                                      headerName</th>
                                  <th runat="server">
                                      relationOperator</th>
                                  <th runat="server">
                                      conditionValue</th>
                                  <th runat="server">
                                      changeToCurrency</th>
                              </tr>
                              <tr runat="server" ID="itemPlaceholder">
                              </tr>
                          </table>
                      </td>
                  </tr>
                  <tr runat="server">
                      <td runat="server" 
                          style="text-align: center;background-color: #CCCCCC;font-family: Verdana, Arial, Helvetica, sans-serif;color: #000000;">
                      </td>
                  </tr>
              </table>
          </LayoutTemplate>
            <SelectedItemTemplate>
                <tr style="background-color:#008A8C;font-weight: bold;color: #FFFFFF;">
                    <td>
                        <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                            Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                        <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                    </td>
                    <td>
                        <asp:Label ID="groupNameLabel" runat="server" Text='<%# Eval("groupName") %>' />
                    </td>
                    <td>
                        <asp:Label ID="seqPerGroupLabel" runat="server" 
                            Text='<%# Eval("seqPerGroup") %>' />
                    </td>
                    <td>
                        <asp:Label ID="headerNameLabel" runat="server" 
                            Text='<%# Eval("headerName") %>' />
                    </td>
                    <td>
                        <asp:Label ID="relationOperatorLabel" runat="server" 
                            Text='<%# Eval("relationOperator") %>' />
                    </td>
                    <td>
                        <asp:Label ID="conditionValueLabel" runat="server" 
                            Text='<%# Eval("conditionValue") %>' />
                    </td>
                    <td>
                        <asp:Label ID="changeToCurrencyLabel" runat="server" 
                            Text='<%# Eval("changeToCurrency")%>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\inetpub\wwwroot\test\App_Data\param.accdb" 
        ProviderName="System.Data.SqlClient"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT * FROM [Esch_CQ_tbl_BLP_new] ORDER BY [groupName], [seqPerGroup]" 
    DeleteCommand="DELETE FROM [Esch_CQ_tbl_BLP_new] WHERE ([groupName] = ?) AND ([seqPerGroup] = ?)" 
                                  
            UpdateCommand="UPDATE [Esch_CQ_tbl_BLP_new] SET [headerName] = ?, [relationOperator] = ?, [conditionValue] = ?, [changeToCurrency] = ? WHERE ([groupName] = ?) AND ([seqPerGroup] = ?)" 
            
            
            InsertCommand="INSERT INTO [Esch_CQ_tbl_BLP_new] ([groupName], [seqPerGroup], [headerName], [relationOperator], [conditionValue], [changeToCurrency]) VALUES (?, ?, ?, ?, ?, ?)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_groupName" Type="String" />
            <asp:Parameter Name="original_seqPerGroup" Type="Int16" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="groupName" Type="String" />
            <asp:Parameter Name="seqPerGroup" Type="Int16" />
            <asp:Parameter Name="headerName" Type="String" />
            <asp:Parameter Name="relationOperator" Type="String" />
            <asp:Parameter Name="conditionValue" Type="String" />
            <asp:Parameter Name="changeToCurrency" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="headerName" Type="String" />
            <asp:Parameter Name="relationOperator" Type="String" />
            <asp:Parameter Name="conditionValue" Type="String" />
            <asp:Parameter Name="changeToCurrency" Type="String" />
            <asp:Parameter Name="original_groupName" Type="String" />
            <asp:Parameter Name="original_seqPerGroup" Type="Int16" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

