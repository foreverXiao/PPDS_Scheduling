<%@ Page  aspcompat="true"  Title="Current lead time" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="currentLeadtime.aspx.vb" Inherits="SCMrelated_currentLeadtime" Theme="Monochrome" %>

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
        Enabled="False"  />
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
          DataSourceID="SDS1" DataKeyNames="Lead_time_Group,Production_line" 
            ViewStateMode="Enabled" >
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC; ">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="Lead_time_GroupLabel" runat="server" 
                          Text='<%# Eval("Lead_time_Group") %>' />
                  </td>
                  <td>
                      <asp:Label ID="Production_lineLabel" runat="server" 
                          Text='<%# Eval("Production_line") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lngDaily_outputLabel" runat="server" 
                          Text='<%# Eval("lngDaily_output") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lnglead_timeLabel" runat="server" 
                          Text='<%# Eval("lnglead_time") %>' />
                  </td>
                  <td>
                      <asp:Label ID="datStartTimeLabel" runat="server" 
                          Text='<%# Eval("datStartTime") %>' />
                  </td>
                  <td>
                      <asp:Label ID="sng_coefficientLabel" runat="server" 
                          Text='<%# Eval("sng_coefficient") %>' />
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
                      <asp:Label ID="Lead_time_GroupLabel1" runat="server" 
                          Text='<%# Eval("Lead_time_Group") %>' />
                  </td>
                  <td>
                       <asp:Label ID="Production_lineLabel1" runat="server" 
                           Text='<%# Eval("Production_line") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lngDaily_outputTextBox" runat="server" 
                          Text='<%# Bind("lngDaily_output") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lnglead_timeTextBox" runat="server" 
                          Text='<%# Bind("lnglead_time") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="datStartTimeTextBox" runat="server" 
                          Text='<%# Bind("datStartTime") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="sng_coefficientTextBox" runat="server" 
                          Text='<%# Bind("sng_coefficient") %>' />
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
                      <asp:TextBox ID="Lead_time_GroupTextBox" runat="server" 
                          Text='<%# Bind("Lead_time_Group") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="Production_lineTextBox" runat="server" 
                          Text='<%# Bind("Production_line") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lngDaily_outputTextBox" runat="server" 
                          Text='<%# Bind("lngDaily_output") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lnglead_timeTextBox" runat="server" 
                          Text='<%# Bind("lnglead_time") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="datStartTimeTextBox" runat="server" 
                          Text='<%# Bind("datStartTime") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="sng_coefficientTextBox" runat="server" 
                          Text='<%# Bind("sng_coefficient") %>' />
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
                      <asp:Label ID="Lead_time_GroupLabel" runat="server" 
                          Text='<%# Eval("Lead_time_Group") %>' />
                  </td>
                  <td>
                      <asp:Label ID="Production_lineLabel" runat="server" 
                          Text='<%# Eval("Production_line") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lngDaily_outputLabel" runat="server" 
                          Text='<%# Eval("lngDaily_output") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lnglead_timeLabel" runat="server" 
                          Text='<%# Eval("lnglead_time") %>' />
                  </td>
                  <td>
                      <asp:Label ID="datStartTimeLabel" runat="server" 
                          Text='<%# Eval("datStartTime") %>' />
                  </td>
                  <td>
                      <asp:Label ID="sng_coefficientLabel" runat="server" 
                          Text='<%# Eval("sng_coefficient") %>' />
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
                                  <th id="Th1" runat="server">
                                  <asp:Button ID = "btnNew" runat="server" Text="New" CommandName = "new" />
                                  </th>
                                  <th runat="server">
                                      Lead_time_Group</th>
                                  <th runat="server">
                                      Production_line</th>
                                  <th runat="server">
                                      lngDaily_output</th>
                                  <th runat="server">
                                      lnglead_time</th>
                                  <th runat="server">
                                      datStartTime</th>
                                  <th runat="server">
                                      sng_coefficient</th>
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
                        <asp:Label ID="Lead_time_GroupLabel" runat="server" 
                            Text='<%# Eval("Lead_time_Group") %>' />
                    </td>
                    <td>
                        <asp:Label ID="Production_lineLabel" runat="server" 
                            Text='<%# Eval("Production_line") %>' />
                    </td>
                    <td>
                        <asp:Label ID="lngDaily_outputLabel" runat="server" 
                            Text='<%# Eval("lngDaily_output") %>' />
                    </td>
                    <td>
                        <asp:Label ID="lnglead_timeLabel" runat="server" 
                            Text='<%# Eval("lnglead_time") %>' />
                    </td>
                    <td>
                        <asp:Label ID="datStartTimeLabel" runat="server" 
                            Text='<%# Eval("datStartTime") %>' />
                    </td>
                    <td>
                        <asp:Label ID="sng_coefficientLabel" runat="server" 
                            Text='<%# Eval("sng_coefficient") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\IIS\Test\App_Data\db_Resin.mdb" 
        ProviderName="System.Data.OleDb"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT [Lead_time_Group], [Production_line], [lngDaily_output], [lnglead_time], [datStartTime], [sng_coefficient] FROM [Esch_Na_tbl_Lead_Time]" 
    DeleteCommand="DELETE FROM [Esch_Na_tbl_Lead_Time] WHERE ([Lead_time_Group] = ?) AND ([Production_line] = ?)" 
                                  
            UpdateCommand="UPDATE [Esch_Na_tbl_Lead_Time] SET [lngDaily_output] = ?, [lnglead_time] = ?, [datStartTime] = ?, [sng_coefficient] = ? WHERE ([Lead_time_Group] = ?) AND ([Production_line] = ?)" 
                                           
            
            
            
            
            InsertCommand="INSERT INTO [Esch_Na_tbl_Lead_Time] ([Lead_time_Group], [Production_line], [lngDaily_output], [lnglead_time], [datStartTime], [sng_coefficient]) VALUES (?, ?, ?, ?, ?, ?)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_Lead_time_Group" Type="String" />
            <asp:Parameter Name="original_Production_line" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="Lead_time_Group" Type="String" />
            <asp:Parameter Name="Production_line" Type="String" />
            <asp:Parameter Name="lngDaily_output" Type="Int32" />
            <asp:Parameter Name="lnglead_time" Type="Int32" />
            <asp:Parameter Name="datStartTime" Type="DateTime" />
            <asp:Parameter Name="sng_coefficient" Type="Single" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="lngDaily_output" Type="Int32" />
            <asp:Parameter Name="lnglead_time" Type="Int32" />
            <asp:Parameter Name="datStartTime" Type="DateTime" />
            <asp:Parameter Name="sng_coefficient" Type="Single" />
            <asp:Parameter Name="original_Lead_time_Group" Type="String" />
            <asp:Parameter Name="original_Production_line" Type="String" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

