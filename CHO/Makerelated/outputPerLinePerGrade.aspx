<%@ Page  aspcompat="true"  Title="Output per line and grade" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="outputPerLinePerGrade.aspx.vb" Inherits="Makerelated_outputPerLinePerGrade" Theme="Monochrome" %>

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
          DataSourceID="SDS1" DataKeyNames="txt_grade,txt_line_no" 
            ViewStateMode="Enabled" >
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC;">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');" />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="txt_gradeLabel" runat="server" 
                          Text='<%# Eval("txt_grade") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_rateLabel" runat="server" 
                          Text='<%# Eval("int_rate") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txtRemarkLabel" runat="server" 
                          Text='<%# Eval("txtRemark") %>' />
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
                      <asp:Label ID="txt_gradeLabel1" runat="server" 
                          Text='<%# Eval("txt_grade") %>' />
                  </td>
                  <td>
                       <asp:Label ID="txt_line_noLabel1" runat="server" 
                           Text='<%# Eval("txt_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_rateTextBox" runat="server" 
                          Text='<%# Bind("int_rate") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txtRemarkTextBox" runat="server" 
                          Text='<%# Bind("txtRemark") %>' />
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
                      <asp:TextBox ID="txt_gradeTextBox" runat="server" 
                          Text='<%# Bind("txt_grade") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_line_noTextBox" runat="server" 
                          Text='<%# Bind("txt_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_rateTextBox" runat="server" 
                          Text='<%# Bind("int_rate") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txtRemarkTextBox" runat="server" 
                          Text='<%# Bind("txtRemark") %>' />
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
                      <asp:Label ID="txt_gradeLabel" runat="server" 
                          Text='<%# Eval("txt_grade") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_rateLabel" runat="server" 
                          Text='<%# Eval("int_rate") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txtRemarkLabel" runat="server" 
                          Text='<%# Eval("txtRemark") %>' />
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
                                      txt_grade</th>
                                  <th runat="server">
                                      txt_line_no</th>
                                  <th runat="server">
                                      int_rate</th>
                                  <th runat="server">
                                      txtRemark</th>
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
                        <asp:Label ID="txt_gradeLabel" runat="server" 
                            Text='<%# Eval("txt_grade") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_line_noLabel" runat="server" 
                            Text='<%# Eval("txt_line_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="int_rateLabel" runat="server" 
                            Text='<%# Eval("int_rate") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txtRemarkLabel" runat="server" 
                            Text='<%# Eval("txtRemark") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\IIS\Test\App_Data\db_Resin.mdb" 
        ProviderName="System.Data.SqlClient"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT [txt_grade], [txt_line_no], [int_rate], [txtRemark] FROM [Esch_CQ_tbl_output_by_line_by_grade]" 
    DeleteCommand="DELETE FROM [Esch_CQ_tbl_output_by_line_by_grade] WHERE ([txt_grade] = ?) AND ([txt_line_no] = ?)" 
                                  
            UpdateCommand="UPDATE [Esch_CQ_tbl_output_by_line_by_grade] SET [int_rate] = ?, [txtRemark] = ? WHERE ([txt_grade] = ?) AND ([txt_line_no] = ?)" 
            
            
            
            InsertCommand="INSERT INTO [Esch_CQ_tbl_output_by_line_by_grade] ([txt_grade], [txt_line_no], [int_rate], [txtRemark]) VALUES (?, ?, ?, ?)" >
  
        <DeleteParameters>
            <asp:Parameter Name="original_txt_grade" Type="String" />
            <asp:Parameter Name="original_txt_line_no" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="txt_grade" Type="String" />
            <asp:Parameter Name="txt_line_no" Type="String" />
            <asp:Parameter Name="int_rate" Type="Double" />
            <asp:Parameter Name="txtRemark" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="int_rate" Type="Double" />
            <asp:Parameter Name="txtRemark" Type="String" />
            <asp:Parameter Name="original_txt_grade" Type="String" />
            <asp:Parameter Name="original_txt_line_no" Type="String" />
        </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

