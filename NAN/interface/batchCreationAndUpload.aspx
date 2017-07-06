<%@ Page   Title="Batch creation and upload" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="batchCreationAndUpload.aspx.vb" Inherits="interface_batchCreationAndUpload" Theme="Monochrome" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="act" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
    <asp:LinkButton ID="listEDI" runat="server" style="position:absolute;right:150px;">List EDI files</asp:LinkButton>
    <asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy><br />
    <asp:UpdatePanel ID="UP3" runat="server"><ContentTemplate>
        <asp:CheckBox ID="cbAllLines"
            runat="server" CssClass="MyButton" Text=" All" 
            AutoPostBack="True" Checked="True" />&nbsp;&nbsp;<asp:Literal
        ID="Literal1" runat="server">  Production Lines:</asp:Literal>
    <asp:PlaceHolder
            ID="linesCollection" runat="server" ClientIDMode="Static"></asp:PlaceHolder><br /><br />
        <asp:DropDownList ID="ddlColumn" runat="server" ClientIDMode="Static" 
            ViewStateMode="Enabled">
            <asp:ListItem>dat_start_date</asp:ListItem>
            <asp:ListItem>dat_finish_date</asp:ListItem>
    </asp:DropDownList> &nbsp;Between&nbsp;&nbsp;  
        <asp:TextBox ID="earlierTime" runat="server"  Width="72px"></asp:TextBox><asp:Image ID="Img1" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <act:CalendarExtender ID="CE1" runat="server" 
        TargetControlID="earlierTime" PopupButtonID="Img1">
    </act:CalendarExtender> Hour:
        <asp:DropDownList ID="ddlHour1" runat="server">
        </asp:DropDownList> Minute:
        <asp:DropDownList ID="ddlMinute1" runat="server">
        </asp:DropDownList> &nbsp;&nbsp;And&nbsp;&nbsp; <asp:TextBox ID="laterTime" runat="server" Width="72px"></asp:TextBox><asp:Image ID="Img2" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <act:CalendarExtender ID="CE2" runat="server" 
        TargetControlID="laterTime" PopupButtonID="Img2">
    </act:CalendarExtender> Hour:
        <asp:DropDownList ID="ddlHour2" runat="server">
        </asp:DropDownList> Minute:
        <asp:DropDownList ID="ddlMinute2" runat="server">
        </asp:DropDownList><br />
         <hr style="color:white;height:1px" />
        <asp:Button ID="prdctnOrdrs" runat="server" 
            Text="Prepare production orders list" 
            ToolTip="sort out all the orders to be created batch no. based on the selected production lines and timing" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" />
         <asp:Label ID="StatusLabel" runat="server" ></asp:Label>
         </ContentTemplate></asp:UpdatePanel>
      <hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP1" runat="server" ClientIDMode="Static"><ContentTemplate>
        <asp:CheckBox ID="cbAllOrders"
            runat="server" CssClass="MyButton" Text=" All orders" 
            AutoPostBack="True" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="CreateBatch" runat="server" Text="Create Batch No." OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" />
        <asp:Button ID="gnrtFlndFTP" runat="server" Text="Generate EDI file and FTP to OPM server" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" />
    </ContentTemplate></asp:UpdatePanel>
    <hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP2" runat="server" ClientIDMode="Static"><ContentTemplate>
    <asp:Label ID="Message"
        ForeColor="Red"          
        runat="server"/><br />
      <asp:repeater id="Rt1"       
        datasourceid="SDS1"
        runat="server" ClientIDMode="Static">

        <headertemplate>
          <table border="1">
            <tr>
              <td><b>Create<br />Lot ?<br />(Y/N)</b></td>
              <td><b>Line</b></td>
              <td><b>Lot No</b></td>
              <td><b>Item No</b></td>
              <td><b>Quantity<br />(KG)</b></td>
              <td><b>Line<br />Group</b></td>
              <td><b>Currency</b></td>
              <td><b>Package</b></td>
              <td><b>FTP<br />New<br />Lot</b></td>
              <td><b>Starting</b></td>
              <td><b>Ending</b></td>
              <td><b>Formula<br />Version</b></td>
              <td><b>End user</b></td>
              <td><b>POD</b></td>
              <td><b>Ship to</b></td>
              <td><b>Order Key</b></td>
            </tr>
        </headertemplate>

        <itemtemplate>
          <tr><td><asp:CheckBox ID="Y_N" runat="server" Checked = '<%# Eval("bnl_Y_or_N")%>'  /></td><td> <%# Eval("int_line_no") %> </td><td><asp:TextBox ID="L_N" runat="server" Text = '<%# Eval("txt_lot_no") %>'  Width="72px" /></td><td> <%# Eval("txt_item_no") %> </td><td> <%# Eval("planned_production_qty")%> </td><td> <%# Eval("txt_line_group")%> </td><td> <%# Eval("txt_currency")%> </td><td> <%# Eval("txt_package_code")%> </td><td><asp:CheckBox ID="N_B" runat="server" Checked = '<%# Eval("bnlNewBatchNO")%>'  /> </td><td><small> <%# Eval("dat_start_date")%> </small></td><td><small> <%# Eval("dat_finish_date")%> </small></td><td> <%# Eval("txt_formula_version")%> </td><td> <%# Eval("txt_end_user")%> </td><td> <%# Eval("txt_destination")%> </td><td> <%# Eval("txt_ship_cust_no")%> </td><td><asp:Label ID="O_K" runat="server" Text = '<%# Eval("txt_order_key") %>' /></td></tr>
        </itemtemplate>

        <footertemplate>
          </table>
        </footertemplate>
      </asp:repeater><br />  
        <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\IIS\Test\App_Data\db_Resin.mdb" 
        ProviderName="System.Data.OleDb"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT * FROM [Esch_Na_tbl_BatchNO] ORDER BY int_line_no ASC,dat_start_date ASC,dat_finish_date,txt_item_no,txt_currency ASC,txt_package_code,txt_order_key ASC" >
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="cbAllOrders" EventName="CheckedChanged" />
    <asp:AsyncPostBackTrigger ControlID="CreateBatch" EventName="Click" />
    </Triggers></asp:UpdatePanel>
</asp:Content>

