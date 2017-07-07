<%@ Page   Title="Combine production schedule " Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="combineProduction.aspx.vb" Inherits="dragDrop_normalOP_combineProduction" Theme="Monochrome" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
    <asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy><br />
    <asp:UpdatePanel ID="UP3" runat="server"><ContentTemplate>
        <asp:CheckBox ID="cbAllLines"
            runat="server" CssClass="MyButton" Text=" All" 
            AutoPostBack="True" Checked="True" />&nbsp;&nbsp;<asp:Literal
        ID="Literal1" runat="server">  Production Lines:</asp:Literal>
    <asp:PlaceHolder
            ID="linesCollection" runat="server"></asp:PlaceHolder><br /><br />
        <b>Orders'</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:DropDownList ID="ddlColumn" runat="server" ClientIDMode="Static" 
            ViewStateMode="Enabled">
            <asp:ListItem>dat_start_date</asp:ListItem>
            <asp:ListItem>dat_finish_date</asp:ListItem>
    </asp:DropDownList> &nbsp;Between&nbsp;&nbsp;  
        <asp:TextBox ID="earlierTime" runat="server"  Width="72px"></asp:TextBox><asp:Image ID="Img1" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <asp:CalendarExtender ID="CE1" runat="server" 
        TargetControlID="earlierTime" PopupButtonID="Img1">
    </asp:CalendarExtender> Hour:
        <asp:DropDownList ID="ddlHour1" runat="server">
        </asp:DropDownList> Minute:
        <asp:DropDownList ID="ddlMinute1" runat="server">
        </asp:DropDownList> &nbsp;&nbsp;And&nbsp;&nbsp; <asp:TextBox ID="laterTime" runat="server" Width="72px"></asp:TextBox><asp:Image ID="Img2" ImageUrl="~/App_Themes/Monochrome/Images/Calendar.png"  runat="server" />
        <asp:CalendarExtender ID="CE2" runat="server" 
        TargetControlID="laterTime" PopupButtonID="Img2">
    </asp:CalendarExtender> Hour:
        <asp:DropDownList ID="ddlHour2" runat="server">
        </asp:DropDownList> Minute:
        <asp:DropDownList ID="ddlMinute2" runat="server">
        </asp:DropDownList><br /><hr style="color:white;height:1px" /><b> or order in dummy line but its quantity is 
        greater than 0 or order in special lines but its ETD is in between the above two 
        dates <br />(filter out those items when none of their int_order_status does like 'NEW' or 'REV-R' or 'REV-Q')==><asp:CheckBox 
            ID="fltout" runat="server" Checked="True" Text="Apply this rule" 
            TextAlign="Left" /></b>&nbsp;&nbsp;
            <asp:Button ID="prdctnOrdrs" runat="server" 
            Text="List orders for combination" 
            ToolTip="sort out all the orders to be created batch no. based on the selected production lines and timing" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" />
         <hr style="color:white;height:1px" />
         <b>The gap of RSDs among orders is less than or equal to</b>&nbsp;<asp:TextBox 
            ID="gapRSD" runat="server"  Width="60px"></asp:TextBox>&nbsp;&nbsp;<asp:Button ID="prdctnOrdrs2" runat="server" 
            Text="List orders for combination" 
            ToolTip="sort out all the orders to be created batch no. based on the selected production lines and timing" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" />
        <hr style="color:white;height:1px" />
         <asp:Label ID="StatusLabel" runat="server" ></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
         <asp:Button ID="bTdb" runat="server" Text="Write updated data back to order table" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" />
         </ContentTemplate><Triggers></Triggers></asp:UpdatePanel>
    <hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP2" runat="server" ClientIDMode="Static"><ContentTemplate>
    <asp:Label ID="Message"
        ForeColor="Red"          
        runat="server"/>
      <asp:repeater id="Rt1"       
        datasourceid="SDS1"
        runat="server" ClientIDMode="Static" >

        <headertemplate>
          <table  style="background-color:silver; border-width:0px;">
            <tr>
              <td><b>Item No</b></td>
              <td><b>start</b></td>
              <td><b>ETD</b></td>
              <td><b>Line</b></td>
              <td><b>Quantity</b></td>
              <td><b>Lot No</b></td>
              <td><b>order<br />status</b></td>
              <td><b>Currency</b></td>
              <td><b>RDD</b></td>
              <td><b>remark</b></td>
              <td><b>Order<br />Key</b></td>
            </tr>
        </headertemplate>

        <itemtemplate>
          <tr>
            <td nowrap="nowrap" style="color: #0000FF;font-size:x-large;"><%# Eval("txt_item_no")%></td>
            <td><asp:TextBox ID="strt" runat="server" Text = '<%# String.Format("{0:M/d/yyyy HH:mm}",Eval("dat_start_date")) %>'  Width="102px" ClientIDMode="Inherit" BackColor="silver"  ForeColor="blue"  BorderStyle="Groove" /></td>
            <td> <%# String.Format("{0:M/d/yyyy}", Eval("dat_etd"))%> </td>
            <td><asp:TextBox ID="line" runat="server" Text = '<%# Eval("int_line_no") %>'  Width="32px" BackColor="silver"  ForeColor="blue" BorderStyle="Groove" /></td>
            <td><asp:TextBox ID="qty" runat="server" Text = '<%# Eval("planned_production_qty") %>'  Width="42px" BackColor="silver"  ForeColor="blue" BorderStyle="Groove" /></td>
            <td><asp:TextBox ID="Lot" runat="server" Text = '<%# Eval("txt_lot_no") %>'  Width="62px" BackColor="silver"  ForeColor="blue" BorderStyle="Groove" /></td>
            <td> <%# Eval("int_status_key")%> </td>
            <td> <%# Eval("txt_currency")%> </td>         
            <td nowrap="nowrap"> <%# String.Format("{0:M/d/yyyy}",Eval("dat_rdd")) %> </td>         
            <td><asp:TextBox ID="rmk" runat="server" Text = '<%# Eval("txt_remark") %>'  Width="320px" BackColor="silver"  ForeColor="blue" BorderStyle="Groove" /></td>                       
            <td><asp:Label ID="O_K" runat="server" Text = '<%# Eval("txt_order_key") %>' /></td>
          </tr>
        </itemtemplate>

        <footertemplate>
          </table>
        </footertemplate>
      </asp:repeater><br />  
        <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.Jet.OleDb.4.0;Data Source=C:\IIS\Test\App_Data\db_Resin.mdb" 
        ProviderName="System.Data.SqlClient"
        OldValuesParameterFormatString="original_{0}"
        
    
            SelectCommand="SELECT [txt_item_no],[dat_etd], [planned_production_qty], [int_line_no], [txt_lot_no], [dat_start_date], [dat_finish_date],[int_status_key],[txt_remark], [txt_currency],[dat_rdd],  [txt_order_key] FROM [Esch_Na_tbl_similar_item_combination] ORDER BY [txt_item_no_similar] ASC,[txt_item_no] ASC, [dat_start_date] ASC, [int_line_no] ASC" >
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    </Triggers></asp:UpdatePanel>
</asp:Content>

