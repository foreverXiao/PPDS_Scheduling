﻿<%@ Page  aspcompat="true"  Title="Order detail" Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="OrderDetail2.aspx.vb" Inherits="dragDrop_OrderDetail" Theme="Monochrome" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CP1" Runat="Server">
    <asp:ScriptManagerProxy ID="SMP1" runat="server">
    </asp:ScriptManagerProxy>
<br />
      <asp:FileUpload ID="FileUpload1" runat="server" 
    style="margin-bottom: 0px"  ClientIDMode="Static" Width="196px" />&nbsp;&nbsp;
    <asp:Button runat="server" 
        id="UpldUpdate" text="Update" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False"  OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};"  />
    <asp:Button runat="server" 
        id="UpldDel" text="Delete" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};"  />&nbsp;
    <asp:Button runat="server" 
        id="UpldInsrt" text="Insert" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False" OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};"  />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="overwrite" 
        style="left:80px;" runat="server" Text="Upload and overwrite" 
        OnClientClick="if (this.value.indexOf('...') > 0 ){this.disabled=true;}else{this.value +='...';};" 
        ClientIDMode="Static"  />
&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="Download1" runat="server" 
        Text="Download current selection" style="left:100px;" ClientIDMode="Static"  />
    <asp:HiddenField ID="hiddenBT"  runat="server" ClientIDMode="Static"  /><br />
        <hr style="color:white;height:1px" />
      <asp:Label ID="StatusLabel" runat="server" Text=""></asp:Label>
    <asp:LinkButton ID="exceptionR"
          runat="server" Visible="False" Font-Bold="True" ForeColor="Red" 
        ClientIDMode="Static">Click to view exception report</asp:LinkButton>
      <hr style="color:white;height:1px" />
    <asp:UpdatePanel ID="UP1" runat="server" ClientIDMode="Static"><ContentTemplate>
    <asp:DropDownList ID="DDL1" runat="server" ClientIDMode="Static" 
        AutoPostBack="True" ViewStateMode="Enabled">
    </asp:DropDownList><asp:DropDownList ID="DDL2" runat="server" ClientIDMode="Static">
    </asp:DropDownList>
    <asp:TextBox ID="filtercdtn1" runat="server" ClientIDMode="Static" Width="120px" 
            AutoPostBack="False"  ></asp:TextBox>
    <asp:Button ID="Filter1" runat="server" Text="Filter" ClientIDMode="Static"  />
        <asp:Button ID="clrfltr1"
        runat="server" Text="Clear Filter" ClientIDMode="Static" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:LinkButton ID="cmbnPrdctn" runat="server" Text="Combine production" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="mtiOE" runat="server" Text="MTI order entry" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="btchCrtn" runat="server" Text="Go to create Batch No." />
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
          DataSourceID="SDS1" DataKeyNames="txt_order_key">
          <AlternatingItemTemplate>
              <tr style="background-color:#FFF8DC;" >
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');"  />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_keyLabel" runat="server" 
                          Text='<%# Eval("txt_order_key") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_status_keyLabel" runat="server" 
                          Text='<%# Eval("int_status_key") %>' />
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
                      <asp:Label ID="txt_item_noLabel" runat="server" 
                          Text='<%# Eval("txt_item_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_order_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_order_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="planned_production_qtyLabel" runat="server" 
                          Text='<%# Eval("planned_production_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_etdLabel" runat="server" 
                          Text='<%# Eval("dat_etd") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_line_noLabel" runat="server" 
                          Text='<%# Eval("int_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_lot_noLabel" runat="server" 
                          Text='<%# Eval("txt_lot_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_start_dateLabel" runat="server" 
                          Text='<%# Eval("dat_start_date") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_finish_dateLabel" runat="server" 
                          Text='<%# Eval("dat_finish_date") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_new_explantLabel" runat="server" 
                          Text='<%# Eval("dat_new_explant") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_working_hoursLabel" runat="server" 
                          Text='<%# Eval("flt_working_hours") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_change_over_timeLabel" runat="server" 
                          Text='<%# Eval("int_change_over_time") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_completedLabel" runat="server" 
                          Text='<%# Eval("flt_actual_completed") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_rddLabel" runat="server" 
                          Text='<%# Eval("dat_rdd") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_currencyLabel" runat="server" 
                          Text='<%# Eval("txt_currency") %>' />
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
                      <asp:Label ID="txt_gl_classLabel" runat="server" 
                          Text='<%# Eval("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_gradeLabel" runat="server" 
                          Text='<%# Eval("txt_grade") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_colorLabel" runat="server" 
                          Text='<%# Eval("txt_color") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_line_assignLabel" runat="server" 
                          Text='<%# Eval("txt_line_assign") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_spanLabel" runat="server" 
                          Text='<%# Eval("int_span") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_typeLabel" runat="server" 
                          Text='<%# Eval("txt_order_type") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_orgn_codeLabel" runat="server" 
                          Text='<%# Eval("txt_orgn_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_region_codeLabel" runat="server" 
                          Text='<%# Eval("txt_region_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_order_addedLabel" runat="server" 
                          Text='<%# Eval("dat_order_added") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_rev_ex_plantLabel" runat="server" 
                          Text='<%# Eval("dat_rev_ex_plant") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_allocated_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_allocated_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_custLabel" runat="server" 
                          Text='<%# Eval("txt_ship_cust") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_clean_downLabel" runat="server" 
                          Text='<%# Eval("txt_clean_down") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_remarkLabel" runat="server" 
                          Text='<%# Eval("txt_remark") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_line_commentsLabel" runat="server" 
                          Text='<%# Eval("txt_line_comments") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_from_whseLabel" runat="server" 
                          Text='<%# Eval("txt_from_whse") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_cust_noLabel" runat="server" 
                          Text='<%# Eval("txt_ship_cust_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_market_segLabel" runat="server" 
                          Text='<%# Eval("txt_market_seg") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_molderLabel" runat="server" 
                          Text='<%# Eval("txt_molder") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_end_userLabel" runat="server" 
                          Text='<%# Eval("txt_end_user") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_sales_priceLabel" runat="server" 
                          Text='<%# Eval("flt_sales_price") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ess_soLabel" runat="server" 
                          Text='<%# Eval("txt_ess_so") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ess_sol_noLabel" runat="server" 
                          Text='<%# Eval("txt_ess_sol_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_allocated_lotsLabel" runat="server" 
                          Text='<%# Eval("txt_allocated_lots") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_payment_termLabel" runat="server" 
                          Text='<%# Eval("txt_payment_term") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_package_codeLabel" runat="server" 
                          Text='<%# Eval("txt_package_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_tbdLabel" runat="server" 
                          Text='<%# Eval("txt_tbd") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_uploadLabel" runat="server" 
                          Text='<%# Eval("txt_upload") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_batch_statusLabel" runat="server" 
                          Text='<%# Eval("txt_batch_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_actual_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_actual_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_actual_startLabel" runat="server" 
                          Text='<%# Eval("dat_actual_start") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_actual_finishLabel" runat="server" 
                          Text='<%# Eval("dat_actual_finish") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_actual_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_qty_manLabel" runat="server" 
                          Text='<%# Eval("flt_actual_qty_man") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_formula_versionLabel" runat="server" 
                          Text='<%# Eval("int_formula_version") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_VIPLabel" runat="server" 
                          Text='<%# Eval("txt_VIP") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_unallocate_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_unallocate_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_allocation_statusLabel" runat="server" 
                          Text='<%# Eval("txt_allocation_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lng_VIP_lead_timeLabel" runat="server" 
                          Text='<%# Eval("lng_VIP_lead_time") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lng_AdvanceOfRevisionLabel" runat="server" 
                          Text='<%# Eval("lng_AdvanceOfRevision") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_payment_statusLabel" runat="server" 
                          Text='<%# Eval("txt_payment_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_statusLabel" runat="server" 
                          Text='<%# Eval("txt_order_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_process_technicsLabel" runat="server" 
                          Text='<%# Eval("txt_process_technics") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_start_from_qaLabel" runat="server" 
                          Text='<%# Eval("dat_start_from_qa") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_qty_from_qaLabel" runat="server" 
                          Text='<%# Eval("flt_actual_qty_from_qa") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_FDALabel" runat="server" 
                          Text='<%# Eval("txt_FDA") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_remark_spclLabel" runat="server" 
                          Text='<%# Eval("txt_remark_spcl") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_auxiliary_codeLabel" runat="server" 
                          Text='<%# Eval("txt_auxiliary_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_auxiliary_code_for_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_auxiliary_code_for_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_marginLabel" runat="server" 
                          Text='<%# Eval("flt_margin") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_FromUserLabel" runat="server" 
                          Text='<%# Eval("txt_FromUser") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ToUserLabel" runat="server" 
                          Text='<%# Eval("txt_ToUser") %>' />
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
                      <asp:Label ID="txt_order_keyLabel1" runat="server" 
                          Text='<%# Eval("txt_order_key") %>' />
                  </td>
                  <td>
                       <asp:TextBox ID="int_status_keyTextBox" runat="server" 
                           Text='<%# Bind("int_status_key") %>' />
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
                      <asp:TextBox ID="txt_item_noTextBox" runat="server" 
                          Text='<%# Bind("txt_item_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_order_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_order_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="planned_production_qtyTextBox" runat="server" 
                          Text='<%# Bind("planned_production_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_etdTextBox" runat="server" 
                          Text='<%# Bind("dat_etd") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_line_noTextBox" runat="server" 
                          Text='<%# Bind("int_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_lot_noTextBox" runat="server" 
                          Text='<%# Bind("txt_lot_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_start_dateTextBox" runat="server" 
                          Text='<%# Bind("dat_start_date") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_finish_dateTextBox" runat="server" 
                          Text='<%# Bind("dat_finish_date") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_new_explantTextBox" runat="server" 
                          Text='<%# Bind("dat_new_explant") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_working_hoursTextBox" runat="server" 
                          Text='<%# Bind("flt_working_hours") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_change_over_timeTextBox" runat="server" 
                          Text='<%# Bind("int_change_over_time") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_completedTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_completed") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_rddTextBox" runat="server" 
                          Text='<%# Bind("dat_rdd") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_currencyTextBox" runat="server" 
                          Text='<%# Bind("txt_currency") %>' />
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
                      <asp:TextBox ID="txt_gl_classTextBox" runat="server" 
                          Text='<%# Bind("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_gradeTextBox" runat="server" 
                          Text='<%# Bind("txt_grade") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_colorTextBox" runat="server" 
                          Text='<%# Bind("txt_color") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_line_assignTextBox" runat="server" 
                          Text='<%# Bind("txt_line_assign") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_spanTextBox" runat="server" 
                          Text='<%# Bind("int_span") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_typeTextBox" runat="server" 
                          Text='<%# Bind("txt_order_type") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_orgn_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_orgn_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_region_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_region_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_order_addedTextBox" runat="server" 
                          Text='<%# Bind("dat_order_added") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_rev_ex_plantTextBox" runat="server" 
                          Text='<%# Bind("dat_rev_ex_plant") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_allocated_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_allocated_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_custTextBox" runat="server" 
                          Text='<%# Bind("txt_ship_cust") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_clean_downTextBox" runat="server" 
                          Text='<%# Bind("txt_clean_down") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_remarkTextBox" runat="server" 
                          Text='<%# Bind("txt_remark") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_line_commentsTextBox" runat="server" 
                          Text='<%# Bind("txt_line_comments") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_from_whseTextBox" runat="server" 
                          Text='<%# Bind("txt_from_whse") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_cust_noTextBox" runat="server" 
                          Text='<%# Bind("txt_ship_cust_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_market_segTextBox" runat="server" 
                          Text='<%# Bind("txt_market_seg") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_molderTextBox" runat="server" 
                          Text='<%# Bind("txt_molder") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_end_userTextBox" runat="server" 
                          Text='<%# Bind("txt_end_user") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_sales_priceTextBox" runat="server" 
                          Text='<%# Bind("flt_sales_price") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ess_soTextBox" runat="server" 
                          Text='<%# Bind("txt_ess_so") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ess_sol_noTextBox" runat="server" 
                          Text='<%# Bind("txt_ess_sol_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_allocated_lotsTextBox" runat="server" 
                          Text='<%# Bind("txt_allocated_lots") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_payment_termTextBox" runat="server" 
                          Text='<%# Bind("txt_payment_term") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_package_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_package_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_tbdTextBox" runat="server" 
                          Text='<%# Bind("txt_tbd") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_uploadTextBox" runat="server" 
                          Text='<%# Bind("txt_upload") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_batch_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_batch_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_actual_line_noTextBox" runat="server" 
                          Text='<%# Bind("txt_actual_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_actual_startTextBox" runat="server" 
                          Text='<%# Bind("dat_actual_start") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_actual_finishTextBox" runat="server" 
                          Text='<%# Bind("dat_actual_finish") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_qty_manTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_qty_man") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_formula_versionTextBox" runat="server" 
                          Text='<%# Bind("int_formula_version") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_VIPTextBox" runat="server" 
                          Text='<%# Bind("txt_VIP") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_unallocate_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_unallocate_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_allocation_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_allocation_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lng_VIP_lead_timeTextBox" runat="server" 
                          Text='<%# Bind("lng_VIP_lead_time") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lng_AdvanceOfRevisionTextBox" runat="server" 
                          Text='<%# Bind("lng_AdvanceOfRevision") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_payment_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_payment_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_order_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_process_technicsTextBox" runat="server" 
                          Text='<%# Bind("txt_process_technics") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_start_from_qaTextBox" runat="server" 
                          Text='<%# Bind("dat_start_from_qa") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_qty_from_qaTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_qty_from_qa") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_FDATextBox" runat="server" 
                          Text='<%# Bind("txt_FDA") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_remark_spclTextBox" runat="server" 
                          Text='<%# Bind("txt_remark_spcl") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_auxiliary_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_auxiliary_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_auxiliary_code_for_line_noTextBox" runat="server" 
                          Text='<%# Bind("txt_auxiliary_code_for_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_marginTextBox" runat="server" 
                          Text='<%# Bind("flt_margin") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_FromUserTextBox" runat="server" 
                          Text='<%# Bind("txt_FromUser") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ToUserTextBox" runat="server" 
                          Text='<%# Bind("txt_ToUser") %>' />
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
                      <asp:TextBox ID="txt_order_keyTextBox" runat="server" 
                          Text='<%# Bind("txt_order_key") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_status_keyTextBox" runat="server" 
                          Text='<%# Bind("int_status_key") %>' />
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
                      <asp:TextBox ID="txt_item_noTextBox" runat="server" 
                          Text='<%# Bind("txt_item_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_order_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_order_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="planned_production_qtyTextBox" runat="server" 
                          Text='<%# Bind("planned_production_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_etdTextBox" runat="server" Text='<%# Bind("dat_etd") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_line_noTextBox" runat="server" 
                          Text='<%# Bind("int_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_lot_noTextBox" runat="server" 
                          Text='<%# Bind("txt_lot_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_start_dateTextBox" runat="server" 
                          Text='<%# Bind("dat_start_date") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_finish_dateTextBox" runat="server" 
                          Text='<%# Bind("dat_finish_date") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_new_explantTextBox" runat="server" 
                          Text='<%# Bind("dat_new_explant") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_working_hoursTextBox" runat="server" 
                          Text='<%# Bind("flt_working_hours") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_change_over_timeTextBox" runat="server" 
                          Text='<%# Bind("int_change_over_time") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_completedTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_completed") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_rddTextBox" runat="server" 
                          Text='<%# Bind("dat_rdd") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_currencyTextBox" runat="server" 
                          Text='<%# Bind("txt_currency") %>' />
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
                      <asp:TextBox ID="txt_gl_classTextBox" runat="server" 
                          Text='<%# Bind("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_gradeTextBox" runat="server" 
                          Text='<%# Bind("txt_grade") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_colorTextBox" runat="server" 
                          Text='<%# Bind("txt_color") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_line_assignTextBox" runat="server" 
                          Text='<%# Bind("txt_line_assign") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_spanTextBox" runat="server" 
                          Text='<%# Bind("int_span") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_typeTextBox" runat="server" 
                          Text='<%# Bind("txt_order_type") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_orgn_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_orgn_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_region_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_region_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_order_addedTextBox" runat="server" 
                          Text='<%# Bind("dat_order_added") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_rev_ex_plantTextBox" runat="server" 
                          Text='<%# Bind("dat_rev_ex_plant") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_allocated_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_allocated_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_custTextBox" runat="server" 
                          Text='<%# Bind("txt_ship_cust") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_clean_downTextBox" runat="server" 
                          Text='<%# Bind("txt_clean_down") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_remarkTextBox" runat="server" 
                          Text='<%# Bind("txt_remark") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_line_commentsTextBox" runat="server" 
                          Text='<%# Bind("txt_line_comments") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_from_whseTextBox" runat="server" 
                          Text='<%# Bind("txt_from_whse") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ship_cust_noTextBox" runat="server" 
                          Text='<%# Bind("txt_ship_cust_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_market_segTextBox" runat="server" 
                          Text='<%# Bind("txt_market_seg") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_molderTextBox" runat="server" 
                          Text='<%# Bind("txt_molder") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_end_userTextBox" runat="server" 
                          Text='<%# Bind("txt_end_user") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_sales_priceTextBox" runat="server" 
                          Text='<%# Bind("flt_sales_price") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ess_soTextBox" runat="server" 
                          Text='<%# Bind("txt_ess_so") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ess_sol_noTextBox" runat="server" 
                          Text='<%# Bind("txt_ess_sol_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_allocated_lotsTextBox" runat="server" 
                          Text='<%# Bind("txt_allocated_lots") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_payment_termTextBox" runat="server" 
                          Text='<%# Bind("txt_payment_term") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_package_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_package_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_tbdTextBox" runat="server" 
                          Text='<%# Bind("txt_tbd") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_uploadTextBox" runat="server" 
                          Text='<%# Bind("txt_upload") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_batch_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_batch_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_actual_line_noTextBox" runat="server" 
                          Text='<%# Bind("txt_actual_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_actual_startTextBox" runat="server" 
                          Text='<%# Bind("dat_actual_start") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_actual_finishTextBox" runat="server" 
                          Text='<%# Bind("dat_actual_finish") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_qty_manTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_qty_man") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="int_formula_versionTextBox" runat="server" 
                          Text='<%# Bind("int_formula_version") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_VIPTextBox" runat="server" 
                          Text='<%# Bind("txt_VIP") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_unallocate_qtyTextBox" runat="server" 
                          Text='<%# Bind("flt_unallocate_qty") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_allocation_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_allocation_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lng_VIP_lead_timeTextBox" runat="server" 
                          Text='<%# Bind("lng_VIP_lead_time") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="lng_AdvanceOfRevisionTextBox" runat="server" 
                          Text='<%# Bind("lng_AdvanceOfRevision") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_payment_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_payment_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_order_statusTextBox" runat="server" 
                          Text='<%# Bind("txt_order_status") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_process_technicsTextBox" runat="server" 
                          Text='<%# Bind("txt_process_technics") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="dat_start_from_qaTextBox" runat="server" 
                          Text='<%# Bind("dat_start_from_qa") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_actual_qty_from_qaTextBox" runat="server" 
                          Text='<%# Bind("flt_actual_qty_from_qa") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_FDATextBox" runat="server" 
                          Text='<%# Bind("txt_FDA") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_remark_spclTextBox" runat="server" 
                          Text='<%# Bind("txt_remark_spcl") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_auxiliary_codeTextBox" runat="server" 
                          Text='<%# Bind("txt_auxiliary_code") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_auxiliary_code_for_line_noTextBox" runat="server" 
                          Text='<%# Bind("txt_auxiliary_code_for_line_no") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="flt_marginTextBox" runat="server" 
                          Text='<%# Bind("flt_margin") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_FromUserTextBox" runat="server" 
                          Text='<%# Bind("txt_FromUser") %>' />
                  </td>
                  <td>
                      <asp:TextBox ID="txt_ToUserTextBox" runat="server" 
                          Text='<%# Bind("txt_ToUser") %>' />
                  </td>
              </tr>
          </InsertItemTemplate>
          <ItemTemplate>
              <tr style="background-color:#DCDCDC; color: #000000;">
                  <td>
                      <asp:LinkButton ID="DeleteButton" runat="server" CommandName="Delete" 
                          Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');"  />
                      <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_keyLabel" runat="server" 
                          Text='<%# Eval("txt_order_key") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_status_keyLabel" runat="server" 
                          Text='<%# Eval("int_status_key") %>' />
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
                      <asp:Label ID="txt_item_noLabel" runat="server" 
                          Text='<%# Eval("txt_item_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_order_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_order_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="planned_production_qtyLabel" runat="server" 
                          Text='<%# Eval("planned_production_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_etdLabel" runat="server" 
                          Text='<%# Eval("dat_etd") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_line_noLabel" runat="server" 
                          Text='<%# Eval("int_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_lot_noLabel" runat="server" 
                          Text='<%# Eval("txt_lot_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_start_dateLabel" runat="server" 
                          Text='<%# Eval("dat_start_date") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_finish_dateLabel" runat="server" 
                          Text='<%# Eval("dat_finish_date") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_new_explantLabel" runat="server" 
                          Text='<%# Eval("dat_new_explant") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_working_hoursLabel" runat="server" 
                          Text='<%# Eval("flt_working_hours") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_change_over_timeLabel" runat="server" 
                          Text='<%# Eval("int_change_over_time") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_completedLabel" runat="server" 
                          Text='<%# Eval("flt_actual_completed") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_rddLabel" runat="server" 
                          Text='<%# Eval("dat_rdd") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_currencyLabel" runat="server" 
                          Text='<%# Eval("txt_currency") %>' />
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
                      <asp:Label ID="txt_gl_classLabel" runat="server" 
                          Text='<%# Eval("txt_gl_class") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_gradeLabel" runat="server" 
                          Text='<%# Eval("txt_grade") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_colorLabel" runat="server" 
                          Text='<%# Eval("txt_color") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_line_assignLabel" runat="server" 
                          Text='<%# Eval("txt_line_assign") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_spanLabel" runat="server" 
                          Text='<%# Eval("int_span") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_typeLabel" runat="server" 
                          Text='<%# Eval("txt_order_type") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_orgn_codeLabel" runat="server" 
                          Text='<%# Eval("txt_orgn_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_region_codeLabel" runat="server" 
                          Text='<%# Eval("txt_region_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_order_addedLabel" runat="server" 
                          Text='<%# Eval("dat_order_added") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_rev_ex_plantLabel" runat="server" 
                          Text='<%# Eval("dat_rev_ex_plant") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_allocated_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_allocated_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_custLabel" runat="server" 
                          Text='<%# Eval("txt_ship_cust") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_clean_downLabel" runat="server" 
                          Text='<%# Eval("txt_clean_down") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_remarkLabel" runat="server" 
                          Text='<%# Eval("txt_remark") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_line_commentsLabel" runat="server" 
                          Text='<%# Eval("txt_line_comments") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_from_whseLabel" runat="server" 
                          Text='<%# Eval("txt_from_whse") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ship_cust_noLabel" runat="server" 
                          Text='<%# Eval("txt_ship_cust_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_market_segLabel" runat="server" 
                          Text='<%# Eval("txt_market_seg") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_molderLabel" runat="server" 
                          Text='<%# Eval("txt_molder") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_end_userLabel" runat="server" 
                          Text='<%# Eval("txt_end_user") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_sales_priceLabel" runat="server" 
                          Text='<%# Eval("flt_sales_price") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ess_soLabel" runat="server" 
                          Text='<%# Eval("txt_ess_so") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ess_sol_noLabel" runat="server" 
                          Text='<%# Eval("txt_ess_sol_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_allocated_lotsLabel" runat="server" 
                          Text='<%# Eval("txt_allocated_lots") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_payment_termLabel" runat="server" 
                          Text='<%# Eval("txt_payment_term") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_package_codeLabel" runat="server" 
                          Text='<%# Eval("txt_package_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_tbdLabel" runat="server" 
                          Text='<%# Eval("txt_tbd") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_uploadLabel" runat="server" 
                          Text='<%# Eval("txt_upload") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_batch_statusLabel" runat="server" 
                          Text='<%# Eval("txt_batch_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_actual_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_actual_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_actual_startLabel" runat="server" 
                          Text='<%# Eval("dat_actual_start") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_actual_finishLabel" runat="server" 
                          Text='<%# Eval("dat_actual_finish") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_actual_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_qty_manLabel" runat="server" 
                          Text='<%# Eval("flt_actual_qty_man") %>' />
                  </td>
                  <td>
                      <asp:Label ID="int_formula_versionLabel" runat="server" 
                          Text='<%# Eval("int_formula_version") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_VIPLabel" runat="server" 
                          Text='<%# Eval("txt_VIP") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_unallocate_qtyLabel" runat="server" 
                          Text='<%# Eval("flt_unallocate_qty") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_allocation_statusLabel" runat="server" 
                          Text='<%# Eval("txt_allocation_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lng_VIP_lead_timeLabel" runat="server" 
                          Text='<%# Eval("lng_VIP_lead_time") %>' />
                  </td>
                  <td>
                      <asp:Label ID="lng_AdvanceOfRevisionLabel" runat="server" 
                          Text='<%# Eval("lng_AdvanceOfRevision") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_payment_statusLabel" runat="server" 
                          Text='<%# Eval("txt_payment_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_order_statusLabel" runat="server" 
                          Text='<%# Eval("txt_order_status") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_process_technicsLabel" runat="server" 
                          Text='<%# Eval("txt_process_technics") %>' />
                  </td>
                  <td>
                      <asp:Label ID="dat_start_from_qaLabel" runat="server" 
                          Text='<%# Eval("dat_start_from_qa") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_actual_qty_from_qaLabel" runat="server" 
                          Text='<%# Eval("flt_actual_qty_from_qa") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_FDALabel" runat="server" 
                          Text='<%# Eval("txt_FDA") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_remark_spclLabel" runat="server" 
                          Text='<%# Eval("txt_remark_spcl") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_auxiliary_codeLabel" runat="server" 
                          Text='<%# Eval("txt_auxiliary_code") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_auxiliary_code_for_line_noLabel" runat="server" 
                          Text='<%# Eval("txt_auxiliary_code_for_line_no") %>' />
                  </td>
                  <td>
                      <asp:Label ID="flt_marginLabel" runat="server" 
                          Text='<%# Eval("flt_margin") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_FromUserLabel" runat="server" 
                          Text='<%# Eval("txt_FromUser") %>' />
                  </td>
                  <td>
                      <asp:Label ID="txt_ToUserLabel" runat="server" 
                          Text='<%# Eval("txt_ToUser") %>' />
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
                                  </th>
                                  <th runat="server">
                                      order key</th>
                                  <th runat="server">
                                      status<br />key</th>
                                  <th runat="server">
                                      order no</th>
                                  <th runat="server">
                                      order<br />line<br />no</th>
                                  <th runat="server">
                                      local so</th>
                                  <th runat="server">
                                      item<br />no</th>
                                  <th runat="server">
                                      order<br />qty</th>
                                  <th runat="server">
                                      planned<br />production<br />qty</th>
                                  <th runat="server">
                                      dat_etd</th>
                                  <th runat="server">
                                      line<br />no</th>
                                  <th runat="server">
                                      lot<br />no</th>
                                  <th runat="server">
                                      dat_start_date</th>
                                  <th runat="server">
                                      dat_finish_date</th>
                                  <th runat="server">
                                      dat_new_explant</th>
                                  <th runat="server">
                                      flt_working_hours</th>
                                  <th runat="server">
                                      int_change_over_time</th>
                                  <th runat="server">
                                      flt_actual_completed</th>
                                  <th runat="server">
                                      dat_rdd</th>
                                  <th runat="server">
                                      txt_currency</th>
                                  <th runat="server">
                                      txt_destination</th>
                                  <th runat="server">
                                      txt_ship_method</th>
                                  <th runat="server">
                                      txt_gl_class</th>
                                  <th runat="server">
                                      txt_grade</th>
                                  <th runat="server">
                                      txt_color</th>
                                  <th runat="server">
                                      txt_line_assign</th>
                                  <th runat="server">
                                      int_span</th>
                                  <th runat="server">
                                      txt_order_type</th>
                                  <th runat="server">
                                      txt_orgn_code</th>
                                  <th runat="server">
                                      txt_region_code</th>
                                  <th runat="server">
                                      dat_order_added</th>
                                  <th runat="server">
                                      dat_rev_ex_plant</th>
                                  <th runat="server">
                                      flt_allocated_qty</th>
                                  <th runat="server">
                                      txt_ship_cust</th>
                                  <th runat="server">
                                      txt_clean_down</th>
                                  <th runat="server">
                                      txt_remark</th>
                                  <th runat="server">
                                      txt_line_comments</th>
                                  <th runat="server">
                                      txt_from_whse</th>
                                  <th runat="server">
                                      txt_ship_cust_no</th>
                                  <th runat="server">
                                      txt_market_seg</th>
                                  <th runat="server">
                                      txt_molder</th>
                                  <th runat="server">
                                      txt_end_user</th>
                                  <th runat="server">
                                      flt_sales_price</th>
                                  <th runat="server">
                                      txt_ess_so</th>
                                  <th runat="server">
                                      txt_ess_sol_no</th>
                                  <th runat="server">
                                      txt_allocated_lots</th>
                                  <th runat="server">
                                      txt_payment_term</th>
                                  <th runat="server">
                                      txt_package_code</th>
                                  <th runat="server">
                                      txt_tbd</th>
                                  <th runat="server">
                                      txt_upload</th>
                                  <th runat="server">
                                      txt_batch_status</th>
                                  <th runat="server">
                                      txt_actual_line_no</th>
                                  <th runat="server">
                                      dat_actual_start</th>
                                  <th runat="server">
                                      dat_actual_finish</th>
                                  <th runat="server">
                                      flt_actual_qty</th>
                                  <th runat="server">
                                      flt_actual_qty_man</th>
                                  <th runat="server">
                                      int_formula_version</th>
                                  <th runat="server">
                                      txt_VIP</th>
                                  <th runat="server">
                                      flt_unallocate_qty</th>
                                  <th runat="server">
                                      txt_allocation_status</th>
                                  <th runat="server">
                                      lng_VIP_lead_time</th>
                                  <th runat="server">
                                      lng_AdvanceOfRevision</th>
                                  <th runat="server">
                                      txt_payment_status</th>
                                  <th runat="server">
                                      txt_order_status</th>
                                  <th runat="server">
                                      txt_process_technics</th>
                                  <th runat="server">
                                      dat_start_from_qa</th>
                                  <th runat="server">
                                      flt_actual_qty_from_qa</th>
                                  <th runat="server">
                                      txt_FDA</th>
                                  <th runat="server">
                                      txt_remark_spcl</th>
                                  <th runat="server">
                                      txt_auxiliary_code</th>
                                  <th runat="server">
                                      txt_auxiliary_code_for_line_no</th>
                                  <th runat="server">
                                      flt_margin</th>
                                  <th runat="server">
                                      txt_FromUser</th>
                                  <th runat="server">
                                      txt_ToUser</th>
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
                            Text="Delete" OnClientClick="return confirm('Are you sure you want to delete this item?');"  />
                        <asp:Button ID="EditButton" runat="server" CommandName="Edit" Text="Edit" />
                    </td>
                    <td>
                        <asp:Label ID="txt_order_keyLabel" runat="server" 
                            Text='<%# Eval("txt_order_key") %>' />
                    </td>
                    <td>
                        <asp:Label ID="int_status_keyLabel" runat="server" 
                            Text='<%# Eval("int_status_key") %>' />
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
                        <asp:Label ID="txt_item_noLabel" runat="server" 
                            Text='<%# Eval("txt_item_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_order_qtyLabel" runat="server" 
                            Text='<%# Eval("flt_order_qty") %>' />
                    </td>
                    <td>
                        <asp:Label ID="planned_production_qtyLabel" runat="server" 
                            Text='<%# Eval("planned_production_qty") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_etdLabel" runat="server" Text='<%# Eval("dat_etd") %>' />
                    </td>
                    <td>
                        <asp:Label ID="int_line_noLabel" runat="server" 
                            Text='<%# Eval("int_line_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_lot_noLabel" runat="server" 
                            Text='<%# Eval("txt_lot_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_start_dateLabel" runat="server" 
                            Text='<%# Eval("dat_start_date") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_finish_dateLabel" runat="server" 
                            Text='<%# Eval("dat_finish_date") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_new_explantLabel" runat="server" 
                            Text='<%# Eval("dat_new_explant") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_working_hoursLabel" runat="server" 
                            Text='<%# Eval("flt_working_hours") %>' />
                    </td>
                    <td>
                        <asp:Label ID="int_change_over_timeLabel" runat="server" 
                            Text='<%# Eval("int_change_over_time") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_actual_completedLabel" runat="server" 
                            Text='<%# Eval("flt_actual_completed") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_rddLabel" runat="server" 
                            Text='<%# Eval("dat_rdd") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_currencyLabel" runat="server" 
                            Text='<%# Eval("txt_currency") %>' />
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
                        <asp:Label ID="txt_gl_classLabel" runat="server" 
                            Text='<%# Eval("txt_gl_class") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_gradeLabel" runat="server" 
                            Text='<%# Eval("txt_grade") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_colorLabel" runat="server" 
                            Text='<%# Eval("txt_color") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_line_assignLabel" runat="server" 
                            Text='<%# Eval("txt_line_assign") %>' />
                    </td>
                    <td>
                        <asp:Label ID="int_spanLabel" runat="server" 
                            Text='<%# Eval("int_span") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_order_typeLabel" runat="server" 
                            Text='<%# Eval("txt_order_type") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_orgn_codeLabel" runat="server" 
                            Text='<%# Eval("txt_orgn_code") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_region_codeLabel" runat="server" 
                            Text='<%# Eval("txt_region_code") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_order_addedLabel" runat="server" 
                            Text='<%# Eval("dat_order_added") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_rev_ex_plantLabel" runat="server" 
                            Text='<%# Eval("dat_rev_ex_plant") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_allocated_qtyLabel" runat="server" 
                            Text='<%# Eval("flt_allocated_qty") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ship_custLabel" runat="server" 
                            Text='<%# Eval("txt_ship_cust") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_clean_downLabel" runat="server" 
                            Text='<%# Eval("txt_clean_down") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_remarkLabel" runat="server" 
                            Text='<%# Eval("txt_remark") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_line_commentsLabel" runat="server" 
                            Text='<%# Eval("txt_line_comments") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_from_whseLabel" runat="server" 
                            Text='<%# Eval("txt_from_whse") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ship_cust_noLabel" runat="server" 
                            Text='<%# Eval("txt_ship_cust_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_market_segLabel" runat="server" 
                            Text='<%# Eval("txt_market_seg") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_molderLabel" runat="server" 
                            Text='<%# Eval("txt_molder") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_end_userLabel" runat="server" 
                            Text='<%# Eval("txt_end_user") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_sales_priceLabel" runat="server" 
                            Text='<%# Eval("flt_sales_price") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ess_soLabel" runat="server" 
                            Text='<%# Eval("txt_ess_so") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ess_sol_noLabel" runat="server" 
                            Text='<%# Eval("txt_ess_sol_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_allocated_lotsLabel" runat="server" 
                            Text='<%# Eval("txt_allocated_lots") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_payment_termLabel" runat="server" 
                            Text='<%# Eval("txt_payment_term") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_package_codeLabel" runat="server" 
                            Text='<%# Eval("txt_package_code") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_tbdLabel" runat="server" 
                            Text='<%# Eval("txt_tbd") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_uploadLabel" runat="server" 
                            Text='<%# Eval("txt_upload") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_batch_statusLabel" runat="server" 
                            Text='<%# Eval("txt_batch_status") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_actual_line_noLabel" runat="server" 
                            Text='<%# Eval("txt_actual_line_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_actual_startLabel" runat="server" 
                            Text='<%# Eval("dat_actual_start") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_actual_finishLabel" runat="server" 
                            Text='<%# Eval("dat_actual_finish") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_actual_qtyLabel" runat="server" 
                            Text='<%# Eval("flt_actual_qty") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_actual_qty_manLabel" runat="server" 
                            Text='<%# Eval("flt_actual_qty_man") %>' />
                    </td>
                    <td>
                        <asp:Label ID="int_formula_versionLabel" runat="server" 
                            Text='<%# Eval("int_formula_version") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_VIPLabel" runat="server" 
                            Text='<%# Eval("txt_VIP") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_unallocate_qtyLabel" runat="server" 
                            Text='<%# Eval("flt_unallocate_qty") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_allocation_statusLabel" runat="server" 
                            Text='<%# Eval("txt_allocation_status") %>' />
                    </td>
                    <td>
                        <asp:Label ID="lng_VIP_lead_timeLabel" runat="server" 
                            Text='<%# Eval("lng_VIP_lead_time") %>' />
                    </td>
                    <td>
                        <asp:Label ID="lng_AdvanceOfRevisionLabel" runat="server" 
                            Text='<%# Eval("lng_AdvanceOfRevision") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_payment_statusLabel" runat="server" 
                            Text='<%# Eval("txt_payment_status") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_order_statusLabel" runat="server" 
                            Text='<%# Eval("txt_order_status") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_process_technicsLabel" runat="server" 
                            Text='<%# Eval("txt_process_technics") %>' />
                    </td>
                    <td>
                        <asp:Label ID="dat_start_from_qaLabel" runat="server" 
                            Text='<%# Eval("dat_start_from_qa") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_actual_qty_from_qaLabel" runat="server" 
                            Text='<%# Eval("flt_actual_qty_from_qa") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_FDALabel" runat="server" 
                            Text='<%# Eval("txt_FDA") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_remark_spclLabel" runat="server" 
                            Text='<%# Eval("txt_remark_spcl") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_auxiliary_codeLabel" runat="server" 
                            Text='<%# Eval("txt_auxiliary_code") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_auxiliary_code_for_line_noLabel" runat="server" 
                            Text='<%# Eval("txt_auxiliary_code_for_line_no") %>' />
                    </td>
                    <td>
                        <asp:Label ID="flt_marginLabel" runat="server" 
                            Text='<%# Eval("flt_margin") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_FromUserLabel" runat="server" 
                            Text='<%# Eval("txt_FromUser") %>' />
                    </td>
                    <td>
                        <asp:Label ID="txt_ToUserLabel" runat="server" 
                            Text='<%# Eval("txt_ToUser") %>' />
                    </td>
                </tr>
          </SelectedItemTemplate>
            </asp:ListView>
    <asp:SqlDataSource ID="SDS1" runat="server" 
        ConnectionString="Provider=Microsoft.ACE.OleDb.12.0;Data Source=C:\inetpub\wwwroot\SHA\App_Data\db_Resin.accdb" 
        ProviderName="System.Data.SqlClient"
        OldValuesParameterFormatString="original_{0}"
        
    SelectCommand="SELECT * FROM [Esch_CQ_tbl_orders] ORDER BY [int_line_no], [dat_start_date]" 
    DeleteCommand="DELETE FROM [Esch_CQ_tbl_orders] WHERE [txt_order_key] = @original_txt_order_key " 
                                  
            
            
            InsertCommand="INSERT INTO [Esch_CQ_tbl_orders] ([int_status_key], [txt_local_so], [txt_item_no], [flt_order_qty], [flt_unallocate_qty], [planned_production_qty], [dat_etd], [int_line_no], [txt_lot_no], [dat_start_date], [dat_finish_date], [dat_new_explant], [int_span], [dat_rdd], [dat_order_added], [txt_currency], [txt_remark], [txt_VIP], [txt_order_key], [txt_process_technics], [txt_order_status], [flt_actual_completed], [flt_actual_qty], [flt_actual_qty_man], [dat_start_from_qa], [flt_actual_qty_from_qa], [txt_destination], [txt_ship_method], [txt_molder], [txt_end_user], [txt_payment_term], [lng_VIP_lead_time], [lng_AdvanceOfRevision], [txt_allocation_status], [txt_FDA], [txt_payment_status], [txt_remark_spcl], [txt_auxiliary_code], [txt_auxiliary_code_for_line_no], [txt_order_no], [txt_order_line_no], [flt_working_hours], [int_change_over_time], [txt_gl_class], [txt_grade], [txt_color], [txt_line_assign], [txt_order_type], [txt_orgn_code], [txt_region_code], [dat_rev_ex_plant], [flt_allocated_qty], [txt_ship_cust], [txt_clean_down], [txt_line_comments], [txt_from_whse], [txt_ship_cust_no], [txt_market_seg], [flt_sales_price], [flt_margin], [txt_ess_so], [txt_ess_sol_no], [txt_allocated_lots], [txt_package_code], [txt_tbd], [txt_upload], [txt_batch_status], [txt_actual_line_no], [dat_actual_start], [dat_actual_finish], [int_formula_version], [txt_FromUser], [txt_ToUser]) VALUES (@int_status_key, @txt_local_so, @txt_item_no, @flt_order_qty, @flt_unallocate_qty, @planned_production_qty, @dat_etd, @int_line_no, @txt_lot_no, @dat_start_date, @dat_finish_date, @dat_new_explant, @int_span, @dat_rdd, @dat_order_added, @txt_currency, @txt_remark, @txt_VIP, @txt_order_key, @txt_process_technics, @txt_order_status, @flt_actual_completed, @flt_actual_qty, @flt_actual_qty_man, @dat_start_from_qa, @flt_actual_qty_from_qa, @txt_destination, @txt_ship_method, @txt_molder, @txt_end_user, @txt_payment_term, @lng_VIP_lead_time, @lng_AdvanceOfRevision, @txt_allocation_status, @txt_FDA, @txt_payment_status, @txt_remark_spcl, @txt_auxiliary_code, @txt_auxiliary_code_for_line_no, @txt_order_no, @txt_order_line_no, @flt_working_hours, @int_change_over_time, @txt_gl_class, @txt_grade, @txt_color, @txt_line_assign, @txt_order_type, @txt_orgn_code, @txt_region_code, @dat_rev_ex_plant, @flt_allocated_qty, @txt_ship_cust, @txt_clean_down, @txt_line_comments, @txt_from_whse, @txt_ship_cust_no, @txt_market_seg, @flt_sales_price, @flt_margin, @txt_ess_so, @txt_ess_sol_no, @txt_allocated_lots, @txt_package_code, @txt_tbd, @txt_upload, @txt_batch_status, @txt_actual_line_no, @dat_actual_start, @dat_actual_finish, @int_formula_version, @txt_FromUser, @txt_ToUser)" 
        UpdateCommand="UPDATE [Esch_CQ_tbl_orders] SET [int_status_key] = @int_status_key, [txt_local_so] = @txt_local_so, [txt_item_no] = @txt_item_no, [flt_order_qty] = @flt_order_qty, [flt_unallocate_qty] = @flt_unallocate_qty, [planned_production_qty] = @planned_production_qty, [dat_etd] = @dat_etd, [int_line_no] = @int_line_no, [txt_lot_no] = @txt_lot_no, [dat_start_date] = @dat_start_date, [dat_finish_date] = @dat_finish_date, [dat_new_explant] = @dat_new_explant, [int_span] = @int_span, [dat_rdd] = @dat_rdd, [dat_order_added] = @dat_order_added, [txt_currency] = @txt_currency, [txt_remark] = @txt_remark, [txt_VIP] = @txt_VIP, [txt_process_technics] = @txt_process_technics, [txt_order_status] = @txt_order_status, [flt_actual_completed] = @flt_actual_completed, [flt_actual_qty] = @flt_actual_qty, [flt_actual_qty_man] = @flt_actual_qty_man, [dat_start_from_qa] = @dat_start_from_qa, [flt_actual_qty_from_qa] = @flt_actual_qty_from_qa, [txt_destination] = @txt_destination, [txt_ship_method] = @txt_ship_method, [txt_molder] = @txt_molder, [txt_end_user] = @txt_end_user, [txt_payment_term] = @txt_payment_term, [lng_VIP_lead_time] = @lng_VIP_lead_time, [lng_AdvanceOfRevision] = @lng_AdvanceOfRevision, [txt_allocation_status] = @txt_allocation_status, [txt_FDA] = @txt_FDA, [txt_payment_status] = @txt_payment_status, [txt_remark_spcl] = @txt_remark_spcl, [txt_auxiliary_code] = @txt_auxiliary_code, [txt_auxiliary_code_for_line_no] = @txt_auxiliary_code_for_line_no, [txt_order_no] = @txt_order_no, [txt_order_line_no] = @txt_order_line_no, [flt_working_hours] = @flt_working_hours, [int_change_over_time] = @int_change_over_time, [txt_gl_class] = @txt_gl_class, [txt_grade] = @txt_grade, [txt_color] = @txt_color, [txt_line_assign] = @txt_line_assign, [txt_order_type] = @txt_order_type, [txt_orgn_code] = @txt_orgn_code, [txt_region_code] = @txt_region_code, [dat_rev_ex_plant] = @dat_rev_ex_plant, [flt_allocated_qty] = @flt_allocated_qty, [txt_ship_cust] = @txt_ship_cust, [txt_clean_down] = @txt_clean_down, [txt_line_comments] = @txt_line_comments, [txt_from_whse] = @txt_from_whse, [txt_ship_cust_no] = @txt_ship_cust_no, [txt_market_seg] = @txt_market_seg, [flt_sales_price] = @flt_sales_price, [flt_margin] = @flt_margin, [txt_ess_so] = @txt_ess_so, [txt_ess_sol_no] = @txt_ess_sol_no, [txt_allocated_lots] = @txt_allocated_lots, [txt_package_code] = @txt_package_code, [txt_tbd] = @txt_tbd, [txt_upload] = @txt_upload, [txt_batch_status] = @txt_batch_status, [txt_actual_line_no] = @txt_actual_line_no, [dat_actual_start] = @dat_actual_start, [dat_actual_finish] = @dat_actual_finish, [int_formula_version] = @int_formula_version, [txt_FromUser] = @txt_FromUser, [txt_ToUser] = @txt_ToUser WHERE [txt_order_key] = @original_txt_order_key " >
  <DeleteParameters><asp:Parameter Name="original_txt_order_key" Type="String" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="int_status_key" Type="String" />
            <asp:Parameter Name="txt_local_so" Type="String" />
            <asp:Parameter Name="txt_item_no" Type="String" />
            <asp:Parameter Name="flt_order_qty" Type="Int32" />
            <asp:Parameter Name="flt_unallocate_qty" Type="Int32" />
            <asp:Parameter Name="planned_production_qty" Type="Int32" />
            <asp:Parameter Name="dat_etd" Type="DateTime" />
            <asp:Parameter Name="int_line_no" Type="Int32" />
            <asp:Parameter Name="txt_lot_no" Type="String" />
            <asp:Parameter Name="dat_start_date" Type="DateTime" />
            <asp:Parameter Name="dat_finish_date" Type="DateTime" />
            <asp:Parameter Name="dat_new_explant" Type="DateTime" />
            <asp:Parameter Name="int_span" Type="Int32" />
            <asp:Parameter Name="dat_rdd" Type="DateTime" />
            <asp:Parameter Name="dat_order_added" Type="DateTime" />
            <asp:Parameter Name="txt_currency" Type="String" />
            <asp:Parameter Name="txt_remark" Type="String" />
            <asp:Parameter Name="txt_VIP" Type="String" />
            <asp:Parameter Name="txt_order_key" Type="String" />
            <asp:Parameter Name="txt_process_technics" Type="String" />
            <asp:Parameter Name="txt_order_status" Type="String" />
            <asp:Parameter Name="flt_actual_completed" Type="Single" />
            <asp:Parameter Name="flt_actual_qty" Type="Int32" />
            <asp:Parameter Name="flt_actual_qty_man" Type="Int32" />
            <asp:Parameter Name="dat_start_from_qa" Type="DateTime" />
            <asp:Parameter Name="flt_actual_qty_from_qa" Type="Int32" />
            <asp:Parameter Name="txt_destination" Type="String" />
            <asp:Parameter Name="txt_ship_method" Type="String" />
            <asp:Parameter Name="txt_molder" Type="String" />
            <asp:Parameter Name="txt_end_user" Type="String" />
            <asp:Parameter Name="txt_payment_term" Type="String" />
            <asp:Parameter Name="lng_VIP_lead_time" Type="Int32" />
            <asp:Parameter Name="lng_AdvanceOfRevision" Type="Int32" />
            <asp:Parameter Name="txt_allocation_status" Type="String" />
            <asp:Parameter Name="txt_FDA" Type="String" />
            <asp:Parameter Name="txt_payment_status" Type="String" />
            <asp:Parameter Name="txt_remark_spcl" Type="String" />
            <asp:Parameter Name="txt_auxiliary_code" Type="String" />
            <asp:Parameter Name="txt_auxiliary_code_for_line_no" Type="String" />
            <asp:Parameter Name="txt_order_no" Type="String" />
            <asp:Parameter Name="txt_order_line_no" Type="String" />
            <asp:Parameter Name="flt_working_hours" Type="Int32" />
            <asp:Parameter Name="int_change_over_time" Type="Int32" />
            <asp:Parameter Name="txt_gl_class" Type="String" />
            <asp:Parameter Name="txt_grade" Type="String" />
            <asp:Parameter Name="txt_color" Type="String" />
            <asp:Parameter Name="txt_line_assign" Type="String" />
            <asp:Parameter Name="txt_order_type" Type="String" />
            <asp:Parameter Name="txt_orgn_code" Type="String" />
            <asp:Parameter Name="txt_region_code" Type="String" />
            <asp:Parameter Name="dat_rev_ex_plant" Type="DateTime" />
            <asp:Parameter Name="flt_allocated_qty" Type="Int32" />
            <asp:Parameter Name="txt_ship_cust" Type="String" />
            <asp:Parameter Name="txt_clean_down" Type="String" />
            <asp:Parameter Name="txt_line_comments" Type="String" />
            <asp:Parameter Name="txt_from_whse" Type="String" />
            <asp:Parameter Name="txt_ship_cust_no" Type="String" />
            <asp:Parameter Name="txt_market_seg" Type="String" />
            <asp:Parameter Name="flt_sales_price" Type="Single" />
            <asp:Parameter Name="flt_margin" Type="Single" />
            <asp:Parameter Name="txt_ess_so" Type="String" />
            <asp:Parameter Name="txt_ess_sol_no" Type="String" />
            <asp:Parameter Name="txt_allocated_lots" Type="String" />
            <asp:Parameter Name="txt_package_code" Type="String" />
            <asp:Parameter Name="txt_tbd" Type="String" />
            <asp:Parameter Name="txt_upload" Type="String" />
            <asp:Parameter Name="txt_batch_status" Type="String" />
            <asp:Parameter Name="txt_actual_line_no" Type="String" />
            <asp:Parameter Name="dat_actual_start" Type="DateTime" />
            <asp:Parameter Name="dat_actual_finish" Type="DateTime" />
            <asp:Parameter Name="int_formula_version" Type="Int32" />
            <asp:Parameter Name="txt_FromUser" Type="String" />
            <asp:Parameter Name="txt_ToUser" Type="String" />
        </InsertParameters>
        <UpdateParameters>
            <asp:Parameter Name="int_status_key" Type="String" />
            <asp:Parameter Name="txt_local_so" Type="String" />
            <asp:Parameter Name="txt_item_no" Type="String" />
            <asp:Parameter Name="flt_order_qty" Type="Int32" />
            <asp:Parameter Name="flt_unallocate_qty" Type="Int32" />
            <asp:Parameter Name="planned_production_qty" Type="Int32" />
            <asp:Parameter Name="dat_etd" Type="DateTime" />
            <asp:Parameter Name="int_line_no" Type="Int32" />
            <asp:Parameter Name="txt_lot_no" Type="String" />
            <asp:Parameter Name="dat_start_date" Type="DateTime" />
            <asp:Parameter Name="dat_finish_date" Type="DateTime" />
            <asp:Parameter Name="dat_new_explant" Type="DateTime" />
            <asp:Parameter Name="int_span" Type="Int32" />
            <asp:Parameter Name="dat_rdd" Type="DateTime" />
            <asp:Parameter Name="dat_order_added" Type="DateTime" />
            <asp:Parameter Name="txt_currency" Type="String" />
            <asp:Parameter Name="txt_remark" Type="String" />
            <asp:Parameter Name="txt_VIP" Type="String" />
            <asp:Parameter Name="txt_process_technics" Type="String" />
            <asp:Parameter Name="txt_order_status" Type="String" />
            <asp:Parameter Name="flt_actual_completed" Type="Single" />
            <asp:Parameter Name="flt_actual_qty" Type="Int32" />
            <asp:Parameter Name="flt_actual_qty_man" Type="Int32" />
            <asp:Parameter Name="dat_start_from_qa" Type="DateTime" />
            <asp:Parameter Name="flt_actual_qty_from_qa" Type="Int32" />
            <asp:Parameter Name="txt_destination" Type="String" />
            <asp:Parameter Name="txt_ship_method" Type="String" />
            <asp:Parameter Name="txt_molder" Type="String" />
            <asp:Parameter Name="txt_end_user" Type="String" />
            <asp:Parameter Name="txt_payment_term" Type="String" />
            <asp:Parameter Name="lng_VIP_lead_time" Type="Int32" />
            <asp:Parameter Name="lng_AdvanceOfRevision" Type="Int32" />
            <asp:Parameter Name="txt_allocation_status" Type="String" />
            <asp:Parameter Name="txt_FDA" Type="String" />
            <asp:Parameter Name="txt_payment_status" Type="String" />
            <asp:Parameter Name="txt_remark_spcl" Type="String" />
            <asp:Parameter Name="txt_auxiliary_code" Type="String" />
            <asp:Parameter Name="txt_auxiliary_code_for_line_no" Type="String" />
            <asp:Parameter Name="txt_order_no" Type="String" />
            <asp:Parameter Name="txt_order_line_no" Type="String" />
            <asp:Parameter Name="flt_working_hours" Type="Int32" />
            <asp:Parameter Name="int_change_over_time" Type="Int32" />
            <asp:Parameter Name="txt_gl_class" Type="String" />
            <asp:Parameter Name="txt_grade" Type="String" />
            <asp:Parameter Name="txt_color" Type="String" />
            <asp:Parameter Name="txt_line_assign" Type="String" />
            <asp:Parameter Name="txt_order_type" Type="String" />
            <asp:Parameter Name="txt_orgn_code" Type="String" />
            <asp:Parameter Name="txt_region_code" Type="String" />
            <asp:Parameter Name="dat_rev_ex_plant" Type="DateTime" />
            <asp:Parameter Name="flt_allocated_qty" Type="Int32" />
            <asp:Parameter Name="txt_ship_cust" Type="String" />
            <asp:Parameter Name="txt_clean_down" Type="String" />
            <asp:Parameter Name="txt_line_comments" Type="String" />
            <asp:Parameter Name="txt_from_whse" Type="String" />
            <asp:Parameter Name="txt_ship_cust_no" Type="String" />
            <asp:Parameter Name="txt_market_seg" Type="String" />
            <asp:Parameter Name="flt_sales_price" Type="Single" />
            <asp:Parameter Name="flt_margin" Type="Single" />
            <asp:Parameter Name="txt_ess_so" Type="String" />
            <asp:Parameter Name="txt_ess_sol_no" Type="String" />
            <asp:Parameter Name="txt_allocated_lots" Type="String" />
            <asp:Parameter Name="txt_package_code" Type="String" />
            <asp:Parameter Name="txt_tbd" Type="String" />
            <asp:Parameter Name="txt_upload" Type="String" />
            <asp:Parameter Name="txt_batch_status" Type="String" />
            <asp:Parameter Name="txt_actual_line_no" Type="String" />
            <asp:Parameter Name="dat_actual_start" Type="DateTime" />
            <asp:Parameter Name="dat_actual_finish" Type="DateTime" />
            <asp:Parameter Name="int_formula_version" Type="Int32" />
            <asp:Parameter Name="txt_FromUser" Type="String" />
            <asp:Parameter Name="txt_ToUser" Type="String" />
            <asp:Parameter Name="original_txt_order_key" Type="String" />
            </UpdateParameters>
  
    </asp:SqlDataSource>
    </ContentTemplate><Triggers>
    <asp:AsyncPostBackTrigger ControlID="Filter1" EventName="Click" />
    <asp:AsyncPostBackTrigger ControlID="clrfltr1" EventName="Click" />
    </Triggers></asp:UpdatePanel><br />
    Last upload time:
    <asp:Label ID="lblUpdateTime" runat="server" ClientIDMode="Static"></asp:Label>
    <br />
    <br />
    <br /> 
    <asp:linkButton runat="server" 
        id="delCS" text="Delete current selection" 
    ClientIDMode="Static" ViewStateMode="Disabled" EnableViewState="False" OnClientClick="return confirm('Are you sure you want to delete all the selected orders ?');"  />
    
</asp:Content>

