Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class dragDrop_normalOP_MTI
    Inherits FrequentPlanActions



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString
        
        maxRowNumber = 5001

        pageLoadInitiate(SDS1, DDL1, DDL2, filtercdtn1, Filter1, Download1, hiddenBT)

    End Sub


    Protected Sub Download1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Download1.Click

        downloadExcelFileFromSqlDataSource(SDS1, StatusLabel, hiddenBT)


    End Sub


    Protected Sub Filter1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Filter1.Click
        filterClickHnadler(filtercdtn1, DDL1, DDL2, SDS1, LV1)


    End Sub



    Protected Sub DDL1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL1.TextChanged
        DDLchangeSelection(DDL1, DDL2)
    End Sub

    Protected Sub clrfltr1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles clrfltr1.Click
        clrfltrClickHnadler(SDS1, LV1)
    End Sub

    Protected Sub UpldDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldDel.Click

        UpldDel_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)

    End Sub


    Protected Sub UpldInsrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldInsrt.Click

        UpldInsrt_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)


    End Sub


    '''update table per the data in excel file
    Protected Sub UpldUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldUpdate.Click

        UpldUpdate_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)

    End Sub

    Protected Sub overwrite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles overwrite.Click

        overwrite_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)

    End Sub


    Public Overrides Function dataValidityCheck(ByRef dtTbl As DataTable) As String


        Dim msstring As String = String.Empty

        Try
            For Each rowexcel As DataRow In dtTbl.Rows
                If DBNull.Value.Equals(rowexcel("sequence")) Then
                    msstring = "Empty value in column sequence is not allowed."
                    Exit For
                Else
                    If Not IsNumeric(rowexcel("sequence")) Then
                        msstring = "You need input number in field sequence."
                        Exit For
                    End If
                End If


                If DBNull.Value.Equals(rowexcel("item")) Then
                    msstring = "Empty value in column item is not allowed."
                    Exit For
                Else
                    If rowexcel("item").ToString.IndexOf("-") < 1 Then
                        msstring = "You need input correct item code in field item."
                        Exit For
                    End If
                End If

                If DBNull.Value.Equals(rowexcel("quantity")) Then
                    msstring = "Empty value in column quantity is not allowed."
                    Exit For
                Else
                    If Not IsNumeric(rowexcel("quantity")) Then
                        msstring = "You need input number in field quantity."
                        Exit For
                    End If
                End If

                If DBNull.Value.Equals(rowexcel("line")) Then
                    msstring = "Empty value in column line is not allowed."
                    Exit For
                Else
                    If Not IsNumeric(rowexcel("line")) Then
                        msstring = "You need input number in field line."
                        Exit For
                    End If
                End If


                If DBNull.Value.Equals(rowexcel("startDate")) Then
                    msstring = "Empty value in column startDate is not allowed."
                    Exit For
                Else
                    If Not IsDate(rowexcel("startDate")) Then
                        msstring = "You need input date in field startDate."
                        Exit For
                    End If
                End If


                If String.IsNullOrEmpty(rowexcel("txt_order_no")) OrElse String.IsNullOrEmpty(rowexcel("txt_order_line_no")) Then
                    msstring = "Empty value in field txt_order_no or txt_order_line_no is not allowed."
                    Exit For
                End If


            Next

        Catch
            msstring = "There is something wrong with the data validity."
        End Try

        Return msstring

    End Function






    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    ''' 
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating

        Dim messageText As StringBuilder = New StringBuilder()


        If String.IsNullOrEmpty(e.NewValues("quantity")) Then
            messageText.Append("Empty value in field quantity is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("quantity")) Then
                messageText.Append("You need input number in field quantity.")
            End If
        End If


        If String.IsNullOrEmpty(e.NewValues("item")) Then
            messageText.Append("Empty value in field item is not allowed.")
        Else
            If e.NewValues("item").ToString.IndexOf("-") < 1 Then
                messageText.Append("You need input correct item code in field item.")
            End If
        End If



        If String.IsNullOrEmpty(e.NewValues("line")) Then
            messageText.Append("Empty value in field line is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("line")) Then
                messageText.Append("You need input number in field line.")
            End If
        End If



        If String.IsNullOrEmpty(e.NewValues("startDate")) Then
            messageText.Append("Empty value in field startDate is not allowed.")
        Else
            If Not IsDate(e.NewValues("startDate")) Then
                messageText.Append("You need input date in field startDate.")
            End If
        End If


        If String.IsNullOrEmpty(e.NewValues("txt_order_no")) OrElse String.IsNullOrEmpty(e.NewValues("txt_order_line_no")) Then
            messageText.Append("Empty value in field txt_order_no or txt_order_line_no is not allowed.")
        End If


        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If


    End Sub

    ''' <summary>
    ''' reveal insertion template
    ''' </summary>
    Protected Sub LV1_ItemCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewCommandEventArgs) Handles LV1.ItemCommand
        If e.CommandName.Equals("New", StringComparison.OrdinalIgnoreCase) Then
            Dim me1 As ListView = CType(sender, ListView)
            me1.InsertItemPosition = InsertItemPosition.FirstItem
        End If
    End Sub


    Protected Sub LV1_ItemCanceling(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewCancelEventArgs) Handles LV1.ItemCanceling

        If e.CancelMode = ListViewCancelMode.CancelingInsert Then
            Dim me1 As ListView = CType(sender, ListView)
            me1.InsertItemPosition = InsertItemPosition.None
        End If

    End Sub


    Protected Sub LV1_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewInsertedEventArgs) Handles LV1.ItemInserted
        Dim me1 As ListView = CType(sender, ListView)
        me1.InsertItemPosition = InsertItemPosition.None
    End Sub


    ''' <summary>
    ''' add MTI orders to table Esch_Sh_tbl_orders
    ''' </summary>
    Protected Sub bckToDB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bckToDB.Click

        Dim userName As String = lockKeyTable(priority.addMTI)
        If String.IsNullOrEmpty(userName) Then
            Dim conn As OleDbConnection = New OleDbConnection(ConfigurationManager.ConnectionStrings(dbConnectionName).ProviderName & ConfigurationManager.ConnectionStrings(dbConnectionName).ConnectionString)
            Dim dtUpdateTo0 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Sh_tbl_orders WHERE txt_order_key = ''", conn)
            Dim cmdbAccessCmdBuilder As New OleDbCommandBuilder(dtUpdateTo0)
            dtUpdateTo0.InsertCommand = cmdbAccessCmdBuilder.GetInsertCommand()
            Dim dtTbl0 As DataTable = New DataTable()
            dtUpdateTo0.Fill(dtTbl0)


            Dim dtUpdateTo1 As OleDbDataAdapter = New OleDbDataAdapter("SELECT * FROM Esch_Sh_tbl_MTI_add WHERE decideToAdd = 1 ", conn)
            Dim cmdbAccessCmdBuilder1 As New OleDbCommandBuilder(dtUpdateTo1)
            dtUpdateTo1.UpdateCommand = cmdbAccessCmdBuilder1.GetUpdateCommand()
            Dim dtTbl1 As DataTable = New DataTable()
            dtUpdateTo1.Fill(dtTbl1)

            msgPopUP(dtTbl1.Rows.Count & " MTI orders  added.", Message, False, False)


            Dim newR As DataRow
            For Each r1 As DataRow In dtTbl1.Rows
                newR = dtTbl0.NewRow()
                newR.Item("txt_order_no") = r1.Item("txt_order_no")
                newR.Item("txt_order_line_no") = r1.Item("txt_order_line_no")
                newR.Item("txt_order_key") = newR.Item("txt_order_no") & "-" & newR.Item("txt_order_line_no")
                newR.Item("int_status_key") = newOrderStatus
                newR.Item("txt_item_no") = r1.Item("item")
                newR.Item("flt_order_qty") = r1.Item("quantity")
                newR.Item("flt_unallocate_qty") = r1.Item("quantity")
                newR.Item("planned_production_qty") = r1.Item("quantity")
                newR.Item("dat_etd") = DateTime.Today
                newR.Item("dat_order_added") = DateTime.Today
                newR.Item("dat_rdd") = DateTime.Today.AddDays(2).Date
                newR.Item("txt_orgn_code") = valueOf("strOrganization")
                newR.Item("txt_order_type") = "MTI"
                newR.Item("txt_local_so") = r1.Item("txt_local_so")
                newR.Item("int_line_no") = r1.Item("line")
                newR.Item("txt_gl_class") = r1.Item("txt_gl_class")
                newR.Item("dat_start_date") = r1.Item("startDate")
                newR.Item("txt_remark") = r1.Item("remark")
                newR.Item("txt_gl_class") = r1.Item("txt_gl_class")
                newR.Item("txt_grade") = r1.Item("item").ToString.Split(New Char() {"-"}, StringSplitOptions.RemoveEmptyEntries)(0)
                newR.Item("txt_color") = r1.Item("item").ToString.Split(New Char() {"-"}, StringSplitOptions.RemoveEmptyEntries)(1)

                dtTbl0.Rows.Add(newR)


                r1.Item("decideToAdd") = False 'use to mark that this item has been added to table Esch_Sh_tbl_orders

            Next

            Dim hasException As Boolean = False
            Dim newRows() As DataRow = dtTbl0.Select(Nothing, Nothing, System.Data.DataViewRowState.Added)
            'calculateRSD1(conn, newRows)
            finishTime_exPlantDate_Span1(conn, newRows, hasException)
            'DoAssignScrewDieAndFDA(" And (txt_order_type = 'MTI')")
            AssignScrewDieAndFDA(conn, newRows)

            If hasException Then
                CacheInsert("Oexception", 1)
            End If

            Try
                dtUpdateTo0.Update(dtTbl0)

                dtUpdateTo1.Update(dtTbl1)
            Catch ex As Exception
                msgPopUP("Might duplicate txt_order_key <br />" & ex.Message, Message, True, False)
            End Try

            cmdbAccessCmdBuilder1.Dispose()

            dtTbl0.Dispose()
            dtUpdateTo0.Dispose()

            dtTbl1.Dispose()
            dtUpdateTo1.Dispose()

            conn.Dispose()

            LV1.DataBind()

            unlockKeyTable(priority.addMTI)

        Else
            Message.Text = "<div style='color:red;font-size:150%;'>" & userName & " is using the order detail table" & "</div>"

        End If
    End Sub

    Protected Sub btOM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btOM.Click
        Response.Redirect("~/dragDrop/OrderDetail.aspx")
    End Sub

    
End Class

