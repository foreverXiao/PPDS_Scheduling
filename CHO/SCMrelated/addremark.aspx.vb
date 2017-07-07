Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class SCMrelated_addremark
    Inherits InteracWithExcel



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString

        maxRowNumber = 25001


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
        Dim columNameExist As String = getColumnNameOf_Esch_CQ_tbl_orders().ToLower()

        Try
            Dim i As Integer = 0
            Dim j As Integer = dtTbl.Columns.Count

            For Each rowexcel As DataRow In dtTbl.Rows

                i = 0
                While i < j
                    If DBNull.Value.Equals(rowexcel.Item(i)) Then
                        msstring = "Empty value in column '" & dtTbl.Columns(i).ColumnName & "' is not allowed."
                        Exit For
                    End If
                    i += 1
                End While

                'column name has to be found in table Esch_CQ_tbl_orders
                If columNameExist.IndexOf("," & rowexcel("headerName").ToString.ToLower & ",") < 0 Then
                    msstring = "Column name as '" & rowexcel("headerName") & "' was NOT found in [ detail order ] file! You input a wrong column name."
                    Exit For
                End If

            Next
        Catch
            msstring = "There is something wrong with the data validity."
        End Try

        Return msstring

    End Function

    Protected Sub LV1_ItemInserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewInsertEventArgs) Handles LV1.ItemInserting
        ' Cancel the update operation if any of the fields is empty
        ' or null.
        Dim messageText As StringBuilder = New StringBuilder()
        Dim columNameExist As String = getColumnNameOf_Esch_CQ_tbl_orders().ToLower()

        For Each de As DictionaryEntry In e.Values
            ' Check if the value is null or empty
            If de.Value Is Nothing OrElse de.Value.ToString().Trim().Length = 0 Then
                messageText.Append("Cannot insert an empty value.<br />")
            End If
        Next

        'column name has to be found in table Esch_CQ_tbl_orders
        If columNameExist.IndexOf("," & e.Values("headerName").ToString.ToLower & ",") < 0 Then
            messageText.Append("Column name as '" & e.Values("headerName") & "' was NOT found in [ detail order ] file! You input a wrong column name.<br />")

        End If


        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If
    End Sub


    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    ''' 
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating

        Dim messageText As StringBuilder = New StringBuilder()


        If String.IsNullOrEmpty(e.NewValues("remarkToBeAdded")) Then
            messageText.Append("Empty value  in field remarkToBeAdded is not allowed.")
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


    
End Class

