Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class SCMrelated_spcluserackrule
    Inherits FrequentPlanActions

 


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        
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
                If String.IsNullOrEmpty(rowexcel("txt_item_no")) OrElse String.IsNullOrEmpty(rowexcel("txt_end_user")) Then
                    msstring = "txt_item_no and txt_end_user are  not allowed empty value."
                    Exit For
                End If

                If Not IsNumeric(rowexcel("int_extraDays")) Then
                    msstring = "There is something wrong with row which txt_item_no = " & rowexcel("txt_item_no") & " and txt_end_user = " & rowexcel("txt_end_user") & " ! Maybe int_extraDays is not a valid number."
                    Exit For
                End If
            Next
        Catch
            msstring = "There is something wrong with the data validity."
        End Try

        Return msstring

    End Function

    Protected Sub LV1_ItemUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdatedEventArgs) Handles LV1.ItemUpdated
        Dim index1 As Integer = LV1.EditIndex

        Dim DK As DataKey = LV1.DataKeys(index1)

        Message.Text = DoCalculateRSD(" And (txt_item_no = '" & DK.Item("txt_item_no") & "')")
    End Sub


    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    ''' 
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating

        Dim messageText As StringBuilder = New StringBuilder()

        If String.IsNullOrEmpty(e.NewValues("int_extraDays")) Then
            messageText.Append("Empty value in field int_extraDays is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("int_extraDays")) Then messageText.Append("Number is needed in field int_extraDays.")
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

        Message.Text = DoCalculateRSD(" And (txt_item_no = '" & e.Values("txt_item_no") & "')")

    End Sub


End Class

