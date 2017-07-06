Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class Makerelated_scheduleDifferencesInBetweens
    Inherits InteracWithExcel

    


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


        Return String.Empty

    End Function


    ''' <summary>
    ''' You can do some checking on data validity; if some exception is thrown out, the updating would be cancelled
    ''' </summary>
    ''' 
    Protected Sub LV1_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewUpdateEventArgs) Handles LV1.ItemUpdating

        Dim messageText As StringBuilder = New StringBuilder()

        If String.IsNullOrEmpty(e.NewValues("int_rate")) Then
            messageText.Append("Empty value in field int_rate is not allowed.")
        Else
            If Not IsNumeric(e.NewValues("int_rate")) Then messageText.Append("Number is needed in field int_rate.")
        End If

        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If


    End Sub



   

    

End Class

