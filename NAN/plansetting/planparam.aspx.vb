Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Linq



Partial Class plansetting_planparam
    Inherits InteracWithExcel



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        
        maxRowNumber = 501

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
        signalRstPlanParam()
    End Sub


    Protected Sub UpldInsrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldInsrt.Click

        UpldInsrt_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        signalRstPlanParam()

    End Sub


    '''update table per the data in excel file
    Protected Sub UpldUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldUpdate.Click

        UpldUpdate_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        signalRstPlanParam()
    End Sub

    Protected Sub overwrite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles overwrite.Click

        overwrite_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        signalRstPlanParam()
    End Sub


    Public Overrides Function dataValidityCheck(ByRef dtTbl As DataTable) As String



        Dim msstring As String = String.Empty

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

        For Each de As DictionaryEntry In e.NewValues

            If de.Value Is Nothing OrElse de.Value.ToString().Trim().Length = 0 Then
                messageText.Append("Empty value  is not allowed.<br />")
                Exit For
            End If

            If e.Keys("paramname").ToString().ToLower.IndexOf("time") > -1 Then
                If Not IsDate(de.Value) Then
                    messageText.Append("A date is supposed to be entered.<br />")
                    Exit For
                End If
            End If


            If e.Keys("paramname").ToString().ToLower.IndexOf("day") > -1 Then
                If Not IsNumeric(de.Value) Then
                    messageText.Append("A number is supposed to be entered.<br />")
                    Exit For
                End If
            End If


            If e.Keys("paramname").ToString().ToLower.IndexOf("auto") > -1 Then
                If de.Value.ToString.ToLower <> "yes" AndAlso de.Value.ToString.ToLower <> "no" Then
                    messageText.Append("'yes' or 'no' is supposed to be entered.<br />")
                    Exit For
                End If
            End If



        Next



        If messageText.Length > 0 Then
            e.Cancel = True
            Message.Text = messageText.ToString()
        End If


    End Sub

    Protected Sub SDS1_Deleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles SDS1.Deleted
        signalRstPlanParam()
    End Sub

    Protected Sub SDS1_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles SDS1.Inserted
        signalRstPlanParam()
    End Sub


    Protected Sub SqlDataSource1_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles SDS1.Updated

        ' reset this application variable to force to refetch the planning parameter in the DataExchange.aspx.vb function

        signalRstPlanParam
    End Sub


    Protected Sub signalRstPlanParam()
        Cache.Remove("production")
        CacheRemove("checker")

        CacheRemove("schedule")

    End Sub

    
End Class

