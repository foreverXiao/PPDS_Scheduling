Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class Makerelated_colorCode
    Inherits InteracWithExcel

    Public Delegate Sub asychrSub()

    Private colorList() As String = {"black", "silver", "gray", "white", "maroon", "red", "purple", "fuchsia", "green", "lime", "olive", "yellow", "navy", "blue", "teal", "aqua"}


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SDS1.ConnectionString = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString


        maxRowNumber = 15001



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
        rewriteCSSfile()

    End Sub


    Protected Sub UpldInsrt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldInsrt.Click

        UpldInsrt_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        rewriteCSSfile()


    End Sub


    '''update table per the data in excel file
    Protected Sub UpldUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UpldUpdate.Click

        UpldUpdate_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        rewriteCSSfile()

    End Sub

    Protected Sub overwrite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles overwrite.Click

        overwrite_ClickHnadler(SDS1, LV1, StatusLabel, FileUpload1)
        rewriteCSSfile()

    End Sub


    Protected Sub SDS1_Deleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles SDS1.Deleted
        rewriteCSSfile()
    End Sub

    Protected Sub SDS1_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles SDS1.Inserted
        rewriteCSSfile()
    End Sub

    Protected Sub SDS1_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles SDS1.Updated
        rewriteCSSfile()
    End Sub


    


    Public Overrides Function dataValidityCheck(ByRef dtTbl As DataTable) As String


        Dim messageText As StringBuilder = New StringBuilder()

        Try
            For Each rowexcel As DataRow In dtTbl.Rows
                If DBNull.Value.Equals(rowexcel("red")) OrElse DBNull.Value.Equals(rowexcel("green")) OrElse DBNull.Value.Equals(rowexcel("blue")) Then
                    messageText.AppendLine("Empty value in field red,green and blue is not allowed.<br />")
                    Exit For
                End If

                If CInt(rowexcel("red")) < 0 OrElse CInt(rowexcel("red")) > 255 OrElse CInt(rowexcel("green")) < 0 OrElse CInt(rowexcel("green")) > 255 OrElse CInt(rowexcel("blue")) < 0 OrElse CInt(rowexcel("blue")) > 255 Then
                    messageText.AppendLine("Value for (red,green,blue) should be between 0 and 255 .<br />")
                    Exit For
                End If

                If False AndAlso Not DBNull.Value.Equals(rowexcel("commonName")) AndAlso Not colorList.Contains(rowexcel("commonName").ToString.ToLower()) Then

                    messageText.AppendLine("Color name should be one in the list (" & String.Join(",", colorList))

                    'For i As Integer = 1 To colorList.Count - 1
                    '    messageText.Append("," & colorList(i))
                    'Next

                    'messageText.Append(").")

                End If

            Next

        Catch ex As Exception
            messageText.AppendLine("There is something wrong with the data validity." & ex.Message)
        End Try

        Return messageText.ToString

    End Function

    Protected Sub LV1_ItemInserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ListViewInsertEventArgs) Handles LV1.ItemInserting
        ' Cancel the update operation if any of the fields is empty
        ' or null.
        Dim messageText As StringBuilder = New StringBuilder()

        For Each de As DictionaryEntry In e.Values
            ' Check if the value is null or empty except field txtRemark
            If (de.Value Is Nothing OrElse de.Value.ToString().Trim().Length = 0) And Not (de.Key.ToString = "txtRemark") Then
                messageText.Append("Cannot insert an empty value.<br />")
            End If
        Next


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

        If String.IsNullOrEmpty(e.NewValues("red")) OrElse String.IsNullOrEmpty(e.NewValues("green")) OrElse String.IsNullOrEmpty(e.NewValues("blue")) Then
            messageText.AppendLine("Empty value in field red,green and blue is not allowed.<br />")
        Else
            If CInt(e.NewValues("red")) < 0 OrElse CInt(e.NewValues("red")) > 255 OrElse CInt(e.NewValues("green")) < 0 OrElse CInt(e.NewValues("green")) > 255 OrElse CInt(e.NewValues("blue")) < 0 OrElse CInt(e.NewValues("blue")) > 255 Then
                messageText.AppendLine("Value for (red,green,blue) should be between 0 and 255 .<br />")

            End If
        End If


        If Not String.IsNullOrEmpty(e.NewValues("commonName")) AndAlso Not colorList.Contains(e.NewValues("commonName").ToString.ToLower()) Then

            messageText.AppendLine("Color name should be one in the list (" & String.Join(",", colorList) & ")")

            'For i As Integer = 1 To colorList.Count - 1
            '    messageText.Append("," & colorList(i))
            'Next

            'messageText.Append(").")

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


    'asynchronous function to re-write color scheme for Gantt block
    Protected Sub rewriteCSSfile()
        'use below function to update table 
        Dim asySubroutine As asychrSub = New asychrSub(AddressOf updateGanttColorScheme)
        asySubroutine.BeginInvoke(Nothing, Nothing)
    End Sub

    'update the content of file GanttChartColor.css  in folder ../dragdrop
    Protected Sub updateGanttColorScheme()

        Dim cssSTR As StringBuilder = New StringBuilder()

        Using connCSS As SqlConnection = New SqlConnection(SDS1.ConnectionString)

            connCSS.Open()

            Dim command As New SqlCommand(SDS1.SelectCommand.ToString(), connCSS)
            Dim reader As SqlDataReader
            reader = command.ExecuteReader()

            Do While reader.Read()
                cssSTR.Append(".g-" & reader("color"))
                If CBool(reader("tranparent")) Then
                    cssSTR.AppendLine(" {background-color:transparent;background-image:url(""trprt.gif"");}" & vbCr)
                Else
                    If DBNull.Value.Equals(reader("commonName")) Then ' see if there is null value in field commonName
                        cssSTR.AppendLine(" {background-color:rgb(" & reader("red") & "," & reader("green") & "," & reader("blue") & ");}" & vbCr)
                    Else
                        cssSTR.AppendLine(" {background-color:" & reader("commonName") & ";}" & vbCr)
                    End If
                End If

            Loop


            reader.Close()

            command.Dispose()

            connCSS.Close()

        End Using


        'write to GanttColor.css file
        Dim pathAndName As String = Server.MapPath("~/App_Themes/GanttChartColor.css")
        'If File.Exists(pathAndName) Then
        '    File.Delete(pathAndName)
        'End If


        Using outFile As New StreamWriter(pathAndName)
            outFile.Write(cssSTR.ToString)
        End Using


    End Sub


    
End Class

