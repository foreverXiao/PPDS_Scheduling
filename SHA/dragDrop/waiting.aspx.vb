
Imports System.Text
Imports System.Data
Imports Microsoft.VisualBasic




Partial Class dragDrop_waiting
    Inherits basepage1




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        StatusLabel.Text = "Waiting..... another page is importing order status now."


    End Sub



    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim compactAccess As Object = CreateObject("DAO.DBEngine.120")
        'compactAccess.CompactDatabase("C:\inetpub\wwwroot\test\App_Data\db_Resin.accdb", "C:\inetpub\wwwroot\test\App_Data\db_Resin1.accdb", , 128)

        compactAccess.CompactDatabase("C:\inetpub\wwwroot\test\App_Data\db_Resin.accdb", "C:\inetpub\wwwroot\test\App_Data\db_Resin1.accdb")


        'Dim a As String = ConfigurationManager.ConnectionStrings("accessDB").ConnectionString

        'Dim b() As String = a.Split(";".ToCharArray)
        'Dim pathAndFile As String
        'For Each str1 As String In b
        '    If str1.ToLower.IndexOf("data source") > -1 Then
        '        pathAndFile = str1.Trim()
        '        Exit For
        '    End If
        'Next


        compactAccess = Nothing

    End Sub
End Class

