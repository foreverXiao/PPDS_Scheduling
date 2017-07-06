Imports basepage1
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
imports System.Collections.Generic



Partial Class interface_orderFilesEDI
    Inherits basepage1


    Private Sub BindGrid()

        Dim contnueOrNot As Boolean = True
        Dim msg As String = String.Empty
        Dim localpath As String = ConfigurationManager.AppSettings("EDIfilesFolder")
        If Not localpath.EndsWith("\") Then localpath &= "\"
        localpath &= "interfaceData\"
        If Not Directory.Exists(localpath) Then
            Try
                Directory.CreateDirectory(localpath)
            Catch
                'errMessage.Append("You might not have right to create folder: " & localpath)
                msg = "<div style = 'color:red;'>" & "You might not have right to create folder: " & localpath & "</div>"
            End Try
        End If


        'get new&revision order status files from ftp server and also copy them to S&FS SBU folder for their use
        Dim localpathsubfolder As String = localpath & "NewRevisionOrder\"

        If Not Directory.Exists(localpathsubfolder) Then
            Try
                Directory.CreateDirectory(localpathsubfolder)
            Catch
                'errMessage.Append("You might not have right to create folder: " & localpath)
                msg = "<div style = 'color:red;'>" & "You might not have right to create folder: " & localpathsubfolder & "</div>"
            End Try
        End If


        Dim files As FileInfo() = New DirectoryInfo(localpathsubfolder).GetFiles(valueOf("strOpenOrderPrefix") & "." & valueOf("strOrganization") & "*")
        'Dim sortedFiles() As String = CType(files, String())
        'Dim compare1 As compare

        'files.Sort(AddressOf Compare1)

        Array.Sort(files, AddressOf Compare)

        'Dim sorted = Directory.GetFiles(".").OrderBy(f >= f)

        GridView1.DataSource = files
        GridView1.DataBind()

        sl.Text = msg

    End Sub

    Private Shared Function Compare(ByVal x As Object, ByVal y As Object) As Integer
        Dim file1 As FileInfo = CType(x, FileInfo)
        Dim file2 As FileInfo = CType(y, FileInfo)

        Return file2.Name.CompareTo(file1.Name)
    End Function



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BindGrid()
        End If
    End Sub

    Private Sub Downloadfile(ByVal fileName As String, ByVal FullFilePath As String)
        Response.AddHeader("Content-Disposition", "attachment; filename=" & fileName)
        Response.TransmitFile(FullFilePath)
        Response.End()
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandName = "Download" Then
            Dim fileInfo() As String = e.CommandArgument.ToString().Split(";")
            Dim FileName As String = fileInfo(1)
            Dim FullPath As String = fileInfo(0)
            Downloadfile(FileName, FullPath)
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
        BindGrid()

    End Sub


End Class
