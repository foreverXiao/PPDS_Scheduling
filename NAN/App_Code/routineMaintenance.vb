Imports Microsoft.VisualBasic
Imports System.Web.Caching
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data

Public Class PGPTbatchAndOrder
    Public organization As String
    Public strFTPserverIP As String
    Public strFTP_ID As String
    Public strFTP_PW As String
    Public strOpenOrderPrefix As String
    Public strOpenOrderPath As String
    Public strBatchStatusUpdatePrefix As String
    Public strBatchStatusUpdatePath As String
End Class


Public Class routineMaintenance
    Inherits System.Web.UI.Page


    Dim cache1 As System.Web.Caching.Cache
    Dim PGPTbatchAndOrder1 As PGPTbatchAndOrder

    Sub New(ByRef Cache2 As System.Web.Caching.Cache, ByRef PGPTbatchAndOrder2 As PGPTbatchAndOrder)
        cache1 = Cache2
        PGPTbatchAndOrder1 = PGPTbatchAndOrder2
    End Sub

    Private Sub New()

    End Sub

    Public Sub routineMaintenance()


        If Today.DayOfWeek = DayOfWeek.Friday Then
            clearAgedScrewMark()
            timerToCompctAfterExplant()
        End If


        deleteEDIfilesForPGPT()  'delete EDI files for orgnization PGPT


    End Sub



    ''' <summary>
    ''' set a timer to compact Access database after ex-plant date is uploaded to ftp server
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub timerToCompctAfterExplant()
        'Dim onRemove As CacheItemRemovedCallback
        'onRemove = New CacheItemRemovedCallback(AddressOf RemovedCallback)
        'cache1.Insert("compactAccess", "go", Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(480), CacheItemPriority.Default, onRemove)
    End Sub


    Protected Sub RemovedCallback(ByVal k As String, ByVal v As Object, ByVal r As CacheItemRemovedReason)

        'If Today.AddHours(19).CompareTo(Now()) < 0 AndAlso Today.AddHours(22).CompareTo(Now()) > 0 Then

        'Dim pathAndFile As String = String.Empty
        'For Each str1 In ConfigurationManager.ConnectionStrings("accessDB").ConnectionString.Split(";".ToCharArray())
        '    If str1.ToLower.IndexOf("data source") > -1 Then
        '        pathAndFile = str1.Substring(str1.ToLower.IndexOf("=") + 1).Trim()
        '        Exit For
        '    End If
        'Next

        'compactAccess(pathAndFile, 32 * 1024 * 1024)

        'For Each str1 In ConfigurationManager.ConnectionStrings("accessDB").ConnectionString.Split(";".ToCharArray())
        '    If str1.ToLower.IndexOf("data source") > -1 Then
        '        pathAndFile = str1.Substring(str1.ToLower.IndexOf("=") + 1).Trim()
        '        Exit For
        '    End If
        'Next

        'compactAccess(pathAndFile, 4 * 1024 * 1024)

        'End If
    End Sub

    'compact access database when its file size grows to certain number
    Protected Sub compactAccess(ByVal pathAndFile As String, Optional ByVal fileSize As Long = 20 * 1024 * 1024)
        If File.Exists(pathAndFile) Then

            Dim fileInfo1 As FileInfo = New FileInfo(pathAndFile)

            If fileInfo1.Length > fileSize Then 'greater than certain quantity

                Dim pathAndFile1 As String = String.Empty
                pathAndFile1 = pathAndFile.Substring(0, pathAndFile.LastIndexOf("\") + 1) & "temp" & pathAndFile.Substring(pathAndFile.LastIndexOf("\") + 1)

                If File.Exists(pathAndFile1) Then
                    File.Delete(pathAndFile1)
                End If

                Dim compactAccess As Object = CreateObject("DAO.DBEngine.120")
                compactAccess.CompactDatabase(pathAndFile, pathAndFile1)
                'compactAccess.CompactDatabase(pathAndFile, pathAndFile1, , 128)

                File.Delete(pathAndFile)
                File.Move(pathAndFile1, pathAndFile)

            End If

        End If


    End Sub


    ''' <summary>
    ''' clear the mark of screw pulling if the order was arranged production 3 months ago or those orders which could not be found in the order details
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub clearAgedScrewMark()
        'get data from table Esch_Na_tbl_batch_no_group_and_batch_rules
        Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("accessDB").ProviderName & ConfigurationManager.ConnectionStrings("accessDB").ConnectionString)
        Dim dtAdapterFrom As SqlDataAdapter = New SqlDataAdapter("SELECT txt_order_key FROM Esch_Na_tbl_orders ", conn)
        Dim dtTableFrom As DataTable = New DataTable
        dtAdapterFrom.Fill(dtTableFrom)

        Dim dtAdapterTo1 As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM Esch_Na_tbl_prductn_prmtr ", conn)
        Dim cmdbAccessCmdBuilder1 As New SqlCommandBuilder(dtAdapterTo1)
        dtAdapterTo1.DeleteCommand = cmdbAccessCmdBuilder1.GetDeleteCommand()
        Dim dtTableTo1 As DataTable = New DataTable
        dtAdapterTo1.Fill(dtTableTo1)
        Dim keys(0) As DataColumn
        keys(0) = dtTableTo1.Columns("txt_order_key")
        dtTableTo1.PrimaryKey = keys



        'iterate through screw pulling table to see if we should delete this record
        For i As Integer = 0 To dtTableTo1.Rows.Count - 1

            Dim rs() As DataRow = dtTableFrom.Select("txt_order_key = '" & dtTableTo1.Rows(i).Item("txt_order_key") & "'")

            If rs.Count = 0 Then
                dtTableTo1.Rows(i).Delete()
            End If

        Next



        dtAdapterTo1.Update(dtTableTo1)



        cmdbAccessCmdBuilder1.Dispose()

        dtTableTo1.Dispose()

        dtTableFrom.Dispose()

        dtAdapterTo1.Dispose()

        dtAdapterFrom.Dispose()

        'conn.Close()
        conn.Dispose()

    End Sub

    ''' <summary>
    ''' This routine is used to delete all the EDI files for organization PGPT because these EDI is useless
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub deleteEDIfilesForPGPT()

        Try

            'get user name and password and ftp server IP to log on to server to get files by ftp method
            Dim organization As String = PGPTbatchAndOrder1.organization
            Dim serverUri As String = PGPTbatchAndOrder1.strFTPserverIP
            If Not serverUri.StartsWith("ftp://") Then serverUri = "ftp://" & serverUri
            If Not serverUri.EndsWith("/") Then serverUri &= "/"

            '1,get new&revision order status files from ftp server ==================================== 
            Dim prefix As String = PGPTbatchAndOrder1.strOpenOrderPrefix
            Dim workingDirectory As String = PGPTbatchAndOrder1.strOpenOrderPath
            If Not workingDirectory.EndsWith("/") Then workingDirectory &= "/"

            Dim olist As List(Of String) = New List(Of String)

            Dim ftpoperation As FTPcls = New FTPcls(PGPTbatchAndOrder1.strFTP_ID, PGPTbatchAndOrder1.strFTP_PW, serverUri & workingDirectory)
            Dim errMsg As String = String.Empty
            olist = ftpoperation.FileListInDirContains(prefix & "." & organization, errMsg) 'Get the file list from ftp server


            olist.Sort(StringComparer.Ordinal)

            'delete the file on the ftp server
            For Each filename As String In olist
                ftpoperation.DeleteFileOnServer(filename)
            Next


            '2,get  batch status files from ftp server ==================================== 
            prefix = PGPTbatchAndOrder1.strBatchStatusUpdatePrefix
            workingDirectory = PGPTbatchAndOrder1.strBatchStatusUpdatePath
            If Not workingDirectory.EndsWith("/") Then workingDirectory &= "/"



            ftpoperation = New FTPcls(PGPTbatchAndOrder1.strFTP_ID, PGPTbatchAndOrder1.strFTP_PW, serverUri & workingDirectory)
            olist = ftpoperation.FileListInDirContains(prefix & "." & organization, errMsg) 'Get the file list from ftp server


            olist.Sort(StringComparer.Ordinal)

            'delete the file on the ftp server
            For Each filename As String In olist
                ftpoperation.DeleteFileOnServer(filename)
            Next


        Finally

        End Try


    End Sub



End Class
