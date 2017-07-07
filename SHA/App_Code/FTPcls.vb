Imports Microsoft.VisualBasic
Imports System.Net
Imports System.IO
Imports basepage1

'this class is to implement ftp functions such as : download file;upload file; list files name in ftp server directory;delete file in ftp server

Public Class FTPcls


    Private _username As String = String.Empty
    Private _password As String = String.Empty
    Private _ftpWebDirectory As String = String.Empty
    Private _wbrequest As FtpWebRequest = Nothing

    Public Property username() As String
        Get
            Return _username
        End Get
        Set(ByVal value As String)
            _password = value
        End Set
    End Property

    Public Property password() As String
        Get
            Return _password
        End Get
        Set(ByVal value As String)
            _password = value
        End Set
    End Property

    Public Property serveruri() As String  'input string and convert it into Uri object
        Set(ByVal value As String)
            If value.StartsWith("ftp://") Then
                If value.EndsWith("/") Then
                    _ftpWebDirectory = value
                Else
                    _ftpWebDirectory = value & "/"
                End If
            End If

        End Set
        Get
            Return _ftpWebDirectory
        End Get
    End Property


    Public Sub New(ByVal username As String, ByVal password As String, ByVal serveruri As String)
        _username = username
        _password = password
        Me.serveruri = serveruri

    End Sub


    'delete file
    Public Function DeleteFileOnServer(ByVal serverFileName As String) As String

        Try
            ' Get the object used to communicate with the server.
            If Not String.IsNullOrEmpty(_ftpWebDirectory) Then
                _wbrequest = CType(WebRequest.Create(_ftpWebDirectory & serverFileName), FtpWebRequest)
                _wbrequest.Credentials = New NetworkCredential(_username, _password)
                _wbrequest.Proxy = Nothing

                _wbrequest.Method = WebRequestMethods.Ftp.DeleteFile
                Dim wbresponse As FtpWebResponse = CType(_wbrequest.GetResponse(), FtpWebResponse)

                wbresponse.Close()

            Else
                Throw New Exception("Invalid ftp server directory !")
            End If


        Catch e As Exception
            Return e.Message
        End Try


        Return "true"

    End Function

    'list out all the files in the directory
    Public Function FileListInDirContains(ByVal filter As String, ByRef errMsg As String) As List(Of String)

        ' Get the object used to communicate with the server.
        Dim oList As New List(Of String)

        Try

            If Not String.IsNullOrEmpty(_ftpWebDirectory) Then

                _wbrequest = CType(WebRequest.Create(_ftpWebDirectory), FtpWebRequest)
                _wbrequest.Credentials = New NetworkCredential(_username, _password)
                _wbrequest.Proxy = Nothing
                _wbrequest.Method = WebRequestMethods.Ftp.ListDirectory
                Dim wbresponse As FtpWebResponse = CType(_wbrequest.GetResponse(), FtpWebResponse)

                Dim sr As StreamReader = New StreamReader(wbresponse.GetResponseStream)
                Dim str As String = sr.ReadLine

                While str IsNot Nothing
                    If str.ToLower().IndexOf(filter.ToLower()) >= 0 Then oList.Add(str)
                    str = sr.ReadLine

                End While

                wbresponse.Close()

            Else
                Throw New Exception("Invalid ftp server directory !")
            End If

        Catch e As Exception
            errMsg = e.Message
        End Try

        Return oList

    End Function

    'Upload one file to server web directory
    Public Function UploadFileToServer1(ByVal localfile As String, ByVal serverFileName As String) As String

        ' Get the object used to communicate with the server.

        Try

            If Not String.IsNullOrEmpty(_ftpWebDirectory) Then
                _wbrequest = CType(WebRequest.Create(_ftpWebDirectory & serverFileName), FtpWebRequest)
                _wbrequest.Credentials = New NetworkCredential(_username, _password)
                _wbrequest.Proxy = Nothing
                _wbrequest.Method = WebRequestMethods.Ftp.UploadFile


                ' Copy the contents of the file to the request stream.
                Dim sourceStream As StreamReader = New StreamReader(localfile)
                Dim fileContents() As Byte = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd())
                sourceStream.Close()

                _wbrequest.ContentLength = fileContents.Length

                Dim requestStream As Stream = _wbrequest.GetRequestStream()
                requestStream.Write(fileContents, 0, fileContents.Length)
                requestStream.Close()

                Dim wbresponse As FtpWebResponse = CType(_wbrequest.GetResponse(), FtpWebResponse)

                wbresponse.Close()

            Else
                Throw New Exception("Invalid ftp server directory !")
            End If

        Catch e As Exception
            Return e.Message
        End Try

        Return "true"

    End Function

    'Upload one file to server web directory
    Public Function UploadFileToServer(ByVal localfile As String, ByVal serverFileName As String) As String

        ' Get the object used to communicate with the server.

        Try

            If Not String.IsNullOrEmpty(_ftpWebDirectory) Then
                _wbrequest = CType(WebRequest.Create(_ftpWebDirectory & serverFileName), FtpWebRequest)
                _wbrequest.Credentials = New NetworkCredential(_username, _password)
                _wbrequest.Proxy = Nothing
                _wbrequest.Method = WebRequestMethods.Ftp.UploadFile



                Dim sourceStream As FileStream = New FileStream(localfile, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                ' Copy the contents of the file to the request stream.


                Dim buffer(16 * 1024) As Byte
                Dim ms As MemoryStream = New MemoryStream()
                Dim intRead As Integer = 0
                intRead = sourceStream.Read(buffer, 0, buffer.Length)
                While intRead > 0
                    ms.Write(buffer, 0, intRead)
                    intRead = sourceStream.Read(buffer, 0, buffer.Length)
                End While

                Dim fileContents() As Byte = ms.ToArray()
                ms.Close()

                sourceStream.Close()
                '_wbrequest.ContentLength = fileContents.Length
                _wbrequest.UseBinary = True

                Dim requestStream As Stream = _wbrequest.GetRequestStream()
                requestStream.Write(fileContents, 0, fileContents.Count)

                requestStream.Close()

                Dim wbresponse As FtpWebResponse = CType(_wbrequest.GetResponse(), FtpWebResponse)



            Else
                Throw New Exception("Invalid ftp server directory !")
            End If

        Catch e As Exception
            Return e.Message
        Finally

        End Try

        Return "true"

    End Function

    'Download one file from server web directory
    Public Function DownFileFrmServer1(ByVal localfile As String, ByVal serverFileName As String) As String

        ' Get the object used to communicate with the server.

        Try

            If Not String.IsNullOrEmpty(_ftpWebDirectory) Then

                _wbrequest = CType(WebRequest.Create(_ftpWebDirectory & serverFileName), FtpWebRequest)
                _wbrequest.Credentials = New NetworkCredential(_username, _password)
                _wbrequest.Proxy = Nothing

                'Dim newFileData() As Byte = _wbrequest.DownloadData(serveruri.ToString())
                'Dim fileString As String = Encoding.UTF8.GetString(newFileData)

                _wbrequest.Method = WebRequestMethods.Ftp.DownloadFile

                Dim wbresponse As FtpWebResponse = DirectCast(_wbrequest.GetResponse(), FtpWebResponse)
                Dim responseStream As Stream = wbresponse.GetResponseStream()
                Dim reader As New StreamReader(responseStream)

                Dim sw As StreamWriter = New StreamWriter(localfile)

                Dim line As String = reader.ReadToEnd()

                'line = reader.ReadLine()
                ' While line IsNot Nothing
                sw.Write(line)
                ' line = reader.ReadLine()
                ' End While


                reader.Close()
                sw.Close()

                wbresponse.Close()


            End If

        Catch e As Exception
            Return e.Message
        End Try

        Return "true"

    End Function


    'Download one file from server web directory
    Public Function DownFileFrmServer(ByVal localfile As String, ByVal serverFileName As String) As String

        ' Get the object used to communicate with the server.

        Try

            If Not String.IsNullOrEmpty(_ftpWebDirectory) Then

                _wbrequest = CType(WebRequest.Create(_ftpWebDirectory & serverFileName), FtpWebRequest)
                _wbrequest.Credentials = New NetworkCredential(_username, _password)
                _wbrequest.Proxy = Nothing

                'Dim newFileData() As Byte = _wbrequest.DownloadData(serveruri.ToString())
                'Dim fileString As String = Encoding.UTF8.GetString(newFileData)

                _wbrequest.Method = WebRequestMethods.Ftp.DownloadFile
                _wbrequest.UseBinary = True

                Dim wbresponse As FtpWebResponse = DirectCast(_wbrequest.GetResponse(), FtpWebResponse)
                Dim responseStream As Stream = wbresponse.GetResponseStream()
                'Dim reader As New StreamReader(responseStream)
                'Dim icount As Integer = 1000

                Dim buffer(16 * 1024) As Byte
                Dim ms As MemoryStream = New MemoryStream()
                Dim intRead As Integer = 0
                intRead = responseStream.Read(buffer, 0, buffer.Length)
                While intRead > 0
                    ms.Write(buffer, 0, intRead)
                    intRead = responseStream.Read(buffer, 0, buffer.Length)
                End While

                Dim fileContents() As Byte = ms.ToArray()
                ms.Close()
                'responseStream.Read(fileContents, 0, fileContents.Count - 1)


                If File.Exists(localfile) Then File.Delete(localfile)
                Dim sw As FileStream = New FileStream(localfile, System.IO.FileMode.CreateNew, System.IO.FileAccess.ReadWrite)
                sw.Write(fileContents, 0, fileContents.Count)

                responseStream.Close()
                sw.Close()

                wbresponse.Close()


            End If

        Catch e As Exception
            Return e.Message
        End Try

        Return "true"

    End Function




End Class
