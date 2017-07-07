Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO


Partial Public Class basepage1
    Inherits System.Web.UI.Page

    Enum priority  'show the priority of operation on table Esch_CQ_tbl_orders
        ganttChart = 5
        addMTI = 6
        excelUpload = 6
        CombineOrBatchCreation = 6
        ImportNewOrderOrBatch = 6
        UploadExplantDate = 6
    End Enum

    Private keytableName As String = "Esch_CQ_tbl_orders"

    Public Const dateSeparator = "'"  'This is for SQL server date
    Public Const newOrderStatus As String = "NEW2" 'Remark for newly imported orders

    Private _rightlevel As Integer = 8

    Private _DBconnection As String = "accessDB"
    Private _paramDBconnection As String = "accessDB"


    ''' <summary>
    ''' connection name to get connection string in web.config
    ''' </summary>>
    Public Property dbConnectionName() As String
        Get
            Return _DBconnection
        End Get
        Set(ByVal value As String)
            _DBconnection = value
        End Set
    End Property


    ''' <summary>
    ''' param database connection name to get connection string in web.config
    ''' </summary>>
    Public Property dbConnForParam() As String
        Get
            Return _paramDBconnection
        End Get
        Set(ByVal value As String)
            _paramDBconnection = value
        End Set
    End Property



    Public Property accesslevel() As Integer
        Get
            Return _rightlevel
        End Get
        Protected Set(ByVal value As Integer)
            If value < 100 Then
                _rightlevel = value
            Else
                _rightlevel = 8
            End If

        End Set
    End Property

    Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) _
        Handles Me.Init


    End Sub

    Public Sub msgPopUP(ByVal message1 As String, ByRef label101 As Object, Optional ByVal warning_color As Boolean = True, Optional ByVal popup_window As Boolean = True)
        If label101 Is Nothing Then
            Return
        End If
        label101.ForeColor = IIf(warning_color, Drawing.Color.Red, Drawing.Color.Black)
        label101.Text = message1
        If popup_window Then Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "script1", "<script language='javascript'>alert('" & message1 & "');</script>", False)
    End Sub

    Public Overridable Function pageTitleReference() As String
        'Utilize page's title
        If String.IsNullOrEmpty(Trim(Page.Title)) Then
            Return "NoTitle"
        Else
            Return Replace(Page.Title, " ", "")
        End If


    End Function

    ''' <summary>
    ''' list out all the production lines
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub initiateProductionLines()

        If CacheFrom("arrayOfLines") Is Nothing Then
            Dim connParam As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)
            connParam.Open()

            Dim command As New SqlCommand("Select distinct int_line_no From Esch_CQ_tbl_output_by_line_only ORDER BY int_line_no ASC", connParam)

            Dim reader As SqlDataReader

            reader = command.ExecuteReader()


            Dim j1 As Long = 0
            Dim arrayLine1() As Integer
            Do While reader.Read()
                ReDim Preserve arrayLine1(j1)
                arrayLine1(j1) = CInt(reader("int_line_no"))
                j1 += 1
            Loop
            reader.Close()

            Dim i As Integer = arrayLine1.Count   'how many lines

            'sort array
            Dim Sorted As Boolean, Temp As Integer
            Sorted = False
            Do While Not Sorted
                Sorted = True
                For j1 = 0 To (i - 2)
                    If arrayLine1(j1) > arrayLine1(j1 + 1) Then
                        Temp = arrayLine1(j1 + 1)
                        arrayLine1(j1 + 1) = arrayLine1(j1)
                        arrayLine1(j1) = Temp
                        Sorted = False
                    End If
                Next
            Loop

            connParam.Close()
            connParam.Dispose()

            CacheInsert("arrayOfLines", arrayLine1)

        End If

    End Sub

    ''' <summary>
    ''' list out the production liness
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function arrayOfLines() As Integer()

        initiateProductionLines()

        Return CType(CacheFrom("arrayOfLines"), Integer())

    End Function



    ''' <summary>
    ''' get the value for the specific system variable  according to system variable table
    ''' </summary>
    ''' <param name="name">the parameter's name</param>
    ''' <returns>return a string,but you can convert the string to any type of data as you want</returns>
    Public Function valueOf(ByVal name As String) As String


        'if found in cathe, return from cache, otherwise create a new cache

        initiateCacheSystmVrbl()

        Dim cacheDictnry As Dictionary(Of String, String) = CType(CacheFrom("systemvariable"), Dictionary(Of String, String))
        If cacheDictnry.ContainsKey(name) Then
            Return cacheDictnry(name)
        Else
            Return String.Empty
        End If


    End Function



    ''' <summary>
    ''' to initiate cache variable from table Esch_CQ_tbl_system_variable
    ''' </summary>
    Public Sub initiateCacheSystmVrbl()

        If CacheFrom("systemvariable") Is Nothing Then

            Dim connstr As String = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString

            Dim conn As SqlConnection = New SqlConnection(connstr)
            Dim command As SqlCommand = New SqlCommand("SELECT txtVariableValue , txtVariableName  FROM  [Esch_CQ_tbl_system_variable]", conn)
            Dim reader As SqlDataReader


            Dim cacheDictionary As New Dictionary(Of String, String)

            Try
                conn.Open()
                reader = command.ExecuteReader()

                While reader.Read()
                    cacheDictionary.Add(reader("txtVariableName"), reader("txtVariableValue"))
                End While


            Finally
                reader.Close()
                command.Dispose()
                conn.Close()
                conn.Dispose()
            End Try


            CacheInsert("systemvariable", cacheDictionary)

        End If



    End Sub

    Overloads Sub CacheInsert(ByVal keyStr As String, ByRef obj As Object)
        Cache.Insert(userIden() & keyStr, obj, Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(480))
    End Sub

    Overloads Sub CacheInsert(ByVal keyStr As String, ByRef obj As Object, ByVal sliding As Integer)

        Cache.Insert(userIden() & keyStr, obj, Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(sliding))
    End Sub


    Sub CacheRemove(ByVal keyStr As String)

        Cache.Remove(userIden() & keyStr)
    End Sub

    Function CacheFrom(ByVal keyStr As String) As Object
        Return Cache(userIden() & keyStr)
    End Function

    Function userIden() As String
        If IsNothing(Request.Cookies("userInfo")) Then
            Return String.Empty
        Else
            Return Server.HtmlEncode(Request.Cookies("userInfo")("id"))
        End If
    End Function


    Function lockKeyTable(ByVal prrty As priority) As String 'who is locking the table
        Dim tablePriority As New NameValueCollection
        If Cache(keytableName) Is Nothing Then
            tablePriority.Add("p", prrty)
            tablePriority.Add("id", userIden())
            Cache.Insert(keytableName, tablePriority, Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(3))
            Return ""
        End If

        tablePriority = CType(Cache(keytableName), NameValueCollection)
        If prrty <= CInt(tablePriority.Item("p")) AndAlso priority.ganttChart <> CInt(tablePriority.Item("p")) Then
            Return tablePriority.Item("id")
        Else
            tablePriority.Item("p") = prrty
            tablePriority.Item("id") = userIden()
            Cache.Insert(keytableName, tablePriority, Nothing, Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(3))
            Return ""
        End If

    End Function

    Function unlockKeyTable(ByVal prrty As priority) As String  'who is locking the table
        Dim tablePriority As New NameValueCollection
        If Cache(keytableName) Is Nothing Then
            Return ""
        End If

        tablePriority = CType(Cache(keytableName), NameValueCollection)
        If prrty < CInt(tablePriority.Item("p")) AndAlso priority.ganttChart <> CInt(tablePriority.Item("p")) Then
            Return tablePriority.Item("id")
        Else
            Cache.Remove(keytableName)
            Return ""
        End If

    End Function



    Private Sub Page_PreLoad(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreLoad
        checkAuthorityForPage()
    End Sub

    Protected Overridable Sub checkAuthorityForPage()
        'if not logged in, then go to login page
        Dim cookiePath As String = Request.ApplicationPath

        If accesslevel() > 2 Then
            If (Request.Cookies("userInfo") Is Nothing) OrElse (CType(Request.Cookies("userInfo")("level"), Integer) < accesslevel()) OrElse (Request.Cookies(cookiePath) Is Nothing) Then
                Response.Redirect("~/usermanage/login.aspx?accLvl=" & accesslevel & "&oldurl=" & IIf(String.IsNullOrEmpty(Request.Params("login")), Request.Url.ToString, String.Empty))
            End If
        End If
    End Sub

End Class



