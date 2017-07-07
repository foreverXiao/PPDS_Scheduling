Imports System.Data.SqlClient

Partial Class checkschedule_login
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Is last time login effective?
        Dim effectiveLogin As Boolean = False

        Dim cookiePath As String = Request.ApplicationPath

        'whether IIS application user login is valid
        If (Request.Cookies("userInfo") IsNot Nothing) AndAlso Request.Cookies(cookiePath) IsNot Nothing Then

            effectiveLogin = True
        End If



        'effetive login, then go to another web page
        If effectiveLogin Then
            If Not String.IsNullOrEmpty(Request.Params("oldurl")) Then
                Response.Redirect(Request.Params("oldurl") & "?login=1")
            End If
        End If


    End Sub

    Protected Sub LoginButton_Click(ByVal sender As Object, ByVal e As System.EventArgs)
  




        Dim connstr As String, nextpage As String = String.Empty

        If username1.Text <> "" AndAlso password1.Text <> "" Then

            connstr = ConfigurationManager.ConnectionStrings("accessDB").ConnectionString

            Dim conn As SqlConnection = New SqlConnection(connstr)
            conn.Open()
            Dim command As New SqlCommand("Select user_name,password,rightlevel From  Esch_Na_tbl_userrole ", conn)
            Dim reader As SqlDataReader


            Try
                reader = command.ExecuteReader()
                While reader.Read()
                    If reader("user_name").ToString.Equals(username1.Text, StringComparison.OrdinalIgnoreCase) AndAlso reader("password").ToString.Equals(password1.Text, StringComparison.OrdinalIgnoreCase) Then
                        Response.Cookies("userInfo")("id") = CStr(reader("user_name"))
                        Response.Cookies("userInfo")("level") = CStr(reader("rightlevel"))

                        Dim cookiePath As String = Request.ApplicationPath
                        Response.Cookies(cookiePath).Value = cookiePath

                        nextpage = Request.Params("oldurl")
                        If String.IsNullOrEmpty(nextpage) Then
                            nextpage = "checker.aspx"
                        End If

                        Exit While
                    End If
                End While

                reader.Close()

            Catch

            End Try





            command.Dispose()
            conn.Close()
            conn.Dispose()



            If Not String.IsNullOrEmpty(nextpage) Then
                Response.Redirect(nextpage)
            Else
                st.Text = "<div style='color:red;'>" & "Failed to log in, please check your user name and password." & "</div>"
            End If

        End If


    End Sub

  
    Protected Sub logout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles logout.Click
        Response.Cookies("userInfo").Expires = DateTime.Now.AddDays(-1)
        st.Text = ""
        username1.Focus()

    End Sub
End Class

