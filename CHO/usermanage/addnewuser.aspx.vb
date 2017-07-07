Imports basepage1
Imports System.Data
Imports System.Data.SqlClient
Partial Class usermanage_addnewuser
    Inherits basepage1




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then
            Dim listdt As DataTable = New DataTable()
            ' Define the columns of the table.
            listdt.Columns.Add(New DataColumn("name", GetType(String)))
            listdt.Columns.Add(New DataColumn("value", GetType(Integer)))

            Dim dr As DataRow
            dr = listdt.NewRow()
            dr("name") = "Anyone"
            dr("value") = 2
            listdt.Rows.Add(dr)

            dr = listdt.NewRow()
            dr("name") = "schedule checker"
            dr("value") = 4
            listdt.Rows.Add(dr)


            dr = listdt.NewRow()
            dr("name") = "PSI or OF"
            dr("value") = 6
            listdt.Rows.Add(dr)


            dr = listdt.NewRow()
            dr("name") = "Planner"
            dr("value") = 8
            listdt.Rows.Add(dr)

            Dim dv As DataView = New DataView(listdt)

            roleDDL1.DataSource = dv
            roleDDL1.DataTextField = "name"
            roleDDL1.DataValueField = "value"
            roleDDL1.DataBind()

        End If

    End Sub

    Protected Sub btSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btSubmit.Click

        Dim db As String, connstr As String
        db = ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString
        connstr = db

        Dim conn As SqlConnection = New SqlConnection(connstr)

        conn.Open()

        Dim stReturn As String = String.Empty

        Dim cmd As SqlCommand = New SqlCommand()

        Try



            cmd.Connection = conn


            cmd.Parameters.Add("@LoginName", SqlDbType.NVarChar).Value = tbUsNm.Text.Trim().ToLower()
            cmd.Parameters.Add("@Password", SqlDbType.NVarChar).Value = tbNewpswrdAgn.Text.Trim()
            cmd.Parameters.Add("@description", SqlDbType.NVarChar).Value = tbUsDscrptn.Text.Trim()
            cmd.Parameters.Add("@authorityLevel", SqlDbType.Int).Value = Convert.ToInt32(roleDDL1.SelectedValue)
            'cmd.CommandText = "SELECT COUNT(*) FROM B_UserInfo WHERE LoginName = @LoginName AND Password = @Password AND Approved = 'True'"
            cmd.CommandText = "INSERT INTO [Esch_CQ_tbl_userrole]  ([user_name],[user_description],[password],[rightlevel]) VALUES(@LoginName, @description, @Password, @authorityLevel)"
            stReturn = cmd.ExecuteScalar()
            cmd.CommandText = "SELECT [user_name] FROM [Esch_CQ_tbl_userrole]  WHERE [user_name] = @LoginName "
            stReturn = cmd.ExecuteScalar()

        Catch ex As Exception

        End Try

        cmd.Dispose()

        conn.Dispose()


        If String.Compare(stReturn, tbUsNm.Text.Trim().ToLower()) Then
            msgPopUP("User's name: " & tbUsNm.Text & " already existed in the system.", lbStatus, True, False)
        Else
            msgPopUP("One new user " & tbUsNm.Text & " has been added in the system. ", lbStatus, False, False)
        End If

    End Sub

    Protected Sub ServerValidation(ByVal source As Object, ByVal args As ServerValidateEventArgs)
        args.IsValid = False

        Try


            If Len(tbUsNm.Text) <= 20 Then
                If Len(tbNewpswrd.Text) <= 20 AndAlso Len(tbNewpswrdAgn.Text) <= 20 Then
                    If Len(tbUsDscrptn) < 40 Then
                        args.IsValid = True
                    Else
                        msgPopUP("The number of characters in user's description  exceeds 20 .", lbStatus, True, False)
                    End If
                Else
                    msgPopUP("The number of characters in new password or confirmed password  exceeds 20 .", lbStatus, True, False)
                End If
            Else
                msgPopUP("The number of characters in user name  exceeds 20 .", lbStatus, True, False)
            End If

        Catch ex As Exception

            args.IsValid = False

        End Try


    End Sub

 
End Class
