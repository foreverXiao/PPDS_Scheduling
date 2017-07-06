Imports basepage1
Imports System.Data
Imports System.Data.OleDb
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
        connstr = ConfigurationManager.ConnectionStrings(dbConnForParam).ProviderName & db

        Dim conn As OleDbConnection = New OleDbConnection(connstr)

        conn.Open()

        Dim i As Integer = 0

        Dim cmd0 As OleDbCommand = New OleDbCommand()
        Dim cmd As OleDbCommand = New OleDbCommand()

        Try
            cmd0.Connection = conn
            cmd0.CommandText = "SELECT * FROM [Esch_Na_tbl_userrole] WHERE [user_name] = '" & tbUsNm.Text & "'"


            cmd.Connection = conn
            cmd.CommandText = "INSERT INTO [Esch_Na_tbl_userrole]  ([user_name],[user_description],[password],[rightlevel]) VALUES('" & tbUsNm.Text & "','" & tbUsDscrptn.Text & "','" & tbNewpswrdAgn.Text & "'," & roleDDL1.SelectedValue & ")"




            If Not cmd0.ExecuteReader().Read() Then
                i = cmd.ExecuteNonQuery()
            Else
                msgPopUP("User's name: " & tbUsNm.Text & " already existed in the system.", lbStatus, True, False)

            End If

        Catch ex As Exception

        End Try

        cmd.Dispose()
        cmd0.Dispose()



        conn.Dispose()


        If i = 1 Then
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
