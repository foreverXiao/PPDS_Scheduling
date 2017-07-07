Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Linq



Partial Class dragDrop_ganttColor
    Inherits basepage1

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Response.Clear()

        Dim cssSTR As StringBuilder = New StringBuilder()

        Using connCSS As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings(dbConnForParam).ConnectionString)

            connCSS.Open()

            Dim command As New SqlCommand("SELECT * FROM [Esch_Na_tbl_colorCode] ", connCSS)
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

        Response.ContentType = "text/css"
        Response.Write(cssSTR.ToString)
        Response.End()


    End Sub

End Class


