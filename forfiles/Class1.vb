Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient

Public Class Class5


    Public Class DbConn
        Public Function GetConnection() As MySqlConnection
            Return New MySqlConnection("server=localhost;userid=root;password=;database=psystem")
        End Function
    End Class
End Class
