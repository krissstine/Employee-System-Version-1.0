Imports MySql.Data.MySqlClient

Public Class y
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Basic validation
        If String.IsNullOrWhiteSpace(TextBox2.Text) OrElse String.IsNullOrWhiteSpace(TextBox3.Text) Then
            MessageBox.Show("Please enter both username and password.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Try
            ' Connect to the psystem database and insert into the user table
            Using conn As New MySqlConnection("server=localhost;userid=root;password=;database=psystem")
                Dim query As String = "INSERT INTO user (username, password) VALUES (@username, @password)"
                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@username", TextBox2.Text)
                    cmd.Parameters.AddWithValue("@password", TextBox3.Text)

                    conn.Open()
                    cmd.ExecuteNonQuery()
                    conn.Close()
                End Using
            End Using

            MessageBox.Show("Account created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Hide()
            Form1.Show()

        Catch ex As Exception
            MessageBox.Show("Error saving account: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub y_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Optional: initialize form
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Form1.Show()
        Me.Hide()
    End Sub
End Class