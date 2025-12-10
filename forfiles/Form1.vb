Imports MySql.Data.MySqlClient

Public Class Form1
    ' MySQL connection string
    Private connectionString As String = "server=localhost;userid=root;password=;database=psystem"

    ' Shared variables to hold the logged-in user's information
    Public Shared CurrentUser As String = ""
    Public Shared CurrentUserID As Integer = 0
    Public Shared UserIDDisplay As String = "" ' Format: "1-krot"

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ' Clear textboxes before opening registration form
        TextBox1.Clear()
        TextBox2.Clear()

        ' Open registration form
        y.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Basic validation
        If String.IsNullOrWhiteSpace(TextBox1.Text) OrElse String.IsNullOrWhiteSpace(TextBox2.Text) Then
            MessageBox.Show("Please enter both username and password.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Try
            Using conn As New MySqlConnection(connectionString)
                ' Updated query to get userID along with authentication
                Dim query As String = "SELECT userID FROM user WHERE username = @username AND password = @password LIMIT 1"
                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@username", TextBox1.Text)
                    cmd.Parameters.AddWithValue("@password", TextBox2.Text)
                    conn.Open()

                    Dim result As Object = cmd.ExecuteScalar()

                    If result IsNot Nothing Then
                        ' Login successful - store user information
                        Dim userID As Integer = Convert.ToInt32(result)
                        Dim username As String = TextBox1.Text

                        ' Set shared variables
                        CurrentUserID = userID
                        CurrentUser = username
                        UserIDDisplay = userID & "-" & username ' Format: "1-krot"

                        MessageBox.Show("Login successful!", "Welcome", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' Clear textboxes after successful login
                        TextBox1.Clear()
                        TextBox2.Clear()

                        ' Create new instance and pass the data
                        Dim employeeForm As New t()
                        employeeForm.Show()
                        Me.Hide()
                    Else
                        MessageBox.Show("Invalid username or password.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error during login: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub
End Class