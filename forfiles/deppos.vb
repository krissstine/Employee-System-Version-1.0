Imports MySql.Data.MySqlClient

Public Class deppos
    ' Reusable connection function
    Public Function GetConnection() As MySqlConnection
        Return New MySqlConnection("server=localhost;userid=root;password=;database=psystem")
    End Function

    ' Save department and position from TextBoxes
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Using conn As MySqlConnection = GetConnection()
                conn.Open()

                ' Get current user ID
                Dim currentUserID As Integer = Form1.CurrentUserID

                ' Insert department name from TextBox1 with userID
                Dim insertDeptQuery As String = "INSERT INTO department (departmentname, userID) VALUES (@deptName, @userID)"
                Using deptCmd As New MySqlCommand(insertDeptQuery, conn)
                    deptCmd.Parameters.AddWithValue("@deptName", TextBox1.Text)
                    deptCmd.Parameters.AddWithValue("@userID", currentUserID)
                    deptCmd.ExecuteNonQuery()
                End Using

                ' Insert position name from TextBox2 with userID
                Dim insertPosQuery As String = "INSERT INTO position (positionname, userID) VALUES (@posName, @userID)"
                Using posCmd As New MySqlCommand(insertPosQuery, conn)
                    posCmd.Parameters.AddWithValue("@posName", TextBox2.Text)
                    posCmd.Parameters.AddWithValue("@userID", currentUserID)
                    posCmd.ExecuteNonQuery()
                End Using
            End Using

            MessageBox.Show("Department and position saved successfully!", "Success")

            ' Open t form, refresh its data, and select latest entries
            Dim tForm As New t()
            tForm.Show()
            tForm.RefreshData()
            ' Select the most recently added department and position
            tForm.ComboBox2.SelectedIndex = tForm.ComboBox2.Items.Count - 1
            tForm.ComboBox3.SelectedIndex = tForm.ComboBox3.Items.Count - 1
            Me.Hide()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim tForm As New t()
        tForm.Show()
        Me.Close()
    End Sub

    Private Sub deppos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Clear previous data
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()

        ' Display userID-username in TextBox3 when form loads
        TextBox3.Text = Form1.UserIDDisplay ' Format: "5-krot"
    End Sub
End Class