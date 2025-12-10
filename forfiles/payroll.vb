Imports MySql.Data.MySqlClient

Public Class payroll
    ' Add this connection string at the class level
    Private connectionString As String = "server=localhost;userid=root;password=;database=psystem"
    Private currentUserID As Integer = 0

    ' Reusable connection function
    Public Function GetConnection() As MySqlConnection
        Return New MySqlConnection(connectionString)
    End Function

    Private Sub LoadComboBoxData(combo As ComboBox, query As String, displayMember As String, valueMember As String)
        Try
            Using conn As MySqlConnection = GetConnection()
                Dim adapter As New MySqlDataAdapter(query, conn)
                adapter.SelectCommand.Parameters.AddWithValue("@userID", currentUserID)
                Dim table As New DataTable()
                adapter.Fill(table)
                combo.DataSource = table
                combo.DisplayMember = displayMember
                combo.ValueMember = valueMember
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading combo box data: " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim employeeForm As New t()
        employeeForm.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ' Get the selected hire date from DateTimePicker1
            Dim selectedHireDate As Date = DateTimePicker1.Value.Date

            ' SQL query to fetch employees hired on the selected date with department and position names AND userID filter
            Dim query As String = "SELECT e.employeeID, e.Full_name, e.gender, e.Birthdate, e.email, " &
                              "d.departmentname, p.positionname, e.rate, e.hireddate, e.status " &
                              "FROM employee e " &
                              "LEFT JOIN department d ON e.departmentID = d.departmentID " &
                              "LEFT JOIN position p ON e.positionID = p.positionID " &
                              "WHERE DATE(e.hireddate) = @selectedDate AND e.userID = @userID " &
                              "ORDER BY e.Full_name ASC"

            Using conn As New MySqlConnection(connectionString)
                Dim adapter As New MySqlDataAdapter(query, conn)

                ' Add parameters
                adapter.SelectCommand.Parameters.AddWithValue("@selectedDate", selectedHireDate)
                adapter.SelectCommand.Parameters.AddWithValue("@userID", currentUserID)

                Dim table As New DataTable()
                adapter.Fill(table)

                ' Check if any records were found
                If table.Rows.Count > 0 Then
                    ' Clear previous data before displaying new data
                    DataGridView1.DataSource = Nothing

                    ' Display results in DataGridView
                    DataGridView1.DataSource = table

                    ' Format Birthdate column
                    If DataGridView1.Columns.Contains("Birthdate") Then
                        DataGridView1.Columns("Birthdate").DefaultCellStyle.Format = "yyyy-MM-dd"
                    End If

                    ' Format hireddate column
                    If DataGridView1.Columns.Contains("hireddate") Then
                        DataGridView1.Columns("hireddate").DefaultCellStyle.Format = "yyyy-MM-dd"
                    End If

                    ' Optional: Auto-fit columns for better visibility
                    DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

                    ' Display success message with count
                    MessageBox.Show("Found " & table.Rows.Count & " employee(s) hired on " &
                               selectedHireDate.ToString("MMMM dd, yyyy"),
                               "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ' No records found
                    DataGridView1.DataSource = Nothing
                    MessageBox.Show("No employees found hired on " & selectedHireDate.ToString("MMMM dd, yyyy"),
                               "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As MySqlException
            MessageBox.Show("Database Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            ' Get the selected payroll date from DateTimePicker1
            Dim selectedPayrollDate As Date = DateTimePicker1.Value.Date

            ' SQL query to fetch payroll records for the selected date AND userID filter
            Dim query As String = "SELECT PayrollDate, EmployeeID, EmployeeName, " &
                              "DailyRate, DaysWorked, GrossPay, Benefits, Deduction, NetPay, CreatedAt " &
                              "FROM payrollrecords " &
                              "WHERE DATE(PayrollDate) = @selectedDate AND userID = @userID " &
                              "ORDER BY EmployeeName ASC"

            Using conn As MySqlConnection = GetConnection()
                Dim adapter As New MySqlDataAdapter(query, conn)

                ' Add parameters
                adapter.SelectCommand.Parameters.AddWithValue("@selectedDate", selectedPayrollDate)
                adapter.SelectCommand.Parameters.AddWithValue("@userID", currentUserID)

                Dim table As New DataTable()
                adapter.Fill(table)

                ' Check if any records were found
                If table.Rows.Count > 0 Then
                    ' Clear previous data before displaying new data
                    DataGridView1.DataSource = Nothing

                    ' Display results in DataGridView1
                    DataGridView1.DataSource = table

                    ' Format currency columns in DataGridView1
                    If DataGridView1.Columns.Contains("DailyRate") Then
                        DataGridView1.Columns("DailyRate").DefaultCellStyle.Format = "₱#,##0.00"
                    End If
                    If DataGridView1.Columns.Contains("GrossPay") Then
                        DataGridView1.Columns("GrossPay").DefaultCellStyle.Format = "₱#,##0.00"
                    End If
                    If DataGridView1.Columns.Contains("Benefits") Then
                        DataGridView1.Columns("Benefits").DefaultCellStyle.Format = "₱#,##0.00"
                    End If
                    If DataGridView1.Columns.Contains("Deduction") Then
                        DataGridView1.Columns("Deduction").DefaultCellStyle.Format = "₱#,##0.00"
                    End If
                    If DataGridView1.Columns.Contains("NetPay") Then
                        DataGridView1.Columns("NetPay").DefaultCellStyle.Format = "₱#,##0.00"
                    End If
                    If DataGridView1.Columns.Contains("PayrollDate") Then
                        DataGridView1.Columns("PayrollDate").DefaultCellStyle.Format = "yyyy-MM-dd"
                    End If
                    If DataGridView1.Columns.Contains("CreatedAt") Then
                        DataGridView1.Columns("CreatedAt").DefaultCellStyle.Format = "yyyy-MM-dd"
                    End If

                    ' Auto-fit columns for better visibility
                    DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

                    ' Display success message with count
                    MessageBox.Show("Found " & table.Rows.Count & " payroll record(s) for " &
                                   selectedPayrollDate.ToString("MMMM dd, yyyy"),
                                   "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ' No records found - clear the grid
                    DataGridView1.DataSource = Nothing
                    MessageBox.Show("No payroll records found for " & selectedPayrollDate.ToString("MMMM dd, yyyy"),
                                   "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using

        Catch ex As MySqlException
            MessageBox.Show("Database Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Using conn As MySqlConnection = GetConnection()
                conn.Open()

                ' Validate selections
                If ComboBox1.SelectedIndex = -1 OrElse ComboBox2.SelectedIndex = -1 Then
                    MessageBox.Show("Please select both Department and Position.")
                    Exit Sub
                End If

                Dim selectedDeptID As Integer = Convert.ToInt32(ComboBox1.SelectedValue)
                Dim selectedPosID As Integer = Convert.ToInt32(ComboBox2.SelectedValue)

                ' Clear DataGridView before generating new data
                DataGridView1.Columns.Clear()
                DataGridView1.Rows.Clear()
                DataGridView1.AllowUserToAddRows = False

                ' Query filtered by department, position, and logged-in user
                Dim query As String = "SELECT employeeID, Full_name, rate, hireddate " &
                                  "FROM employee " &
                                  "WHERE departmentID = @deptID AND positionID = @posID AND userID = @userID " &
                                  "ORDER BY Full_name ASC"

                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@deptID", selectedDeptID)
                    cmd.Parameters.AddWithValue("@posID", selectedPosID)
                    cmd.Parameters.AddWithValue("@userID", currentUserID) ' logged-in user filter

                    Using reader As MySqlDataReader = cmd.ExecuteReader()
                        ' Add columns once
                        DataGridView1.Columns.Add("EmployeeID", "Employee ID")
                        DataGridView1.Columns.Add("FullName", "Full Name")
                        DataGridView1.Columns.Add("Rate", "Rate")
                        DataGridView1.Columns.Add("HiredDate", "Hired Date")

                        If reader.HasRows Then
                            While reader.Read()
                                Dim fullName As String = reader("Full_name").ToString().Trim()
                                Dim rateValue As Double

                                If Not String.IsNullOrWhiteSpace(fullName) AndAlso
                               Double.TryParse(reader("rate").ToString(), rateValue) AndAlso
                               rateValue > 0 Then

                                    Dim rowIndex As Integer = DataGridView1.Rows.Add()
                                    DataGridView1.Rows(rowIndex).Cells("EmployeeID").Value = reader("employeeID").ToString()
                                    DataGridView1.Rows(rowIndex).Cells("FullName").Value = fullName
                                    DataGridView1.Rows(rowIndex).Cells("Rate").Value = "₱" & Format(rateValue, "N2")
                                    DataGridView1.Rows(rowIndex).Cells("HiredDate").Value = CDate(reader("hireddate")).ToString("yyyy-MM-dd")
                                End If
                            End While

                            DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                            MessageBox.Show("Found " & DataGridView1.Rows.Count & " employee(s) in " &
                                        ComboBox1.Text & " - " & ComboBox2.Text,
                                        "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            MessageBox.Show("No employees found for the selected department and position.")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error retrieving employee data: " & ex.Message)
        End Try
    End Sub
    ' Add these methods to load ComboBoxes with user-filtered data
    Private Sub LoadDepartments()
        Try
            Using conn As MySqlConnection = GetConnection()
                conn.Open()
                ' Load only departments where the current user is assigned
                Dim query As String = "SELECT DISTINCT d.departmentID, d.departmentName FROM department d " &
                                  "INNER JOIN employee e ON d.departmentID = e.departmentID " &
                                  "WHERE e.userID = @userID ORDER BY d.departmentName ASC"

                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@userID", currentUserID)
                    Using reader As MySqlDataReader = cmd.ExecuteReader()
                        ComboBox1.DataSource = Nothing
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        ComboBox1.DataSource = dt
                        ComboBox1.DisplayMember = "departmentName"
                        ComboBox1.ValueMember = "departmentID"
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading departments: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadPositions()
        Try
            Using conn As MySqlConnection = GetConnection()
                conn.Open()
                ' Load only positions where the current user has employees
                Dim query As String = "SELECT DISTINCT p.positionID, p.positionName FROM position p " &
                                  "INNER JOIN employee e ON p.positionID = e.positionID " &
                                  "WHERE e.userID = @userID ORDER BY p.positionName ASC"

                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@userID", currentUserID)
                    Using reader As MySqlDataReader = cmd.ExecuteReader()
                        ComboBox2.DataSource = Nothing
                        Dim dt As New DataTable()
                        dt.Load(reader)
                        ComboBox2.DataSource = dt
                        ComboBox2.DisplayMember = "positionName"
                        ComboBox2.ValueMember = "positionID"
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading positions: " & ex.Message)
        End Try
    End Sub

    Private Sub payroll_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized


        ' Get userID from Form1
        currentUserID = Form1.CurrentUserID

        If currentUserID <= 0 Then
            MessageBox.Show("User not logged in. Please log in again.", "Error")
            Me.Close()
            Return
        End If

        ' Clear all previous data completely
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        ComboBox1.DataSource = Nothing
        ComboBox1.Items.Clear()

        ComboBox2.DataSource = Nothing
        ComboBox2.Items.Clear()

        ' Load Department ComboBox from department table with userID filter
        LoadComboBoxData(ComboBox1, "SELECT departmentID, departmentname FROM department WHERE userID = @userID ORDER BY departmentID ASC", "departmentname", "departmentID")

        ' Load Position ComboBox from position table with userID filter
        LoadComboBoxData(ComboBox2, "SELECT positionID, positionname FROM position WHERE userID = @userID ORDER BY positionID ASC", "positionname", "positionID")
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim searchText As String = TextBox1.Text.Trim().ToLower()

        ' Clear previous highlighting
        For Each row As DataGridViewRow In DataGridView1.Rows
            row.DefaultCellStyle.BackColor = Color.White
        Next

        ' If search text is empty, do nothing
        If String.IsNullOrWhiteSpace(searchText) Then
            Return
        End If

        ' Search and highlight entire row
        Dim found As Boolean = False
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim rowMatch As Boolean = False

            For Each cell As DataGridViewCell In row.Cells
                If cell.Value IsNot Nothing Then
                    Dim cellText As String = cell.Value.ToString().ToLower()

                    ' If match found in this row
                    If cellText.Contains(searchText) Then
                        rowMatch = True
                        Exit For
                    End If
                End If
            Next

            ' Highlight entire row if match found
            If rowMatch Then
                row.DefaultCellStyle.BackColor = Color.Plum ' Purple highlight
                found = True
            End If
        Next

        ' Scroll to first matching row
        If found Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.DefaultCellStyle.BackColor = Color.Plum Then
                    DataGridView1.FirstDisplayedScrollingRowIndex = row.Index
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            ' Get the selected date from DateTimePicker2
            Dim selectedFillingDate As Date = DateTimePicker2.Value.Date
            Dim fillingDateFormatted As String = selectedFillingDate.ToString("yyyy-MM-dd")

            ' SQL query to fetch ALL employeedet from history table based on fillingdate and userID
            Dim query As String = "SELECT employeedet FROM history WHERE DATE(fillingdate) = @fillingDate AND userID = @userID"

            Using conn As New MySqlConnection(connectionString)
                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@fillingDate", fillingDateFormatted)
                    cmd.Parameters.AddWithValue("@userID", currentUserID)
                    conn.Open()

                    Using reader As MySqlDataReader = cmd.ExecuteReader()
                        If reader.HasRows Then
                            ' Clear RichTextBox1 first
                            RichTextBox1.Clear()

                            ' Loop through all records and append to RichTextBox1
                            While reader.Read()
                                Dim employeeDetails As String = reader("employeedet").ToString()

                                ' Append each record with a separator
                                If RichTextBox1.TextLength > 0 Then
                                    RichTextBox1.AppendText(Environment.NewLine & New String("="c, 70) & Environment.NewLine & Environment.NewLine)
                                End If

                                RichTextBox1.AppendText(employeeDetails)
                            End While

                            MessageBox.Show("All employee details loaded successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            RichTextBox1.Clear()
                            MessageBox.Show("No employee details found for " & selectedFillingDate.ToString("MMMM dd, yyyy"), "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End Using
                End Using
            End Using

        Catch ex As MySqlException
            MessageBox.Show("Database Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        RichTextBox1.ReadOnly = True

    End Sub
End Class