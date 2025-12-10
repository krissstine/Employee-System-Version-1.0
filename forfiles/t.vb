Imports MySql.Data.MySqlClient

Public Class t
    Private connectionString As String = "server=localhost;userid=root;password=;database=psystem"


    Private currentUserID As Integer = 0

    Private Sub t_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.WindowState = FormWindowState.Maximized


        ' covers the entire screen

        ' Get userID from Form1 shared variable
        currentUserID = Form1.CurrentUserID

        ' Display userID-username in TextBox2
        TextBox2.Text = Form1.UserIDDisplay ' Will show "1-krot"

        ' Check if user is logged in
        If currentUserID <= 0 Then
            MessageBox.Show("User not logged in. Please log in again.", "Error")
            Form1.Show()
            Me.Close()
            Return
        End If

        ' Now load the data
        RefreshData()

        Dim today As Date = Date.Today
        ' Set valid birthdate range for DateTimePicker3
        DateTimePicker3.MinDate = today.AddYears(-80)  ' Oldest allowed: 80 years old
        DateTimePicker3.MaxDate = today.AddYears(-15)  ' Youngest allowed: 15 years old
        ' Optional: Set default value to 15-year-old
        DateTimePicker3.Value = DateTimePicker3.MaxDate
    End Sub
    Public Sub RefreshData()
        LoadComboBoxData(ComboBox2, "SELECT departmentID, departmentname FROM department WHERE userID = @userID ORDER BY departmentID ASC", "departmentname", "departmentID")
        LoadComboBoxData(ComboBox3, "SELECT positionID, positionname FROM position WHERE userID = @userID ORDER BY positionID ASC", "positionname", "positionID")
        LoadEmployeesGrid()
    End Sub

    ' Generic method to load comboboxes
    Private Sub LoadComboBoxData(combo As ComboBox, query As String, displayMember As String, valueMember As String)
        Try
            Using conn As New MySqlConnection(connectionString)
                Dim adapter As New MySqlDataAdapter(query, conn)
                adapter.SelectCommand.Parameters.AddWithValue("@userID", currentUserID)
                Dim table As New DataTable()
                adapter.Fill(table)
                combo.DataSource = table
                combo.DisplayMember = displayMember
                combo.ValueMember = valueMember
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading data: " & ex.Message)
        End Try
    End Sub

    ' Load employees into DataGridView with Department and Position Names
    Private Sub LoadEmployeesGrid()
        Try
            Using conn As New MySqlConnection(connectionString)
                Dim adapter As New MySqlDataAdapter("SELECT e.employeeID, e.Full_name, e.gender, e.Birthdate, e.email, " &
                                            "d.departmentname, p.positionname, e.rate, e.hireddate, e.status " &
                                            "FROM employee e " &
                                            "LEFT JOIN department d ON e.departmentID = d.departmentID " &
                                            "LEFT JOIN position p ON e.positionID = p.positionID " &
                                            "WHERE e.userID = @userID " &
                                            "ORDER BY e.employeeID ASC", conn)
                adapter.SelectCommand.Parameters.AddWithValue("@userID", currentUserID)
                Dim table As New DataTable()
                adapter.Fill(table)
                DataGridView1.DataSource = table

                ' Format Birthdate column to yyyy-MM-dd
                If DataGridView1.Columns.Contains("Birthdate") Then
                    DataGridView1.Columns("Birthdate").DefaultCellStyle.Format = "yyyy-MM-dd"
                End If

                ' Format hireddate column to yyyy-MM-dd
                If DataGridView1.Columns.Contains("hireddate") Then
                    DataGridView1.Columns("hireddate").DefaultCellStyle.Format = "yyyy-MM-dd"
                End If

                ' Format rate column as currency
                If DataGridView1.Columns.Contains("rate") Then
                    DataGridView1.Columns("rate").DefaultCellStyle.Format = "₱#,##0.00"
                End If

                ' Auto-fit columns for better visibility
                DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading employees: " & ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ' Trim all inputs once
            Dim fullName As String = TextBox1.Text.Trim()
            Dim email As String = TextBox4.Text.Trim().ToLowerInvariant()
            Dim gender As String = If(ComboBox4.SelectedIndex >= 0, ComboBox4.Text.Trim(), String.Empty)
            Dim statusText As String = If(ComboBox1.SelectedIndex >= 0, ComboBox1.Text.Trim(), String.Empty)
            Dim deptValue As Object = ComboBox2.SelectedValue
            Dim posValue As Object = ComboBox3.SelectedValue

            ' VALIDATION 1: Full Name
            If String.IsNullOrWhiteSpace(fullName) Then
                MessageBox.Show("Please enter Full Name", "Validation Error")
                Return
            End If

            If Not System.Text.RegularExpressions.Regex.IsMatch(fullName, "^\p{L}+(?:[ \p{L}'-]+\p{L})*$") Then
                MessageBox.Show("Full Name can only contain letters, spaces, hyphens, and apostrophes", "Validation Error")
                Return
            End If

            If fullName.Length < 3 Then
                MessageBox.Show("Full Name must be at least 3 characters long", "Validation Error")
                Return
            End If

            ' VALIDATION 2: Gender
            If String.IsNullOrEmpty(gender) Then
                MessageBox.Show("Please select Gender", "Validation Error")
                Return
            End If

            ' VALIDATION 3: Birthdate + Age bounds
            Dim birthDate As Date = DateTimePicker3.Value.Date
            Dim today As Date = Date.Today
            If birthDate > today Then
                MessageBox.Show("Birthdate cannot be in the future", "Validation Error")
                Return
            End If
            Dim age As Integer = today.Year - birthDate.Year
            If birthDate > today.AddYears(-age) Then age -= 1
            If age < 15 OrElse age > 75 Then
                MessageBox.Show("Age must be between 15 and 75 years", "Validation Error")
                Return
            End If

            ' VALIDATION 4: Email (simple robust pattern)
            If String.IsNullOrWhiteSpace(email) Then
                MessageBox.Show("Please enter Email", "Validation Error")
                Return
            End If
            If Not System.Text.RegularExpressions.Regex.IsMatch(email, "^[^@\s]+@[^@\s]+\.[^@\s]+$") Then
                MessageBox.Show("Please enter a valid email address", "Validation Error")
                Return
            End If

            ' VALIDATION 5: Daily Rate (decimal, positive)
            Dim rateValue As Decimal
            If Not Decimal.TryParse(TextBox6.Text.Trim(), rateValue) Then
                MessageBox.Show("Daily Rate must be a valid number", "Validation Error")
                Return
            End If
            If rateValue <= 0D Then
                MessageBox.Show("Daily Rate must be greater than 0", "Validation Error")
                Return
            End If

            ' VALIDATION 6: Status
            If String.IsNullOrEmpty(statusText) Then
                MessageBox.Show("Please select Status", "Validation Error")
                Return
            End If

            ' VALIDATION 7: Department
            If ComboBox2.SelectedIndex = -1 OrElse deptValue Is Nothing Then
                MessageBox.Show("Please select Department", "Validation Error")
                Return
            End If

            ' VALIDATION 8: Position
            If ComboBox3.SelectedIndex = -1 OrElse posValue Is Nothing Then
                MessageBox.Show("Please select Position", "Validation Error")
                Return
            End If

            ' Dates
            Dim hireDate As Date = DateTimePicker1.Value.Date
            If hireDate > today.AddDays(1) Then
                MessageBox.Show("Hire date cannot be in the future", "Validation Error")
                Return
            End If

            Using conn As New MySqlConnection(connectionString)
                conn.Open()
                Using tx = conn.BeginTransaction()
                    ' Check duplicate email FOR THIS USER ONLY
                    Using checkCmd As New MySqlCommand("SELECT COUNT(*) FROM employee WHERE email = @email AND userID = @userID", conn, tx)
                        checkCmd.Parameters.Add("@email", MySqlDbType.VarChar).Value = email
                        checkCmd.Parameters.Add("@userID", MySqlDbType.Int32).Value = currentUserID
                        Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                        If exists > 0 Then
                            MessageBox.Show("Email already exists in your records", "Validation Error")
                            tx.Rollback()
                            Return
                        End If
                    End Using

                    ' Insert with strong types AND userID
                    Dim query As String = "INSERT INTO employee " &
                "(Full_name, gender, Birthdate, email, rate, hireddate, status, departmentID, positionID, userID) " &
                "VALUES (@name, @gender, @birth, @email, @rate, @hire, @status, @dept, @pos, @userID)"
                    Using cmd As New MySqlCommand(query, conn, tx)
                        cmd.Parameters.Add("@name", MySqlDbType.VarChar).Value = fullName
                        cmd.Parameters.Add("@gender", MySqlDbType.VarChar).Value = gender
                        cmd.Parameters.Add("@birth", MySqlDbType.Date).Value = birthDate
                        cmd.Parameters.Add("@email", MySqlDbType.VarChar).Value = email
                        cmd.Parameters.Add("@rate", MySqlDbType.Decimal).Value = rateValue
                        cmd.Parameters.Add("@hire", MySqlDbType.Date).Value = hireDate
                        cmd.Parameters.Add("@status", MySqlDbType.VarChar).Value = statusText

                        ' Ensure IDs are the correct type (Int32 or the bound type)
                        Dim deptId As Integer = Convert.ToInt32(deptValue)
                        Dim posId As Integer = Convert.ToInt32(posValue)
                        cmd.Parameters.Add("@dept", MySqlDbType.Int32).Value = deptId
                        cmd.Parameters.Add("@pos", MySqlDbType.Int32).Value = posId
                        cmd.Parameters.Add("@userID", MySqlDbType.Int32).Value = currentUserID

                        cmd.ExecuteNonQuery()
                    End Using

                    tx.Commit()
                End Using
            End Using

            ' SUCCESS UI - APPEND INSTEAD OF CLEAR
            MessageBox.Show("Employee saved successfully!", "Success")

            ' Display employee data in RichTextBox1 - APPEND TO EXISTING TEXT
            Dim departmentName As String = ComboBox2.Text
            Dim positionName As String = ComboBox3.Text

            ' Add separator if there's already content
            If RichTextBox1.TextLength > 0 Then
                RichTextBox1.AppendText(Environment.NewLine & New String("="c, 70) & Environment.NewLine & Environment.NewLine)
            End If

            With RichTextBox1
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .SelectionColor = Color.Green
                .AppendText("✓ EMPLOYEE RECORD SAVED SUCCESSFULLY" & Environment.NewLine)
                .AppendText(New String("="c, 70) & Environment.NewLine & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .SelectionColor = Color.DarkBlue
                .AppendText("EMPLOYEE INFORMATION" & Environment.NewLine)
                .AppendText(New String("-"c, 70) & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Full Name: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(fullName & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Gender: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(gender & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Birthdate: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(birthDate.ToString("MMMM dd, yyyy") & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Age: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(age.ToString() & " years old" & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Email: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(email & Environment.NewLine & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .SelectionColor = Color.DarkBlue
                .AppendText("WORK INFORMATION" & Environment.NewLine)
                .AppendText(New String("-"c, 70) & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Hire Date: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(hireDate.ToString("MMMM dd, yyyy") & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Status: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(statusText & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Department: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(departmentName & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Position: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText(positionName & Environment.NewLine & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .SelectionColor = Color.DarkBlue
                .AppendText("COMPENSATION" & Environment.NewLine)
                .AppendText(New String("-"c, 70) & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Black
                .AppendText("Daily Rate: ")
                .SelectionColor = Color.DarkGreen
                .SelectionFont = New Font(.Font, FontStyle.Bold)
                .AppendText("₱" & rateValue.ToString("#,##0.00") & "/day" & Environment.NewLine & Environment.NewLine)

                .SelectionFont = New Font(.Font, FontStyle.Regular)
                .SelectionColor = Color.Gray
                .AppendText(New String("="c, 70) & Environment.NewLine)
                .SelectionFont = New Font(.Font, FontStyle.Italic)
                .AppendText("Saved On: " & DateTime.Now.ToString("MMMM dd, yyyy @ hh:mm:ss tt") & Environment.NewLine)

                ' SCROLL TO BOTTOM
                .ScrollToCaret()
                .Focus()
            End With

            RefreshData()

            ' Clear inputs (keep dates if you prefer)
            TextBox1.Clear()
            ComboBox4.SelectedIndex = -1
            TextBox4.Clear()
            TextBox6.Clear()
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox3.SelectedIndex = -1

        Catch ex As MySqlException
            MessageBox.Show("Database Error: " & ex.Message, "Error Saving Employee")
            ShowRichError(ex.Message)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error Saving Employee")
            ShowRichError(ex.Message)
        End Try
    End Sub

    Private Sub ShowRichError(errorMessage As String)
        RichTextBox1.Clear()
        With RichTextBox1
            .SelectionFont = New Font(.Font, FontStyle.Bold)
            .SelectionColor = Color.Red
            .AppendText("❌ ERROR SAVING EMPLOYEE" & Environment.NewLine)
            .AppendText(New String("="c, 70) & Environment.NewLine & Environment.NewLine)

            .SelectionFont = New Font(.Font, FontStyle.Regular)
            .SelectionColor = Color.DarkRed
            .AppendText("Error Details:" & Environment.NewLine)
            .SelectionFont = New Font(.Font, FontStyle.Bold)
            .AppendText(errorMessage & Environment.NewLine & Environment.NewLine)

            .SelectionFont = New Font(.Font, FontStyle.Regular)
            .SelectionColor = Color.Gray
            .AppendText(New String("="c, 70) & Environment.NewLine)
            .SelectionFont = New Font(.Font, FontStyle.Italic)
            .AppendText("Time: " & DateTime.Now.ToString("MMMM dd, yyyy @ hh:mm:ss tt"))

            .ScrollToCaret()
        End With
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' **Validate selection**
        If DataGridView1.SelectedRows.Count <> 1 Then
            MessageBox.Show("Please select exactly one row to delete.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

        ' **Validate column existence by name**
        If Not DataGridView1.Columns.Contains("employeeID") Then
            MessageBox.Show("The 'employeeID' column was not found in the grid.", "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' **Read and validate the cell value**
        Dim cell As DataGridViewCell = selectedRow.Cells("employeeID")
        Dim employeeID As String = If(cell.Value, "").ToString().Trim()

        If String.IsNullOrWhiteSpace(employeeID) Then
            MessageBox.Show("Employee ID is empty or invalid in the selected row.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' **Confirm deletion**
        Dim confirmResult As DialogResult = MessageBox.Show("Are you sure you want to delete this employee record?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirmResult <> DialogResult.Yes Then Exit Sub

        ' **Database deletion**
        Dim connString As String = "server=localhost;userid=root;password=;database=psystem"
        Using conn As New MySqlConnection(connString)
            Try
                conn.Open()
                Dim deleteQuery As String = "DELETE FROM employee WHERE employeeID = @employeeID"
                Using deleteCmd As New MySqlCommand(deleteQuery, conn)
                    deleteCmd.Parameters.AddWithValue("@employeeID", employeeID)

                    Dim rowsAffected As Integer = deleteCmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Employee record deleted successfully.", "Delete Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        RefreshData()
                    Else
                        MessageBox.Show("No records were deleted. Please verify the Employee ID.", "Delete Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End Using
            Catch ex As MySqlException
                MessageBox.Show("Database error occurred: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch ex As Exception
                MessageBox.Show("Unexpected error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub
    Private Sub ClearEmployeeInputs()
        TextBox1.Clear() ' Full name
        ComboBox4.SelectedIndex = -1 ' Gender
        DateTimePicker3.Value = DateTime.Now ' Birthdate
        TextBox4.Clear() ' Email
        TextBox6.Clear() ' Rate
        DateTimePicker1.Value = DateTime.Now ' Hire date
        ComboBox1.SelectedIndex = -1 ' Status
        ComboBox2.SelectedIndex = -1 ' Department
        ComboBox3.SelectedIndex = -1 ' Position
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Loop through all controls, including nested ones
        ClearControls(Me)

        ' Reset DateTimePicker3 to today's date with year 2010
        DateTimePicker3.Value = New DateTime(2010, DateTime.Today.Month, DateTime.Today.Day)

        ' Optionally, set focus back to the first input
        TextBox1.Focus()
    End Sub
    Private Sub ClearControls(parent As Control)
        For Each ctrl As Control In parent.Controls
            If TypeOf ctrl Is TextBox Then
                DirectCast(ctrl, TextBox).Clear()

            ElseIf TypeOf ctrl Is ComboBox Then
                DirectCast(ctrl, ComboBox).SelectedIndex = -1

            ElseIf TypeOf ctrl Is DateTimePicker Then
                Dim dtp As DateTimePicker = DirectCast(ctrl, DateTimePicker)
                ' Validate before setting to Today
                If DateTime.Today >= dtp.MinDate AndAlso DateTime.Today <= dtp.MaxDate Then
                    dtp.Value = DateTime.Today
                Else
                    dtp.Value = dtp.MinDate ' fallback to safe value
                End If
            End If

            ' Recursively clear nested controls (GroupBox, Panel, etc.)
            If ctrl.HasChildren Then
                ClearControls(ctrl)
            End If
        Next
    End Sub
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        deppos.Show()
        Me.Hide()
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        ' Check if DateTimePicker2 day is 15 or 30 (semi-monthly payroll)
        Dim currentDay As Integer = DateTimePicker2.Value.Day

        If currentDay = 15 OrElse currentDay = 30 Then
            prl.Show()
            Me.Hide()
        Else
            MessageBox.Show("Payroll can only be processed on the 15th or 30th of the month. Current date is: " & DateTimePicker2.Value.ToString("MMMM dd, yyyy"), "Invalid Date", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        Try
            ' Check if RichTextBox1 has content
            If String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                ' No data to save, show warning but continue to next form
                MessageBox.Show("No history data has been saved.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                ' Clear RichTextBox1 and show payroll form anyway
                RichTextBox1.Clear()
                payroll.Show()
                Me.Close()
                Return
            End If

            ' Save data to history table
            Dim connectionString As String = "server=localhost;userid=root;password=;database=psystem"

            Using conn As New MySqlConnection(connectionString)
                conn.Open()

                ' Extract userID from TextBox2 (e.g., "5-krot" -> "5")
                Dim userIDText As String = TextBox2.Text.Split("-"c)(0)
                Dim userID As Integer = Convert.ToInt32(userIDText)

                ' Get filling date from DateTimePicker2
                Dim fillingDate As String = DateTimePicker2.Value.ToString("yyyy-MM-dd")

                ' Get ALL employee details from RichTextBox1 (entire content)
                Dim employeeDetails As String = RichTextBox1.Text

                ' Insert into history table
                Dim query As String = "INSERT INTO history (fillingdate, userID, employeedet) VALUES (@fillingDate, @userID, @employeeDetails)"

                Using cmd As New MySqlCommand(query, conn)
                    ' Use MySqlDbType.LongText to handle large text
                    cmd.Parameters.Add("@fillingDate", MySqlDbType.Date).Value = DateTime.Parse(fillingDate)
                    cmd.Parameters.Add("@userID", MySqlDbType.Int32).Value = userID
                    cmd.Parameters.Add("@employeeDetails", MySqlDbType.LongText).Value = employeeDetails

                    cmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("History saved successfully! All records saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using

            ' Clear RichTextBox1 and show payroll form
            RichTextBox1.Clear()
            payroll.Show()
            Me.Close()

        Catch ex As MySqlException
            MessageBox.Show("Database Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        RichTextBox1.ReadOnly = True

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Ensure one row is selected
        If DataGridView1.SelectedRows.Count = 1 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim employeeID As Integer = 0
            Try
                employeeID = Integer.Parse(selectedRow.Cells("employeeID").Value.ToString())
                Dim con As New MySqlConnection(connectionString)
                Dim query As String = "SELECT employeeID, Full_name, gender, Birthdate, email, departmentID, positionID, rate, hireddate, status FROM employee WHERE employeeID = @employeeID"
                Dim cmd As New MySqlCommand(query, con)
                cmd.Parameters.AddWithValue("@employeeID", employeeID)
                con.Open()
                Dim reader As MySqlDataReader = cmd.ExecuteReader()
                Dim deptID As Integer = 0
                Dim posID As Integer = 0

                If reader.Read() Then
                    TextBox1.Text = reader("Full_name").ToString()
                    ComboBox4.Text = reader("gender").ToString()
                    TextBox4.Text = reader("email").ToString()
                    deptID = Convert.ToInt32(reader("departmentID"))
                    posID = Convert.ToInt32(reader("positionID"))
                    TextBox6.Text = reader("rate").ToString()

                    If Not IsDBNull(reader("Birthdate")) Then
                        DateTimePicker3.Value = Convert.ToDateTime(reader("Birthdate"))
                    End If

                    If Not IsDBNull(reader("hireddate")) Then
                        DateTimePicker1.Value = Convert.ToDateTime(reader("hireddate"))
                    End If

                    ComboBox1.Text = reader("status").ToString()
                Else
                    MessageBox.Show("Employee record not found.", "Lookup Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
                reader.Close()

                ' Lookup department name from departmentID
                Dim deptCmd As New MySqlCommand("SELECT departmentname FROM department WHERE departmentID = @deptID", con)
                deptCmd.Parameters.AddWithValue("@deptID", deptID)
                Dim deptName As String = deptCmd.ExecuteScalar()?.ToString()
                ComboBox2.SelectedValue = deptID

                ' Lookup position name from positionID
                Dim posCmd As New MySqlCommand("SELECT positionname FROM position WHERE positionID = @posID", con)
                posCmd.Parameters.AddWithValue("@posID", posID)
                Dim posName As String = posCmd.ExecuteScalar()?.ToString()
                ComboBox3.SelectedValue = posID

                con.Close()
            Catch ex As Exception
                MessageBox.Show("Error loading employee data: " & ex.Message, "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("Please select exactly one row to edit.", "Edit Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView1.SelectedCells.Count < 1 Then
            MessageBox.Show("Please select a row to update.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim empID As String = DataGridView1.SelectedCells(0).OwningRow.Cells("employeeID").Value?.ToString()

        If String.IsNullOrEmpty(empID) Then
            MessageBox.Show("Employee ID not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' VALIDATION 1: Full Name
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            MessageBox.Show("Please enter Full Name", "Validation Error")
            Return
        End If

        ' Validate Full Name - Check if it contains numbers or special characters
        If Not System.Text.RegularExpressions.Regex.IsMatch(TextBox1.Text, "^[a-zA-Z\s'-]+$") Then
            MessageBox.Show("Full Name can only contain letters, spaces, hyphens, and apostrophes", "Validation Error")
            Return
        End If

        ' Validate Full Name length
        If TextBox1.Text.Length < 3 Then
            MessageBox.Show("Full Name must be at least 3 characters long", "Validation Error")
            Return
        End If

        ' VALIDATION 2: Gender
        If ComboBox4.SelectedIndex = -1 Then
            MessageBox.Show("Please select Gender", "Validation Error")
            Return
        End If

        ' VALIDATION 3: Birthdate
        Dim birthDate As Date = DateTimePicker3.Value
        Dim today As Date = Date.Today
        Dim age As Integer = today.Year - birthDate.Year
        If birthDate > today.AddYears(-age) Then age -= 1

        ' VALIDATION 4: Email
        If String.IsNullOrWhiteSpace(TextBox4.Text) Then
            MessageBox.Show("Please enter Email", "Validation Error")
            Return
        End If

        ' Validate Email format
        If Not TextBox4.Text.Contains("@") Then
            MessageBox.Show("Email must contain '@' symbol", "Validation Error")
            Return
        End If

        If TextBox4.Text.IndexOf("@") = 0 Then
            MessageBox.Show("Email cannot start with '@' symbol", "Validation Error")
            Return
        End If

        If TextBox4.Text.LastIndexOf("@") <> TextBox4.Text.IndexOf("@") Then
            MessageBox.Show("Email cannot contain multiple '@' symbols", "Validation Error")
            Return
        End If

        If Not (TextBox4.Text.Contains(".com") Or TextBox4.Text.Contains(".org") Or TextBox4.Text.Contains(".net") Or TextBox4.Text.Contains(".co") Or TextBox4.Text.Contains(".ph")) Then
            MessageBox.Show("Email must have valid domain (.com, .org, .net, .co, .ph, etc.)", "Validation Error")
            Return
        End If

        ' VALIDATION 5: Daily Rate
        If String.IsNullOrWhiteSpace(TextBox6.Text) Then
            MessageBox.Show("Please enter Rate", "Validation Error")
            Return
        End If

        Dim rateValue As Double
        If Not Double.TryParse(TextBox6.Text, rateValue) Then
            MessageBox.Show("Daily Rate must be a valid number", "Validation Error")
            Return
        End If

        ' VALIDATION 6: Status
        If ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("Please select Status", "Validation Error")
            Return
        End If

        ' VALIDATION 7: Department
        If ComboBox2.SelectedIndex = -1 Then
            MessageBox.Show("Please select Department", "Validation Error")
            Return
        End If

        ' VALIDATION 8: Position
        If ComboBox3.SelectedIndex = -1 Then
            MessageBox.Show("Please select Position", "Validation Error")
            Return
        End If

        Using conn As New MySqlConnection(connectionString)
            Try
                conn.Open()

                Dim updateQuery As String = "UPDATE employee SET Full_name=@name, gender=@gender, Birthdate=@birth, email=@email, " &
                                        "rate=@rate, hireddate=@hire, status=@status, departmentID=@dept, positionID=@pos " &
                                        "WHERE employeeID=@empID"

                Using updateCmd As New MySqlCommand(updateQuery, conn)
                    Dim birthFormatted As String = DateTimePicker3.Value.ToString("yyyy-MM-dd")
                    Dim hireFormatted As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")

                    updateCmd.Parameters.AddWithValue("@empID", empID)
                    updateCmd.Parameters.AddWithValue("@name", TextBox1.Text.Trim())
                    updateCmd.Parameters.AddWithValue("@gender", ComboBox4.Text.Trim())
                    updateCmd.Parameters.AddWithValue("@birth", birthFormatted)
                    updateCmd.Parameters.AddWithValue("@email", TextBox4.Text.Trim())
                    updateCmd.Parameters.AddWithValue("@rate", rateValue)
                    updateCmd.Parameters.AddWithValue("@hire", hireFormatted)
                    updateCmd.Parameters.AddWithValue("@status", ComboBox1.SelectedItem?.ToString())
                    updateCmd.Parameters.AddWithValue("@dept", ComboBox2.SelectedValue)
                    updateCmd.Parameters.AddWithValue("@pos", ComboBox3.SelectedValue)

                    Dim rowsAffected As Integer = updateCmd.ExecuteNonQuery()

                    If rowsAffected > 0 Then
                        MessageBox.Show("Employee record updated successfully.", "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

                        ' Get Department Name from ComboBox2
                        Dim departmentName As String = ComboBox2.Text

                        ' Get Position Name from ComboBox3
                        Dim positionName As String = ComboBox3.Text

                        ' Display updated employee data in RichTextBox1
                        RichTextBox1.Clear()

                        With RichTextBox1
                            ' Header
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .SelectionColor = Color.Green
                            .AppendText("EMPLOYEE RECORD UPDATED SUCCESSFULLY" & Environment.NewLine)
                            .AppendText(New String("="c, 50) & Environment.NewLine & Environment.NewLine)

                            ' Employee details
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .SelectionColor = Color.DarkBlue
                            .AppendText("EMPLOYEE INFORMATION" & Environment.NewLine)
                            .AppendText(New String("-"c, 50) & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Full Name: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(TextBox1.Text & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Gender: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(ComboBox4.Text & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Birthdate: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(birthDate.ToString("MMMM dd, yyyy") & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Age: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(age.ToString() & " years old" & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Email: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(TextBox4.Text & Environment.NewLine & Environment.NewLine)

                            ' Work Information
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .SelectionColor = Color.DarkBlue
                            .AppendText("WORK INFORMATION" & Environment.NewLine)
                            .AppendText(New String("-"c, 50) & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Hire Date: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(DateTimePicker1.Value.ToString("MMMM dd, yyyy") & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Status: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(ComboBox1.SelectedItem?.ToString() & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Department: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(departmentName & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Position: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText(positionName & Environment.NewLine & Environment.NewLine)

                            ' Compensation
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .SelectionColor = Color.DarkBlue
                            .AppendText("COMPENSATION" & Environment.NewLine)
                            .AppendText(New String("-"c, 50) & Environment.NewLine)

                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Black
                            .AppendText("Daily Rate: ")
                            .SelectionColor = Color.DarkGreen
                            .SelectionFont = New Font(.Font, FontStyle.Bold)
                            .AppendText("₱" & Format(rateValue, "#,##0.00") & "/day" & Environment.NewLine & Environment.NewLine)

                            ' Footer with timestamp
                            .SelectionFont = New Font(.Font, FontStyle.Regular)
                            .SelectionColor = Color.Gray
                            .AppendText(New String("="c, 50) & Environment.NewLine)
                            .SelectionFont = New Font(.Font, FontStyle.Italic)
                            .AppendText("Updated On: " & DateTime.Now.ToString("MMMM dd, yyyy @ hh:mm:ss tt"))
                        End With

                        RefreshData()
                    Else
                        MessageBox.Show("No records were updated. Please check if the Employee ID exists.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End Using

            Catch ex As Exception
                MessageBox.Show("Error updating employee data: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                ' Display error in RichTextBox
                RichTextBox1.Clear()
                With RichTextBox1
                    .SelectionFont = New Font(.Font, FontStyle.Bold)
                    .SelectionColor = Color.Red
                    .AppendText("ERROR UPDATING EMPLOYEE" & Environment.NewLine)
                    .AppendText(New String("="c, 50) & Environment.NewLine & Environment.NewLine)
                    .SelectionFont = New Font(.Font, FontStyle.Regular)
                    .SelectionColor = Color.DarkRed
                    .AppendText("Error Details:" & Environment.NewLine)
                    .SelectionFont = New Font(.Font, FontStyle.Bold)
                    .AppendText(ex.Message & Environment.NewLine & Environment.NewLine)
                    .SelectionFont = New Font(.Font, FontStyle.Regular)
                    .SelectionColor = Color.Gray
                    .AppendText("Time: " & DateTime.Now.ToString("MMMM dd, yyyy @ hh:mm:ss tt"))
                End With
            End Try
        End Using
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            ' Check if RichTextBox1 has content
            If String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                ' No data to save, show warning but continue to next form
                MessageBox.Show("No history data has been saved.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                ' Clear RichTextBox1 and show Form1 anyway
                RichTextBox1.Clear()
                Form1.Show()
                Me.Hide()
                Return
            End If

            ' Save data to history table
            Dim connectionString As String = "server=localhost;userid=root;password=;database=psystem"

            Using conn As New MySqlConnection(connectionString)
                conn.Open()

                ' Extract userID from TextBox2 (e.g., "5-krot" -> "5")
                Dim userIDText As String = TextBox2.Text.Split("-"c)(0)
                Dim userID As Integer = Convert.ToInt32(userIDText)

                ' Get filling date from DateTimePicker2
                Dim fillingDate As String = DateTimePicker2.Value.ToString("yyyy-MM-dd")

                ' Get ALL employee details from RichTextBox1 (entire content)
                Dim employeeDetails As String = RichTextBox1.Text

                ' Insert into history table
                Dim query As String = "INSERT INTO history (fillingdate, userID, employeedet) VALUES (@fillingDate, @userID, @employeeDetails)"

                Using cmd As New MySqlCommand(query, conn)
                    ' Use MySqlDbType.LongText to handle large text
                    cmd.Parameters.Add("@fillingDate", MySqlDbType.Date).Value = DateTime.Parse(fillingDate)
                    cmd.Parameters.Add("@userID", MySqlDbType.Int32).Value = userID
                    cmd.Parameters.Add("@employeeDetails", MySqlDbType.LongText).Value = employeeDetails

                    cmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("History saved successfully! All records saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using

            ' Clear RichTextBox1 and show Form1
            RichTextBox1.Clear()
            Form1.Show()
            Me.Hide()

        Catch ex As MySqlException
            MessageBox.Show("Database Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Dim selectedDate As DateTime = DateTimePicker2.Value
        Dim dayOfMonth As Integer = selectedDate.Day

        ' Check if it's the 15th
        If dayOfMonth = 15 Then
            RichTextBox1.BackColor = Color.LightYellow
            RichTextBox1.Text = "⚠️ NOTIFICATION - SEMI-MONTHLY PAYMENT" & vbCrLf & vbCrLf &
                            "Date: " & selectedDate.ToString("MMMM dd, yyyy") & vbCrLf &
                            "Status: Ready for semi-monthly payment for employee" & vbCrLf & vbCrLf &
                            "Action Required: Process payroll"
            ' Check if it's the 30th or 31st (end of month payment)
        ElseIf dayOfMonth = 30 OrElse dayOfMonth = 31 Then
            RichTextBox1.BackColor = Color.LightYellow
            RichTextBox1.Text = "⚠️ NOTIFICATION - SEMI-MONTHLY PAYMENT" & vbCrLf & vbCrLf &
                            "Date: " & selectedDate.ToString("MMMM dd, yyyy") & vbCrLf &
                            "Status: Ready for semi-monthly payment for employee" & vbCrLf & vbCrLf &
                            "Action Required: Process payroll"
        Else
            RichTextBox1.BackColor = Color.White
            RichTextBox1.Text = "No payment due on " & selectedDate.ToString("MMMM dd, yyyy")
        End If
    End Sub



    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub
    ' Allow only digits and limit to 10 characters
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        ' Block non-numeric input (except Backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
            Return
        End If

        ' Block input if already 10 digits
        If Char.IsDigit(e.KeyChar) AndAlso TextBox6.Text.Length >= 10 Then
            e.Handled = True
            MessageBox.Show("You can only enter up to 10 digits.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        Dim searchText As String = TextBox7.Text.Trim().ToLower()

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

                ' Scroll to first match
                If Not found Then
                    DataGridView1.FirstDisplayedScrollingRowIndex = row.Index
                    DataGridView1.CurrentCell = row.Cells(0)
                End If
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

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class