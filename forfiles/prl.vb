Imports MySql.Data.MySqlClient

Public Class prl

    Private currentUserID As Integer = 0
    Private Const MINIMUM_DAYS_EMPLOYED As Integer = 14
    Private Const MAXIMUM_DAYS_WORKED As Integer = 15

    ' Reusable connection function
    Public Function GetConnection() As MySqlConnection
        Return New MySqlConnection("server=localhost;userid=root;password=;database=psystem")
    End Function

    Private Sub prl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Try
            ' Get userID from Form1
            currentUserID = Form1.CurrentUserID

            If currentUserID <= 0 Then
                MessageBox.Show("User not logged in. Please log in again.", "Error")
                Me.Close()
                Return
            End If

            ' Display userID-username in TextBox3
            If Not String.IsNullOrWhiteSpace(Form1.UserIDDisplay) Then
                TextBox3.Text = Form1.UserIDDisplay
                TextBox3.ReadOnly = True
            Else
                MessageBox.Show("User information not available.", "Error")
                Me.Close()
                Return
            End If

            ' Initialize controls
            InitializeControls()
            RefreshData()

            ' 🔒 Freeze columns here
            FreezeColumns()

        Catch ex As Exception
            MessageBox.Show("Error loading form: " & ex.Message, "Error")
            Me.Close()
        End Try
    End Sub

    Private Sub FreezeColumns()
        ' Make all columns read-only first
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.ReadOnly = True
        Next

        ' Allow editing only for Column3, Column5, Column6
        DataGridView1.Columns("Column3").ReadOnly = False
        DataGridView1.Columns("Column5").ReadOnly = False
        DataGridView1.Columns("Column6").ReadOnly = False
    End Sub

    Private Sub InitializeControls()
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.ReadOnly = False
        DataGridView1.Rows.Clear()
        ComboBox1.DataSource = Nothing
        ComboBox2.DataSource = Nothing
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

    Public Sub RefreshData()
        Try
            LoadComboBoxData(ComboBox1, "SELECT departmentID, departmentname FROM department WHERE userID = @userID ORDER BY departmentID ASC", "departmentname", "departmentID")
            LoadComboBoxData(ComboBox2, "SELECT positionID, positionname FROM position WHERE userID = @userID ORDER BY positionID ASC", "positionname", "positionID")
        Catch ex As Exception
            MessageBox.Show("Error refreshing data: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub LoadComboBoxData(combo As ComboBox, query As String, displayMember As String, valueMember As String)
        Try
            Using conn As MySqlConnection = GetConnection()
                Dim adapter As New MySqlDataAdapter(query, conn)
                adapter.SelectCommand.Parameters.AddWithValue("@userID", currentUserID)
                Dim table As New DataTable()
                adapter.Fill(table)

                If table.Rows.Count = 0 Then
                    MessageBox.Show($"No data available for {combo.Name} (User ID: {currentUserID})", "Information")
                End If

                combo.DataSource = table
                combo.DisplayMember = displayMember
                combo.ValueMember = valueMember
            End Using
        Catch ex As Exception
            MessageBox.Show($"Error loading combo box data for User {currentUserID}: {ex.Message}", "Error")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            ' Validation: Check combobox selections
            If Not ValidateComboBoxes() Then
                Return
            End If

            Dim selectedDeptID As Integer = Convert.ToInt32(ComboBox1.SelectedValue)
            Dim selectedPosID As Integer = Convert.ToInt32(ComboBox2.SelectedValue)

            Using conn As MySqlConnection = GetConnection()
                conn.Open()

                ' Query filters by userID to load only current user's employees
                Dim query As String = "SELECT employeeID, Full_name, rate, hireddate FROM employee " &
                                    "WHERE departmentID = @deptID AND positionID = @posID AND userID = @userID " &
                                    "ORDER BY Full_name ASC"

                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@deptID", selectedDeptID)
                    cmd.Parameters.AddWithValue("@posID", selectedPosID)
                    cmd.Parameters.AddWithValue("@userID", currentUserID) ' Filter by current user

                    Using reader As MySqlDataReader = cmd.ExecuteReader()
                        DataGridView1.AllowUserToAddRows = False
                        DataGridView1.Rows.Clear()

                        Dim employeeCount As Integer = 0

                        If reader.HasRows Then
                            While reader.Read()
                                If AddEmployeeToGrid(reader) Then
                                    employeeCount += 1
                                End If
                            End While

                            If employeeCount = 0 Then
                                MessageBox.Show($"No eligible employees found for User {currentUserID} (must have at least 14 days tenure).", "Information")
                            Else
                                MessageBox.Show($"Loaded {employeeCount} eligible employee(s) for User {currentUserID}.", "Success")
                            End If
                        Else
                            MessageBox.Show($"No employees found for User {currentUserID} in the selected department and position.", "Information")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As MySqlException
            MessageBox.Show($"Database Error (User {currentUserID}): {ex.Message}", "Error")
        Catch ex As Exception
            MessageBox.Show($"Error retrieving employee data for User {currentUserID}: {ex.Message}", "Error")
        End Try
    End Sub

    Private Function AddEmployeeToGrid(reader As MySqlDataReader) As Boolean
        Try
            Dim fullName As String = reader("Full_name")?.ToString().Trim()
            Dim rateText As String = reader("rate")?.ToString()
            Dim hiredDateText As String = reader("hireddate")?.ToString()

            ' Validation: Check for null or empty values
            If String.IsNullOrWhiteSpace(fullName) OrElse String.IsNullOrWhiteSpace(rateText) Then
                Return False
            End If

            ' Parse rate
            Dim rateValue As Double
            If Not Double.TryParse(rateText, rateValue) OrElse rateValue <= 0 Then
                Return False
            End If

            ' Parse and validate hire date
            Dim hiredDate As DateTime
            If Not DateTime.TryParse(hiredDateText, hiredDate) Then
                Return False
            End If

            ' Check tenure requirement
            Dim daysEmployed As Integer = CInt((DateTimePicker2.Value.Date - hiredDate.Date).TotalDays)
            If daysEmployed < MINIMUM_DAYS_EMPLOYED Then
                Return False
            End If

            ' Add to grid
            Dim rowIndex As Integer = DataGridView1.Rows.Add()
            DataGridView1.Rows(rowIndex).HeaderCell.Value = reader("employeeID").ToString()
            DataGridView1.Rows(rowIndex).Cells("Column1").Value = fullName
            DataGridView1.Rows(rowIndex).Cells("Column2").Value = "₱" & Format(rateValue, "N2")

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function ValidateComboBoxes() As Boolean
        If ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("Please select a Department.", "Validation Error")
            ComboBox1.Focus()
            Return False
        End If

        If ComboBox2.SelectedIndex = -1 Then
            MessageBox.Show("Please select a Position.", "Validation Error")
            ComboBox2.Focus()
            Return False
        End If

        Return True
    End Function

    ' ============================================================================
    ' DATAGRIDVIEW EDITING EVENTS
    ' ============================================================================

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim tb As TextBox = TryCast(e.Control, TextBox)
        If tb Is Nothing Then Return

        ' Remove old handler
        RemoveHandler tb.KeyPress, AddressOf Column3_KeyPress

        ' Add handler only for Column3 (Days Worked)
        If DataGridView1.CurrentCell.ColumnIndex = DataGridView1.Columns("Column3").Index Then
            AddHandler tb.KeyPress, AddressOf Column3_KeyPress
        End If
    End Sub

    Private Sub Column3_KeyPress(sender As Object, e As KeyPressEventArgs)
        ' Allow control keys (Backspace, Delete, etc.)
        If Char.IsControl(e.KeyChar) Then
            Return
        End If

        ' Allow decimal point
        If e.KeyChar = "."c Then
            Return
        End If

        ' Block non-numeric input
        If Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
            Return
        End If

        Dim tb As TextBox = CType(sender, TextBox)
        Dim futureText As String = tb.Text.Remove(tb.SelectionStart, tb.SelectionLength).Insert(tb.SelectionStart, e.KeyChar)

        ' Validate: Check if result exceeds maximum
        Dim val As Double
        If Double.TryParse(futureText, val) Then
            If val > MAXIMUM_DAYS_WORKED Then
                e.Handled = True
                MessageBox.Show($"Maximum {MAXIMUM_DAYS_WORKED} days allowed per payroll period.", "Limit Reached", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        If DataGridView1.Columns(e.ColumnIndex).Name <> "Column3" Then
            Return
        End If

        Dim inputString As String = e.FormattedValue?.ToString().Trim()

        ' Allow empty cells
        If String.IsNullOrWhiteSpace(inputString) Then
            Return
        End If

        ' Validate: Must be numeric
        Dim days As Double
        If Not Double.TryParse(inputString, days) Then
            MessageBox.Show("Days Worked must be a valid number.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True
            Return
        End If

        ' Validate: Must be between 1 and 15
        If days < 1 OrElse days > MAXIMUM_DAYS_WORKED Then
            MessageBox.Show($"Days Worked must be between 1 and {MAXIMUM_DAYS_WORKED}.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True
            Return
        End If
    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then
            Return
        End If

        Try
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Recalculate Gross Pay if Days Worked changed
            If e.ColumnIndex = DataGridView1.Columns("Column3").Index Then
                RecalculateGrossPay(row)
            End If

            ' Recalculate Net Pay if relevant columns changed
            If e.ColumnIndex = DataGridView1.Columns("Column3").Index OrElse
               e.ColumnIndex = DataGridView1.Columns("Column5").Index OrElse
               e.ColumnIndex = DataGridView1.Columns("Column6").Index Then
                RecalculateNetPay(row)
            End If

            UpdateTotals()
        Catch ex As Exception
            MessageBox.Show("Error in calculation: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub RecalculateGrossPay(row As DataGridViewRow)
        Try
            Dim rateText As String = row.Cells("Column2").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
            Dim daysText As String = row.Cells("Column3").Value?.ToString().Trim()

            Dim rate As Double
            Dim daysWorked As Double

            If Double.TryParse(rateText, rate) AndAlso Double.TryParse(daysText, daysWorked) Then
                Dim grossPay As Double = rate * daysWorked
                row.Cells("Column4").Value = "₱" & Format(grossPay, "N2")
            Else
                row.Cells("Column4").Value = "₱0.00"
            End If
        Catch ex As Exception
            MessageBox.Show("Error calculating Gross Pay: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub RecalculateNetPay(row As DataGridViewRow)
        Try
            Dim grossText As String = row.Cells("Column4").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
            Dim benefitsText As String = row.Cells("Column5").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
            Dim deductionText As String = row.Cells("Column6").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()

            Dim grossPay As Double = 0
            Dim benefits As Double = 0
            Dim deduction As Double = 0

            Double.TryParse(grossText, grossPay)
            Double.TryParse(benefitsText, benefits)
            Double.TryParse(deductionText, deduction)

            Dim netPay As Double = grossPay + benefits - deduction
            row.Cells("Column7").Value = "₱" & Format(netPay, "N2")
        Catch ex As Exception
            MessageBox.Show("Error calculating Net Pay: " & ex.Message, "Error")
        End Try
    End Sub

    Private Sub UpdateTotals()
        UpdateTotalGrossPay()
        UpdateTotalNetPay()
    End Sub

    Private Sub UpdateTotalGrossPay()
        Try
            Dim totalGross As Double = 0

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.IsNewRow Then Continue For

                Dim grossText As String = row.Cells("Column4").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
                Dim grossPay As Double

                If Double.TryParse(grossText, grossPay) Then
                    totalGross += grossPay
                End If
            Next

            TextBox1.Text = "₱" & Format(totalGross, "N2")
        Catch ex As Exception
            TextBox1.Text = "₱0.00"
        End Try
    End Sub

    Private Sub UpdateTotalNetPay()
        Try
            Dim totalNet As Double = 0

            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.IsNewRow Then Continue For

                Dim netText As String = row.Cells("Column7").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
                Dim netPay As Double

                If Double.TryParse(netText, netPay) Then
                    totalNet += netPay
                End If
            Next

            TextBox2.Text = "₱" & Format(totalNet, "N2")
        Catch ex As Exception
            TextBox2.Text = "₱0.00"
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim tForm As New t()
        tForm.Show()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ' Validate: Check if data exists
            If DataGridView1.Rows.Count = 0 Then
                MessageBox.Show("No employees to save. Please load employees first.", "Validation Error")
                Return
            End If

            ' Validate: Check combobox selections
            If Not ValidateComboBoxes() Then
                Return
            End If

            ' Validate: Check dates
            If Not ValidateDates() Then
                Return
            End If

            ' Validate: Check TextBox3 format
            If Not ValidateUserID() Then
                Return
            End If

            Dim savedCount As Integer = 0
            Dim skippedCount As Integer = 0

            Using conn As MySqlConnection = GetConnection()
                conn.Open()

                Dim payrollDateFormatted As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")
                Dim createdAtFormatted As String = DateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss")

                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.IsNewRow Then Continue For

                    If SavePayrollRecord(conn, row, payrollDateFormatted, createdAtFormatted) Then
                        savedCount += 1
                    Else
                        skippedCount += 1
                    End If
                Next
            End Using

            ' Show results
            Dim message As String = $"Successfully saved {savedCount} payroll record(s)."
            If skippedCount > 0 Then
                message &= vbCrLf & $"Skipped {skippedCount} incomplete record(s)."
            End If

            MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

            ' Clear form
            InitializeControls()
            RefreshData()

        Catch ex As MySqlException
            MessageBox.Show("Database Error: " & ex.Message, "Error")
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error")
        End Try
    End Sub

    Private Function SavePayrollRecord(conn As MySqlConnection, row As DataGridViewRow, payrollDateFormatted As String, createdAtFormatted As String) As Boolean
        Try
            ' Validate row data
            If row.Cells("Column1").Value Is Nothing OrElse String.IsNullOrWhiteSpace(row.Cells("Column1").Value.ToString()) Then
                Return False
            End If

            Dim dailyRateText As String = row.Cells("Column2").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
            Dim daysWorkedText As String = row.Cells("Column3").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()

            If String.IsNullOrWhiteSpace(dailyRateText) OrElse String.IsNullOrWhiteSpace(daysWorkedText) Then
                Return False
            End If

            ' Parse values
            Dim dailyRate As Double
            Dim daysWorked As Double

            If Not Double.TryParse(dailyRateText, dailyRate) OrElse Not Double.TryParse(daysWorkedText, daysWorked) Then
                Return False
            End If

            Dim grossPayText As String = row.Cells("Column4").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
            Dim benefitsText As String = row.Cells("Column5").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
            Dim deductionText As String = row.Cells("Column6").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()
            Dim netPayText As String = row.Cells("Column7").Value?.ToString().Replace("₱", "").Replace(",", "").Trim()

            Dim grossPay As Double = If(Double.TryParse(grossPayText, Nothing), Convert.ToDouble(grossPayText), 0)
            Dim benefits As Double = If(Double.TryParse(benefitsText, Nothing), Convert.ToDouble(benefitsText), 0)
            Dim deduction As Double = If(Double.TryParse(deductionText, Nothing), Convert.ToDouble(deductionText), 0)
            Dim netPay As Double = If(Double.TryParse(netPayText, Nothing), Convert.ToDouble(netPayText), 0)

            ' Insert into database - SAVE WITH USERID FOR CURRENT USER ONLY
            Dim query As String = "INSERT INTO payrollrecords " &
                                "(PayrollDate, EmployeeID, EmployeeName, DepartmentID, PositionID, DailyRate, DaysWorked, GrossPay, Benefits, Deduction, NetPay, CreatedAt, userID) " &
                                "VALUES (@payrollDate, @employeeID, @employeeName, @deptID, @posID, @dailyRate, @daysWorked, @grossPay, @benefits, @deduction, @netPay, @createdAt, @userID)"

            Using cmd As New MySqlCommand(query, conn)
                Dim employeeID As String = If(row.HeaderCell.Value IsNot Nothing, row.HeaderCell.Value.ToString(), "0")
                Dim userIDFromTextBox As Integer = Convert.ToInt32(TextBox3.Text.Split("-"c)(0))

                ' Validate userID matches current logged-in user
                If userIDFromTextBox <> currentUserID Then
                    MessageBox.Show($"User ID mismatch. Current User: {currentUserID}, TextBox User: {userIDFromTextBox}", "Error")
                    Return False
                End If

                cmd.Parameters.AddWithValue("@payrollDate", payrollDateFormatted)
                cmd.Parameters.AddWithValue("@employeeID", employeeID)
                cmd.Parameters.AddWithValue("@employeeName", row.Cells("Column1").Value.ToString())
                cmd.Parameters.AddWithValue("@deptID", ComboBox1.SelectedValue)
                cmd.Parameters.AddWithValue("@posID", ComboBox2.SelectedValue)
                cmd.Parameters.AddWithValue("@dailyRate", dailyRate)
                cmd.Parameters.AddWithValue("@daysWorked", daysWorked)
                cmd.Parameters.AddWithValue("@grossPay", grossPay)
                cmd.Parameters.AddWithValue("@benefits", benefits)
                cmd.Parameters.AddWithValue("@deduction", deduction)
                cmd.Parameters.AddWithValue("@netPay", netPay)
                cmd.Parameters.AddWithValue("@createdAt", createdAtFormatted)
                cmd.Parameters.AddWithValue("@userID", userIDFromTextBox) ' Save current user's ID

                cmd.ExecuteNonQuery()
                Return True
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function ValidateDates() As Boolean
        If DateTimePicker1.Value.Date > DateTime.Now.Date Then
            MessageBox.Show("Payroll date cannot be in the future.", "Validation Error")
            Return False
        End If

        If DateTimePicker2.Value.Date > DateTime.Now.Date Then
            MessageBox.Show("Creation date cannot be in the future.", "Validation Error")
            Return False
        End If

        Return True
    End Function

    Private Function ValidateUserID() As Boolean
        Try
            If String.IsNullOrWhiteSpace(TextBox3.Text) OrElse Not TextBox3.Text.Contains("-") Then
                MessageBox.Show("Invalid user format. Please log in again.", "Validation Error")
                Return False
            End If

            Dim userIDPart As String = TextBox3.Text.Split("-"c)(0)
            Dim userID As Integer
            If Not Integer.TryParse(userIDPart, userID) OrElse userID <= 0 Then
                MessageBox.Show("Invalid user ID.", "Validation Error")
                Return False
            End If

            Return True
        Catch ex As Exception
            MessageBox.Show("Error validating user: " & ex.Message, "Error")
            Return False
        End Try
    End Function

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ' Can add additional logic here if needed
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        ' Disable the DateTimePicker so the user cannot access it anymore
        DateTimePicker2.Enabled = False
    End Sub
    ' Run this after you set up your DataGridView (e.g., in Form_Load)

End Class