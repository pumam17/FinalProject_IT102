Public Class frmPassengersAdd
    Private Sub frmAddPassenger_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim strSelect As String = ""
            Dim cmdSelect As OleDb.OleDbCommand
            Dim drSourceTable As OleDb.OleDbDataReader
            Dim dt As DataTable = New DataTable

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                "The application will now close.",
                                Me.Text + " Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()

            End If
            strSelect = "SELECT intStateID, strState From TStates"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dt.Load(drSourceTable)

            cboStates.ValueMember = "intStateID"
            cboStates.DisplayMember = "strState"
            cboStates.DataSource = dt
        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim strFirstName As String
        Dim strLastName As String
        Dim strAddress As String
        Dim strCity As String
        Dim intState As Integer
        Dim intZip As Integer
        Dim intPhoneNumber As Integer
        Dim strEmail As String
        Dim strLogin As String
        Dim strPassword As String
        Dim dtmDOB As String
        Dim blnValidation As Boolean = True
        Get_And_Validate_Input(strFirstName, strLastName, strAddress, strCity, intZip, intPhoneNumber, strEmail, strLogin, strPassword, dtmDOB, blnValidation)
        If blnValidation = True Then
            intState = cboStates.SelectedValue

            Dim cmdAddPassenger As New OleDb.OleDbCommand()
            Dim intPKID As Integer
            Dim intRowsAffected As Integer

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                               "The application will now close.",
                               Me.Text + " Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error)

                Me.Close()
            End If

            cmdAddPassenger.CommandText = "EXECUTE uspAddPassenger " & intPKID & ", '" & strFirstName & "', '" & strLastName & "', '" & strAddress & "', '" & strCity & "', " & intState & ", '" & intZip & "', '" & intPhoneNumber & "', '" & strEmail & "', '" & strLogin & "', '" & strPassword & "', '" & dtmDOB & "'"
            cmdAddPassenger.CommandType = CommandType.StoredProcedure
            cmdAddPassenger = New OleDb.OleDbCommand(cmdAddPassenger.CommandText, m_conAdministrator)
            intRowsAffected = cmdAddPassenger.ExecuteNonQuery()
            If intRowsAffected > 0 Then
                MessageBox.Show("Passenger has been added")
            End If
            CloseDatabaseConnection()
            Close()
            Dim frmPassengersAdd As New frmPassengersAdd
            frmPassengersAdd.Show()
        Else
            Dim frmPassengersAdd As New frmPassengersAdd
            frmPassengersAdd.Show()
            Me.Hide()
            Exit Sub
        End If
    End Sub
    Private Sub Get_And_Validate_Input(ByRef strFirstName As String, ByRef strLastName As String, ByRef strAddress As String, ByRef strCity As String, ByRef intZip As Integer, ByRef intPhoneNumber As Integer, ByRef strEmail As String, ByRef strLogin As String, ByRef strPassword As String, ByRef strDOB As String, ByRef blnValidation As Boolean)
        Validate_FirstName(strFirstName, blnValidation)
        Validate_LastName(strLastName, blnValidation)
        Validate_Address(strAddress, blnValidation)
        Validate_City(strCity, blnValidation)
        Validate_Zip(intZip, blnValidation)
        Validate_Phone(intPhoneNumber, blnValidation)
        Validate_Email(strEmail, blnValidation)
        Validate_Login(strLogin, blnValidation)
        Validate_Password(strPassword, blnValidation)
        Validate_DateOfBirth(strDOB, blnValidation)
    End Sub

    Private Sub Validate_FirstName(ByRef strFirstName As String, ByRef blnValidation As Boolean)
        If txtFristName.Text = String.Empty Then
            MessageBox.Show("First Name Must Exist")
            txtFristName.Focus()
            blnValidation = False
        Else
            strFirstName = txtFristName.Text
        End If
    End Sub

    Private Sub Validate_LastName(ByRef strLastName As String, ByRef blnValidation As Boolean)
        If txtLastName.Text = String.Empty Then
            MessageBox.Show("Last Name Must Exist")
            txtLastName.Focus()
            blnValidation = False
        Else
            strLastName = txtLastName.Text
        End If
    End Sub

    Private Sub Validate_Address(ByRef strAddress As String, ByRef blnValidation As Boolean)
        If txtAddress.Text = String.Empty Then
            MessageBox.Show("Address Must Exist")
            txtAddress.Focus()
            blnValidation = False
        Else
            strAddress = txtAddress.Text
        End If
    End Sub

    Private Sub Validate_City(ByRef strCity As String, ByRef blnValidation As Boolean)
        If txtCity.Text = String.Empty Then
            MessageBox.Show("City Must Exist")
            txtCity.Focus()
            blnValidation = False
        Else
            strCity = txtCity.Text
        End If
    End Sub

    Private Sub Validate_Zip(ByRef intZip As Integer, ByRef blnValidation As Boolean)
        If Integer.TryParse(txtZip.Text, intZip) Then
            If intZip.ToString.Length <> 5 Then
                intZip = txtZip.Text
                MessageBox.Show("Zipcode Must be 5 numbers")
                txtZip.Focus()
                blnValidation = False
            End If
        Else
            intZip = txtZip.Text
        End If
    End Sub

    Private Sub Validate_Phone(ByRef intPhoneNumber As Integer, ByRef blnValidation As Boolean)
        If Integer.TryParse(txtPhone.Text, intPhoneNumber) Then
            intPhoneNumber = txtPhone.Text
            If intPhoneNumber.ToString.Length <> 10 Then
                blnValidation = False
                MessageBox.Show("Phone Number Must be Ten Characters")
                txtPhone.Focus()
            End If
        Else
            MessageBox.Show("Phone Number Must Exist")
            txtPhone.Focus()
            blnValidation = False
        End If
    End Sub

    Private Sub Validate_Email(ByRef strEmail As String, ByRef blnValidation As Boolean)
        If txtEmail.Text = String.Empty Then
            MessageBox.Show("Email Must Exist")
            txtEmail.Focus()
            blnValidation = False
        Else
            strEmail = txtEmail.Text
            If strEmail.IndexOf("@") = -1 Then
                MessageBox.Show("Email Must Include @")
                txtEmail.Focus()
                blnValidation = False
            End If
        End If
    End Sub

    Private Sub Validate_Login(ByRef strLogin As String, ByRef blnValidation As Boolean)
        If txtUsername.Text = String.Empty Then
            MessageBox.Show("Username Must Exist")
            txtUsername.Focus()
            blnValidation = False
        Else
            strLogin = txtUsername.Text
        End If
    End Sub

    Private Sub Validate_Password(ByRef strPassword As String, ByRef blnValidation As Boolean)
        If txtPassword.Text = String.Empty Then
            MessageBox.Show("Password Must Exist")
            txtPassword.Focus()
            blnValidation = False
        Else
            strPassword = txtPassword.Text
        End If
    End Sub
    Private Sub Validate_DateOfBirth(ByRef dtmDOB As String, ByRef blnValidation As Boolean)
        Dim strString As String
        If txtDOB.Text = String.Empty Then
            blnValidation = ValidationDate()
        Else
            dtmDOB = txtDOB.Text
            If dtmDOB.Length <> 10 Then
                blnValidation = ValidationDate()
            Else
                strString = dtmDOB.Substring(0, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDate()
                    End If
                    If strString > 12 Then
                        blnValidation = ValidationDate()
                    End If
                Else
                    blnValidation = ValidationDate()
                End If
                strString = dtmDOB.Substring(3, 2)
                If IsNumeric(strString) Then
                    If strString < 0 Then
                        blnValidation = ValidationDate()
                    End If
                    If strString > 31 Then
                        blnValidation = ValidationDate()
                    End If
                Else
                    blnValidation = ValidationDate()
                End If
                strString = dtmDOB.Substring(6, 4)
                If IsNumeric(strString) Then
                    If dtmDOB.IndexOf("/") <> 2 And 5 Then
                        blnValidation = ValidationDate()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmPLogin As New frmPLogin
        frmPLogin.Show()
        Me.Hide()
    End Sub

    Private Function ValidationDate()
        Dim blnValidation As Boolean
        MessageBox.Show("Date of birth must be vaild format 00/00/0000")
        txtDOB.Focus()
        blnValidation = False
        Return blnValidation
    End Function
End Class