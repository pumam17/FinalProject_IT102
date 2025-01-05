Imports System.Drawing.Text

Public Class frmPLogin
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmMainMenu As New frmMainMenu
        frmMainMenu.Show()
        Me.Hide()
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim strUserName As String
        Dim strPassword As String
        strUserName = txtUserName.Text
        strPassword = txtPassword.Text

        Try
            Dim cmdSelect As OleDb.OleDbCommand        ' this will be used for our Select statement
            Dim objParam As OleDb.OleDbParameter
            Dim drSourceTable As OleDb.OleDbDataReader ' this will be where our result set will

            ' open the DB
            If OpenDatabaseConnectionSQLServer() = False Then
                ' No, warn the user ...
                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                    "The application will now close.",
                                    Me.Text + " Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                ' and close the form/application
                Me.Close()
            End If
            ' Build the select statement

            cmdSelect = New OleDb.OleDbCommand("uspPLogin", m_conAdministrator)
            cmdSelect.CommandType = CommandType.StoredProcedure
            objParam = cmdSelect.Parameters.Add("@strPassengerLoginID", OleDb.OleDbType.VarChar)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = strUserName

            objParam = cmdSelect.Parameters.Add("@strPassengerPassword", OleDb.OleDbType.VarChar)
            objParam.Direction = ParameterDirection.Input
            objParam.Value = strPassword
            drSourceTable = cmdSelect.ExecuteReader
            While drSourceTable.Read()
                strString = drSourceTable("intPassengerID")
            End While
            ' Clean up
            drSourceTable.Close()
            ' close the database connection
            CloseDatabaseConnection()
        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
        LoginConnection()

    End Sub

    Private Sub LoginConnection()
        If CInt(strString) > 0 Then

            Dim frmPassengerMenu As New frmPassengersMenu
            frmPassengerMenu.Show()

            Me.Hide()
        Else
            MessageBox.Show("Incorrect login")
        End If

    End Sub

    Private Sub btnAddPassengers_Click(sender As Object, e As EventArgs) Handles btnAddPassengers.Click
        Dim frmPassengersAdd As New frmPassengersAdd
        frmPassengersAdd.Show()
        Me.Hide()
    End Sub
End Class