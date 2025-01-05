Public Class frmAttendantsDelete
    Private Sub frmAttendantDelete_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Try
            strSelect = "SELECT intAttendantID, strFirstName + ' ' + strLastName as AttendantName From TAttendants"

            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            dt.Load(drSourceTable)

            cboAttendants.ValueMember = "intAttendantID"
            cboAttendants.DisplayMember = "AttendantName"
            cboAttendants.DataSource = dt


        Catch ex As Exception
            ' Log and display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim results As DialogResult
        Dim strDelete As String
        Dim intRowsAffected As Integer
        Dim cmdDelete As OleDb.OleDbCommand
        Try

            If OpenDatabaseConnectionSQLServer() = False Then

                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                   "The application will now close.",
                                   Me.Text + " Error",
                                   MessageBoxButtons.OK, MessageBoxIcon.Error)

                Me.Close()
            End If
            results = MessageBox.Show("Are you sure you want to delete attendant: " & cboAttendants.Text & "?", "Confrim Deletion", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case results
                Case DialogResult.Cancel
                    MessageBox.Show("Action Canceled")
                Case DialogResult.No
                    MessageBox.Show("Action Canceled")
                Case DialogResult.Yes
                    strDelete = "Delete FROM TAttendants Where intAttendantID = " & cboAttendants.SelectedValue.ToString
                    cmdDelete = New OleDb.OleDbCommand(strDelete, m_conAdministrator)
                    intRowsAffected = cmdDelete.ExecuteNonQuery()
            End Select
            If intRowsAffected > 0 Then
                MessageBox.Show("Attendant has been Deleted")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Dim frmAttendantsDelete As New frmAttendantsDelete
        frmAttendantsDelete.Show()
        Me.Hide()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Dim frmAdminA As New frmAdminA
        frmAdminA.Show()
        Me.Hide()
    End Sub
End Class