Imports System.Data.OleDb

Public Class frmRegistration
    Dim inc As Integer
    Dim MaxRows As Integer
    Dim con As New OleDb.OleDbConnection
    Dim dbProvider As String
    Dim dbSource As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim sql As String




    Private Sub frmRegistration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source = C:\Users\erza6\Documents\AddressBook.mdb"

        con.ConnectionString = dbProvider & dbSource
        con.Open()


        sql = "SELECT * FROM tblContacts"
        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(ds, "AddressBook")
        con.Close()

        MaxRows = ds.Tables("AddressBook").Rows.Count
        inc = -1

    End Sub
    Private Sub navigateRecords()
        txtFirstName.Text = ds.Tables("AddressBook").Rows(inc).Item(1)
        txtLastName.Text = ds.Tables("AddressBook").Rows(inc).Item(2)

    End Sub

    Private Sub BtnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNext.Click
        If inc <> MaxRows - 1 Then
            inc = inc + 1
            navigateRecords()
        Else
            MsgBox("No More Rows")
        End If
    End Sub

    Private Sub BtnPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrevious.Click
        If inc > 0 Then
            inc = inc - 1
            navigateRecords()
        ElseIf inc = -1 Then
            MsgBox("No Records Yet")
        ElseIf inc = 0 Then
            MsgBox("First  Record")
        End If
    End Sub

    Private Sub BtnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
        If inc <> MaxRows - 1 Then
            inc = MaxRows - 1
            navigateRecords()

        End If
    End Sub

    Private Sub BtnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click
        If inc <> 0 Then
            inc = 0
            navigateRecords()

        End If
    End Sub

    Private Sub BtnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Dim cb As New OleDb.OleDbCommandBuilder(da)

        ds.Tables("AddressBook").Rows(inc).Item(1) = txtFirstName.Text
        ds.Tables("AddressBook").Rows(inc).Item(2) = txtLastName.Text


        da.Update(ds, "AddressBook")
        MsgBox("Data updated successfully")
    End Sub

    Private Sub BtnCommit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCommit.Click
        If inc <> -1 Then

            Dim cb As New OleDb.OleDbCommandBuilder(da)
            Dim dsNewRow As DataRow

            dsNewRow = ds.Tables("AddressBook").NewRow()

            dsNewRow.Item("FirstName") = txtFirstName.Text
            dsNewRow.Item("SurName") = txtLastName.Text


            ds.Tables("AddressBook").Rows.Add(dsNewRow)
            da.Update(ds, "AddressBook")
            MsgBox("New Record added to the database")

            btnCommit.Enabled = False
            btnAdd.Enabled = True
            btnUpdate.Enabled = True
            btnDelete.Enabled = True

            txtFirstName.Clear()
            txtLastName.Clear()



        End If
    End Sub





    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Dim cb As New OleDb.OleDbCommandBuilder(da)


        If MessageBox.Show("Do you really want to Delete this Record?",
"Delete", MessageBoxButtons.YesNo,
MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then

            ds.Tables("AddressBook").Rows(inc).Delete()
            MaxRows = MaxRows - 1
            da.Update(ds, "AddressBook")
            txtFirstName.Clear()
            txtLastName.Clear()

        Else

            MsgBox("Operation Cancelled")

            Exit Sub

        End If

    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        btnCommit.Enabled = False

        btnAdd.Enabled = True
        btnUpdate.Enabled = True
        btnDelete.Enabled = True

        txtFirstName.Clear()
        txtLastName.Clear()


    End Sub

    Private Sub BtnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        btnCommit.Enabled = True
        btnAdd.Enabled = False
        btnUpdate.Enabled = False
        btnDelete.Enabled = False

        txtFirstName.Clear()
        txtLastName.Clear()

        inc = 0
    End Sub




    Private Sub BtnClick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClick.Click
        frmSearch.Show()
    End Sub

End Class
