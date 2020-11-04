Imports System.Data.OleDb

Public Class frmSearch
    Dim inc As Integer
    Dim MaxRows As Integer
    Dim con As New OleDb.OleDbConnection
    Dim dbProvider As String
    Dim dbSource As String
    Dim ds As New DataSet
    Dim da As OleDb.OleDbDataAdapter
    Dim sql As String

    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        If txtSearch.Text = "" Then
            MessageBox.Show("Please Enter First Name")
        Else

            dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
            dbSource = "Data Source = C:\Users\erza6\Documents\AddressBook.mdb"

            con.ConnectionString = dbProvider & dbSource



            sql = "select* from tblContacts where FirstName " & "like '%" & txtSearch.Text & "%'"
            da = New OleDb.OleDbDataAdapter(sql, con)
            da.Fill(ds, "AddressBook")


            DataGrid1.DataSource = ds.Tables(0)

            btnSearch.Enabled = False
            btnReset.Enabled = True

        End If


    End Sub





    Private Sub BtnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        btnSearch.Enabled = True
        btnReset.Enabled = False
        txtSearch.Text = ""

    End Sub

    Private Sub frmSearch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'AddressBookDataSet1.tblContacts' table. You can move, or remove it, as needed.
        Me.TblContactsTableAdapter.Fill(Me.AddressBookDataSet1.tblContacts)

    End Sub

    Private Sub DataGrid1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid1.CellContentClick

        Dim stuid As String
        stuid = DataGrid1.Item(DataGrid1.CurrentCell.DefaultNewRowValue, 0)
    End Sub
End Class