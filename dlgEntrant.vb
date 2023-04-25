Imports System.Windows.Forms
Imports Microsoft.Data.Sqlite

Public Class dlgEntrant
    Public entrants As New ArrayList()
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        ' Save the template
        My.Settings.template = ListBox1.GetItemText(ListBox1.SelectedItem)      ' retrieve setting
        My.Settings.Save()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dlgEntrant_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Load list of contest entrants
        Dim sql As SqliteCommand, sqldr As SqliteDataReader

        sql = Form1.connect.CreateCommand
        ' Load the email template list
        sql.CommandText = "SELECT * FROM email ORDER BY template"
        sqldr = sql.ExecuteReader
        While sqldr.Read
            ListBox1.Items.Add(sqldr("template"))
        End While
        sqldr.Close()
        ListBox1.SelectedIndex = ListBox1.FindString(My.Settings.template)    ' select the current setting

        ' Create a list of entrants
        entrants.Clear()

        sql.CommandText = $"SELECT * FROM Stations WHERE contestID={Form1.contestID} AND `station` <> ''  AND `name` <> '' ORDER BY station,section"     ' get list of entrants
        sqldr = sql.ExecuteReader
        While sqldr.Read
            entrants.Add(New entrant(sqldr("station"), sqldr("section"), sqldr("name"), sqldr("email")))
        End While
        sqldr.Close()

        ' display them in a DataGridView
        With DataGridView1
            .DataSource = entrants
            .AutoResizeColumns(DataGridViewAutoSizeColumnMode.AllCells)  ' Autosize all cells
        End With
    End Sub
    Public Class entrant
        ' represents a single contest entrant
        Property Generate As Boolean        ' will be rendered as a CheckBox
        Property SendEmail As Boolean       ' will be rendered as a CheckBox
        Property Station As String
        Property Section As String
        Property Name As String
        Property Email As String
        Sub New(station As String, section As String, name As String, email As String)
            _Generate = False
            _SendEmail = False
            _Station = station
            _Section = section
            _Name = name
            _Email = email
        End Sub
    End Class

    Private Sub chkGenerate_CheckedChanged(sender As Object, e As EventArgs) Handles chkGenerate.CheckedChanged
        ' toggle all entrants to state of Generate toggle checkbox
        For Each item In entrants
            item.generate = chkGenerate.Checked
        Next
        DataGridView1.Update()
        DataGridView1.Refresh()
    End Sub

    Private Sub chkEmail_CheckedChanged(sender As Object, e As EventArgs) Handles chkEmail.CheckedChanged
        ' toggle all entrants to state of Email toggle checkbox
        For Each item In entrants
            item.sendemail = chkEmail.Checked
        Next
        DataGridView1.Update()
        DataGridView1.Refresh()
    End Sub

End Class
