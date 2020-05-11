Public Class Form1
    Dim sql As String
    Sub panggildata()
        Call konek()
        da = New OleDb.OleDbDataAdapter("SELECT *FROM tb_kue", conect)
        nds = New DataSet
        nds.Clear()
        da.Fill(nds, "tb_kue")
        DataGridView1.DataSource = nds.Tables("tb_kue")
        DataGridView1.Enabled = True
    End Sub
    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = conect
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sql
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        TextBox1.Text = " "
        TextBox2.Text = " "
        TextBox3.Text = " "
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sql = ("insert into tb_kue (id,nama,resiko) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')")
        Call jalan()
        MsgBox("Data Berhasil Tersimpan")
        Call panggildata()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        sql = "delete from tb_kue where id = " & TextBox1.Text & " "
        Call jalan()
        MsgBox("Data telah terhapus")
        Call panggildata()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form2.Show()
        Me.Hide()
    End Sub

End Class