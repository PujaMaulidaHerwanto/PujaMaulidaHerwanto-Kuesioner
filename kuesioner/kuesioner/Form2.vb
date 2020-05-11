Public Class Form2
    Dim nilai1, nilai2, nilai3, nilai4, nilai5, nilai6, nilai7, nilai8, nilai9, nilai10 As Integer
    Dim nilai11, nilai12, nilai13, nilai14, nilai15, nilai16 As Integer
    Dim nilai17, nilai18, nilai19, nilai20, nilai21 As Integer
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
        Label13.Text = " "
    End Sub
    Sub baru()
        Dim lagi As New System.Data.OleDb.OleDbCommand
        Call konek()
        lagi.Connection = conect
        lagi.CommandType = CommandType.Text
        lagi.CommandText = sql
        lagi.ExecuteNonQuery()
        lagi.Dispose()
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox1.Checked = nilai1
            nilai1 = 1
        Else
            CheckBox1.Checked = nilai1
            nilai1 = 0
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox2.Checked = nilai2
            nilai2 = 1
        Else
            CheckBox2.Checked = nilai2
            nilai2 = 0
        End If
    End Sub
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            CheckBox3.Checked = nilai3
            nilai3 = 1
        Else
            CheckBox3.Checked = nilai3
            nilai3 = 0
        End If
    End Sub


    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            CheckBox4.Checked = nilai4
            nilai4 = 1
        Else
            CheckBox4.Checked = nilai4
            nilai4 = 0
        End If
    End Sub
    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            CheckBox5.Checked = nilai5
            nilai5 = 1
        Else
            CheckBox5.Checked = nilai5
            nilai5 = 0
        End If
    End Sub
    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            CheckBox6.Checked = nilai6
            nilai6 = 1
        Else
            CheckBox6.Checked = nilai6
            nilai6 = 0
        End If
    End Sub
    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            CheckBox7.Checked = nilai7
            nilai7 = 1
        Else
            CheckBox7.Checked = nilai7
            nilai7 = 0
        End If
    End Sub
    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = True Then
            CheckBox8.Checked = nilai8
            nilai8 = 1
        Else
            CheckBox8.Checked = nilai8
            nilai8 = 0
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = True Then
            CheckBox9.Checked = nilai9
            nilai9 = 1
        Else
            CheckBox9.Checked = nilai9
            nilai9 = 0
        End If
    End Sub
    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then
            CheckBox10.Checked = nilai10
            nilai10 = 1
        Else
            CheckBox10.Checked = nilai10
            nilai10 = 0
        End If
    End Sub
    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = True Then
            CheckBox11.Checked = nilai11
            nilai11 = 1
        Else
            CheckBox11.Checked = nilai11
            nilai11 = 0
        End If
    End Sub
    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = True Then
            CheckBox12.Checked = nilai12
            nilai12 = 1
        Else
            CheckBox12.Checked = nilai12
            nilai12 = 0
        End If
    End Sub
    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = True Then
            CheckBox13.Checked = nilai13
            nilai13 = 1
        Else
            CheckBox13.Checked = nilai13
            nilai13 = 0
        End If
    End Sub
    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        If CheckBox14.Checked = True Then
            CheckBox14.Checked = nilai14
            nilai14 = 1
        Else
            CheckBox14.Checked = nilai14
            nilai14 = 0
        End If
    End Sub
    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = True Then
            CheckBox15.Checked = nilai15
            nilai15 = 1
        Else
            CheckBox15.Checked = nilai15
            nilai15 = 0
        End If
    End Sub
    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = True Then
            CheckBox16.Checked = nilai16
            nilai16 = 1
        Else
            CheckBox16.Checked = nilai16
            nilai16 = 0
        End If
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        If CheckBox17.Checked = True Then
            CheckBox17.Checked = nilai17
            nilai17 = 1
        Else
            CheckBox17.Checked = nilai17
            nilai17 = 0
        End If
    End Sub
    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        If CheckBox18.Checked = True Then
            CheckBox18.Checked = nilai18
            nilai18 = 1
        Else
            CheckBox18.Checked = nilai18
            nilai18 = 0
        End If
    End Sub
    Private Sub CheckBox19_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox19.CheckedChanged
        If CheckBox19.Checked = True Then
            CheckBox19.Checked = nilai19
            nilai19 = 1
        Else
            CheckBox19.Checked = nilai19
            nilai19 = 0
        End If
    End Sub
    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox20.CheckedChanged
        If CheckBox20.Checked = True Then
            CheckBox20.Checked = nilai20
            nilai20 = 1
        Else
            CheckBox20.Checked = nilai20
            nilai20 = 0
        End If
    End Sub
    Private Sub CheckBox21_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged
        If CheckBox21.Checked = True Then
            CheckBox21.Checked = nilai21
            nilai21 = 1
        Else
            CheckBox21.Checked = nilai21
            nilai21 = 0
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label8.Text = Val(nilai1) + (nilai2) + (nilai3) + (nilai4) + (nilai5) + (nilai6) + (nilai7) + (nilai8) + (nilai9) + (nilai10)
        Label9.Text = Val(nilai11) + Val(nilai12) + Val(nilai13) + Val(nilai14) + Val(nilai15) + Val(nilai16)
        Label10.Text = Val(nilai17) + Val(nilai18) + Val(nilai19) + Val(nilai20) + Val(nilai21)
        Label11.Text = Val(Label8.Text) + Val(Label9.Text) + Val(Label10.Text)
        If Label11.Text <= 7 Then
            Label13.Text = "Resiko Rendah"
        ElseIf Label11.Text <= 14 Then
            Label13.Text = "Resiko Sedang"
        Else
            Label13.Text = "Resiko Tinggi"
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        sql = "delete from tb_kue where nama = '" & TextBox2.Text & "'"
        Call baru()
        sql = ("insert into tb_kue (id,nama,resiko) values ('" & TextBox1.Text & "','" & TextBox2.Text & "' , '" & Label13.Text & "')")
        Call jalan()
        MsgBox("Data Berhasil Disimpan")
        Call panggildata()

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        Label13.Text = DataGridView1.Item(2, i).Value
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub




End Class
