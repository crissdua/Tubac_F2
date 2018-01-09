Public Class cParametros
    Public con As New Conexion
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Dim frm As New Login
        frm.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text <> "manager" Then
            MessageBox.Show("Unicamente el usuario 'manager' posee privilegios")
        Else
            If con.MakeConnectionSAP() Then
                'MessageBox.Show(con.Connected.ToString)
                Me.Hide()
                Dim frm As New Parametros
                frm.Show()
            Else
                MessageBox.Show("Contraseña erronea")
            End If
        End If
    End Sub
    Private Sub cParametros_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class