Public Class LoginForm1

    ' TODO: inserte el código para realizar autenticación personalizada usando el nombre de usuario y la contraseña proporcionada 
    ' (Consulte http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' El objeto principal personalizado se puede adjuntar al objeto principal del subproceso actual como se indica a continuación: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' donde CustomPrincipal es la implementación de IPrincipal utilizada para realizar la autenticación. 
    ' Posteriormente, My.User devolverá la información de identidad encapsulada en el objeto CustomPrincipal
    ' como el nombre de usuario, nombre para mostrar, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        If UsernameTextBox.Text = "" And PasswordTextBox.Text = "" Then
            Label1.Text = "Introduzca un nombre de usuario y contraseña por favor"
        Else
            If UsernameTextBox.Text = "" Then
                Label1.Text = "Introduzca un nombre de usuario por favor"
            Else
                If UsernameTextBox.Text = "Administrador" Then

                    If PasswordTextBox.Text = "" Then
                        Label1.Text = "Introduzca la contraseña por favor"
                    Else
                        If PasswordTextBox.Text = "841011foto" Then
                            Form1.Show()
                            Me.Hide()
                        Else
                            Label1.Text = "Contraseña invalida"
                        End If
                    End If
                Else
                    Label1.Text = "El usuario no existe"
                End If

                End If
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
    End Sub
End Class
