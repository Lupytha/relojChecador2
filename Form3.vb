Imports System.Data.OleDb

Public Class Form3
    Dim basedatos As OleDbDataAdapter
    Dim Conexion, selectsemanas, selectempleados As String
    Dim cadenaconexion As OleDbConnection
    Dim horarios As DataTable
    Dim semanas, empleados As DataTable
    Dim comando As OleDbCommand
    Dim consulta, nomhorario As String
    Dim lunes, martes, miercoles, jueves, viernes, sabado, domingo As String
    Dim tlunes, tmartes, tmiercoles, tjueves, tviernes, tsabado, tdomingo, ttotals As String
    Dim tlunesb, tmartesb, tmiercolesb, tjuevesb, tviernesb, tsabadob, tdomingob As String
    Dim lunest, martest, miercolest, juevest, viernest, sabadot, domingot As Integer
    Dim ttotal As Integer
    ' Dim rowsem As DataRow

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Conexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Program Files (x86)\TimeWork Reloj Checador\fcf.mdb';" & _
                  "Persist Security Info=True;Jet OLEDB:Database Password=YaDRMlMPtdwYEAdtDNtL"

        selectsemanas = "SELECT * FROM semanas WHERE YEAR(fechaini) = YEAR(Date())"
        selectempleados = "SELECT id, numero, nombre, apellidos FROM  empleado where activo=true"

        basedatos = New OleDbDataAdapter(selectsemanas, Conexion)
        cadenaconexion = New OleDbConnection(Conexion)

        Try
            cadenaconexion.Open()
            comando = New OleDbCommand(selectsemanas, cadenaconexion)
            comando.ExecuteNonQuery()
            semanas = New DataTable
            basedatos.Fill(semanas)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        cadenaconexion.Close()
        ComboBox1.DataSource = semanas
        DataGridView1.DataSource = semanas

        fechaini.Text = Convert.ToDateTime(DataGridView1.CurrentRow.Cells(1).Value()).Date
        fechafin.Text = Convert.ToDateTime(DataGridView1.CurrentRow.Cells(2).Value()).Date


        basedatos = New OleDbDataAdapter(selectempleados, Conexion)
        cadenaconexion = New OleDbConnection(Conexion)

        Try
            cadenaconexion.Open()
            empleados = New DataTable
            basedatos.Fill(empleados)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        cadenaconexion.Close()
        ComboBox2.DataSource = empleados
        DataGridView2.DataSource = empleados

        ent1.Enabled = False
        ent2.Enabled = False
        ent3.Enabled = False
        ent4.Enabled = False
        ent5.Enabled = False
        ent6.Enabled = False
        ent7.Enabled = False

        sal1.Enabled = False
        sal2.Enabled = False
        sal3.Enabled = False
        sal4.Enabled = False
        sal5.Enabled = False
        sal6.Enabled = False
        sal7.Enabled = False

        entb1.Enabled = False
        entb2.Enabled = False
        entb3.Enabled = False
        entb4.Enabled = False
        entb5.Enabled = False
        entb6.Enabled = False
        entb7.Enabled = False

        salb1.Enabled = False
        salb2.Enabled = False
        salb3.Enabled = False
        salb4.Enabled = False
        salb5.Enabled = False
        salb6.Enabled = False
        salb7.Enabled = False

        cbb1.Enabled = False
        cbb2.Enabled = False
        cbb3.Enabled = False
        cbb4.Enabled = False
        cbb5.Enabled = False
        cbb6.Enabled = False
        cbb7.Enabled = False

    End Sub

    Private Sub Salir_Click(sender As Object, e As EventArgs) Handles Salir.Click
        Me.Close()
    End Sub

    Private Sub cb1_CheckedChanged(sender As Object, e As EventArgs) Handles cb1.CheckedChanged
        If cb1.CheckState = 1 Then
            ent1.Enabled = True
            sal1.Enabled = True
            cbb1.Enabled = True
        Else
            ent1.Enabled = False
            sal1.Enabled = False
            cbb1.Enabled = False
            cbb1.CheckState = 0
            entb1.Enabled = False
            salb1.Enabled = False
        End If
    End Sub

    Private Sub cbb1_CheckedChanged(sender As Object, e As EventArgs) Handles cbb1.CheckedChanged
        If cbb1.CheckState = 1 Then
            entb1.Enabled = True
            salb1.Enabled = True
        Else
            entb1.Enabled = False
            salb1.Enabled = False
        End If
    End Sub

    Private Sub cbb4_CheckedChanged(sender As Object, e As EventArgs) Handles cbb4.CheckedChanged
        If cbb4.CheckState = 1 Then
            entb4.Enabled = True
            salb4.Enabled = True
        Else
            entb4.Enabled = False
            salb4.Enabled = False
        End If
    End Sub

    Private Sub cbb3_CheckedChanged(sender As Object, e As EventArgs) Handles cbb3.CheckedChanged
        If cbb3.CheckState = 1 Then
            entb3.Enabled = True
            salb3.Enabled = True
        Else
            entb3.Enabled = False
            salb3.Enabled = False
        End If
    End Sub

    Private Sub cbb2_CheckedChanged(sender As Object, e As EventArgs) Handles cbb2.CheckedChanged
        If cbb2.CheckState = 1 Then
            entb2.Enabled = True
            salb2.Enabled = True
        Else
            entb2.Enabled = False
            salb2.Enabled = False
        End If
    End Sub

    Private Sub cbb7_CheckedChanged(sender As Object, e As EventArgs) Handles cbb7.CheckedChanged
        If cbb7.CheckState = 1 Then
            entb7.Enabled = True
            salb7.Enabled = True
        Else
            entb7.Enabled = False
            salb7.Enabled = False
        End If
    End Sub

    Private Sub cb6_CheckedChanged(sender As Object, e As EventArgs) Handles cb6.CheckedChanged
        If cb6.CheckState = 1 Then
            ent6.Enabled = True
            sal6.Enabled = True
            cbb6.Enabled = True
        Else
            ent6.Enabled = False
            sal6.Enabled = False
            cbb6.Enabled = False
            cbb6.CheckState = 0
            entb6.Enabled = False
            salb6.Enabled = False
        End If
    End Sub

    Private Sub cb5_CheckedChanged(sender As Object, e As EventArgs) Handles cb5.CheckedChanged
        If cb5.CheckState = 1 Then
            ent5.Enabled = True
            sal5.Enabled = True
            cbb5.Enabled = True
        Else
            ent5.Enabled = False
            sal5.Enabled = False
            cbb5.Enabled = False
            cbb5.CheckState = 0
            entb5.Enabled = False
            salb5.Enabled = False
        End If
    End Sub

    Private Sub cb4_CheckedChanged(sender As Object, e As EventArgs) Handles cb4.CheckedChanged
        If cb4.CheckState = 1 Then
            ent4.Enabled = True
            sal4.Enabled = True
            cbb4.Enabled = True
        Else
            ent4.Enabled = False
            sal4.Enabled = False
            cbb4.Enabled = False
            cbb4.CheckState = 0
            entb4.Enabled = False
            salb4.Enabled = False
        End If
    End Sub

    Private Sub cb3_CheckedChanged(sender As Object, e As EventArgs) Handles cb3.CheckedChanged
        If cb3.CheckState = 1 Then
            ent3.Enabled = True
            sal3.Enabled = True
            cbb3.Enabled = True
        Else
            ent3.Enabled = False
            sal3.Enabled = False
            cbb3.Enabled = False
            cbb3.CheckState = 0
            entb3.Enabled = False
            salb3.Enabled = False
        End If
    End Sub

    Private Sub cb2_CheckedChanged(sender As Object, e As EventArgs) Handles cb2.CheckedChanged
        If cb2.CheckState = 1 Then
            ent2.Enabled = True
            sal2.Enabled = True
            cbb2.Enabled = True
        Else
            ent2.Enabled = False
            sal2.Enabled = False
            cbb2.Enabled = False
            cbb2.CheckState = 0
            entb2.Enabled = False
            salb2.Enabled = False
        End If
    End Sub

    Private Sub cb7_CheckedChanged(sender As Object, e As EventArgs) Handles cb7.CheckedChanged
        If cb7.CheckState = 1 Then
            ent7.Enabled = True
            sal7.Enabled = True
            cbb7.Enabled = True
        Else
            ent7.Enabled = False
            sal7.Enabled = False
            cbb7.Enabled = False
            cbb7.CheckState = 0
        End If
    End Sub

    Private Sub cbb5_CheckedChanged(sender As Object, e As EventArgs) Handles cbb5.CheckedChanged
        If cbb5.CheckState = 1 Then
            entb5.Enabled = True
            salb5.Enabled = True
        Else
            entb5.Enabled = False
            salb5.Enabled = False
        End If
    End Sub

    Private Sub cbb6_CheckedChanged(sender As Object, e As EventArgs) Handles cbb6.CheckedChanged
        If cbb6.CheckState = 1 Then
            entb6.Enabled = True
            salb6.Enabled = True
        Else
            entb6.Enabled = False
            salb6.Enabled = False
        End If
    End Sub

    Private Sub Rellenar_Click(sender As Object, e As EventArgs) Handles Rellenar.Click
        ent1.Value = ent1.Value
        ent2.Value = ent1.Value
        ent3.Value = ent1.Value
        ent4.Value = ent1.Value
        ent5.Value = ent1.Value
        ent6.Value = ent1.Value
        ent7.Value = ent1.Value

        sal1.Value = sal1.Value
        sal2.Value = sal1.Value
        sal3.Value = sal1.Value
        sal4.Value = sal1.Value
        sal5.Value = sal1.Value
        sal6.Value = sal1.Value
        sal7.Value = sal1.Value

        entb1.Value = entb1.Value
        entb2.Value = entb1.Value
        entb3.Value = entb1.Value
        entb4.Value = entb1.Value
        entb5.Value = entb1.Value
        entb6.Value = entb1.Value
        entb7.Value = entb1.Value

        salb1.Value = salb1.Value
        salb2.Value = salb1.Value
        salb3.Value = salb1.Value
        salb4.Value = salb1.Value
        salb5.Value = salb1.Value
        salb6.Value = salb1.Value
        salb7.Value = salb1.Value
    End Sub

    Private Sub Guardar_Click(sender As Object, e As EventArgs) Handles Guardar.Click
      
        If cb1.CheckState Then
            If cbb1.CheckState Then
                tlunes = DateDiff(DateInterval.Hour, ent1.Value, sal1.Value)
                tlunesb = DateDiff(DateInterval.Hour, entb1.Value, salb1.Value)
                lunest = tlunes - tlunesb
                ttotal = ttotal + lunest
                lunes = "'" & ent1.Value & "',' " & sal1.Value & "',' " & entb1.Value & "',' " & salb1.Value & "','" & tlunes & "','03/01/2000'"
            Else
                tlunes = DateDiff(DateInterval.Hour, ent1.Value, sal1.Value)
                ttotal = ttotal + tlunes
                lunes = "' " & ent1.Value & "',' " & sal1.Value & "',NULL,NULL,'" & tlunes & "','03/01/2000'"
            End If
        Else
            lunes = "Null,Null,Null,Null,Null,Null"
        End If

        If cb2.CheckState Then
            If cbb2.CheckState Then
                tmartes = DateDiff(DateInterval.Hour, ent2.Value, sal2.Value)
                tmartesb = DateDiff(DateInterval.Hour, entb2.Value, salb2.Value)
                martest = tmartes - tmartesb
                ttotal = ttotal + martest
                martes = "'" & ent2.Value & "',' " & sal2.Value & "',' " & entb2.Value & "',' " & salb2.Value & "','" & tmartes & "','04/01/2000'"
            Else
                tmartes = DateDiff(DateInterval.Hour, ent2.Value, sal2.Value)
                ttotal = ttotal + tmartes
                martes = "'" & ent2.Value & "',' " & sal2.Value & "',NULL,NULL,'" & tmartes & "','04/01/2000'"

            End If
        Else
            martes = "Null,Null,Null,Null,Null,Null"
        End If
        If cb3.CheckState Then
            If cbb3.CheckState Then
                tmiercoles = DateDiff(DateInterval.Hour, ent3.Value, sal3.Value)
                tmiercolesb = DateDiff(DateInterval.Hour, entb3.Value, salb3.Value)
                miercolest = tmiercoles - tmiercolesb
                ttotal = ttotal + miercolest
                miercoles = "' " & ent3.Value & "',' " & sal3.Value & "',' " & entb3.Value & "',' " & salb3.Value & "','" & tmiercoles & "','05/01/2000'"
            Else
                tmiercoles = DateDiff(DateInterval.Hour, ent3.Value, sal3.Value)
                ttotal = ttotal + tmiercoles
                miercoles = "' " & ent3.Value & "',' " & sal3.Value & "',NULL,NULL,'" & tmiercoles & "','05/01/2000'"
            End If
        Else
            miercoles = "Null,Null,Null,Null,Null,Null"
        End If
        If cb4.CheckState Then
            If cbb4.CheckState Then
                tjueves = DateDiff(DateInterval.Hour, ent4.Value, sal4.Value)
                tjuevesb = DateDiff(DateInterval.Hour, entb4.Value, salb4.Value)
                juevest = tjueves - tjuevesb
                ttotal = ttotal + juevest
                jueves = "' " & ent4.Value & "',' " & sal4.Value & "',' " & entb4.Value & "',' " & salb4.Value & "','" & tjueves & "','06/01/2000'"
            Else
                tjueves = DateDiff(DateInterval.Hour, ent4.Value, sal4.Value)
                ttotal = ttotal + tjueves
                jueves = "' " & ent4.Value & "',' " & sal4.Value & "',NULL,NULL,'" & tjueves & "','06/01/2000'"
            End If
        Else
            jueves = "Null,Null,Null,Null,Null,Null"
        End If
        If cb5.CheckState Then
            If cbb5.CheckState Then
                tviernes = DateDiff(DateInterval.Hour, ent5.Value, sal5.Value)
                tviernesb = DateDiff(DateInterval.Hour, entb5.Value, salb5.Value)
                viernest = tviernes - tviernesb
                ttotal = ttotal + viernest
                viernes = "' " & ent5.Value & "',' " & sal5.Value & "',' " & entb5.Value & "',' " & salb5.Value & "','" & tviernes & "','07/01/2000'"
            Else
                tviernes = DateDiff(DateInterval.Hour, ent5.Value, sal5.Value)
                ttotal = ttotal + tviernes
                viernes = "' " & ent5.Value & "',' " & sal5.Value & "',NULL,NULL,'" & tviernes & "','07/01/2000'"
            End If
        Else
            viernes = "Null,Null,Null,Null,Null,Null"
        End If

        If cb6.CheckState Then
            If cbb6.CheckState Then
                tsabado = DateDiff(DateInterval.Hour, ent6.Value, sal6.Value)
                tsabadob = DateDiff(DateInterval.Hour, entb6.Value, salb6.Value)
                sabadot = tsabado - tsabadob
                ttotal = ttotal + sabadot
                sabado = "' " & ent6.Value & "',' " & sal6.Value & "',' " & entb6.Value & "',' " & salb6.Value & "','" & tsabado & "','08/01/2000'"
            Else
                tsabado = DateDiff(DateInterval.Hour, ent6.Value, sal6.Value)
                ttotal = ttotal + tsabado
                sabado = "' " & ent6.Value & "',' " & sal6.Value & "',NULL,NULL,'" & tsabado & "','08/01/2000'"
            End If
        Else
            sabado = "Null,Null,Null,Null,Null,Null"
        End If
        If cb7.CheckState Then
            If cbb7.CheckState Then
                tdomingo = DateDiff(DateInterval.Hour, ent7.Value, sal7.Value)
                tdomingob = DateDiff(DateInterval.Hour, entb7.Value, salb7.Value)
                domingot = tdomingo - tdomingob
                ttotal = ttotal + domingot
                domingo = "' " & ent7.Value & "',' " & sal7.Value & "',' " & entb7.Value & "',' " & salb7.Value & "','" & tdomingo & "','09/01/2000'"
            Else
                tdomingo = DateDiff(DateInterval.Hour, ent7.Value, sal7.Value)
                ttotal = ttotal + tdomingo
                domingo = "' " & ent7.Value & "',' " & sal7.Value & "',NULL,NULL,'" & tdomingo & "','09/01/2000'"
            End If
        Else
            domingo = "Null,Null,Null,Null,Null,Null"
        End If

        'ttotal = (tlunes + tmartes + tmiercoles + tjueves + tviernes + tsabado + tdomingo) / 60
        ttotals = ttotal & ":00"

        MsgBox(" TTOTALS: " & ttotals)

        consulta = "insert into horario (nombre,tipo,[entrada(0)],[salida(0)],[salidabreak(0)],[regresobreak(0)],[tiempo(0)],[comienza(0)], " & _
              "[entrada(1)],[salida(1)],[salidabreak(1)],[regresobreak(1)],[tiempo(1)],[comienza(1)]," & _
              "[entrada(2)],[salida(2)],[salidabreak(2)],[regresobreak(2)],[tiempo(2)],[comienza(2)]," & _
              "[entrada(3)],[salida(3)],[salidabreak(3)],[regresobreak(3)],[tiempo(3)],[comienza(3)]," & _
              "[entrada(4)],[salida(4)],[salidabreak(4)],[regresobreak(4)],[tiempo(4)],[comienza(4)]," & _
              "[entrada(5)],[salida(5)],[salidabreak(5)],[regresobreak(5)],[tiempo(5)],[comienza(5)]," & _
              "[entrada(6)],[salida(6)],[salidabreak(6)],[regresobreak(6)],[tiempo(6)],[comienza(6)])" & _
              "values ('prueba22','Fijo' " & "," & lunes & "," & martes & "," & miercoles & "," & jueves & _
              "," & viernes & "," & sabado & "," & domingo & ")"

        basedatos = New OleDbDataAdapter(consulta, Conexion)
        cadenaconexion = New OleDbConnection(Conexion)

        Try
            cadenaconexion.Open()
            comando = New OleDbCommand(consulta, cadenaconexion)
            MessageBox.Show(consulta)
            comando.ExecuteNonQuery()
            MessageBox.Show(consulta)
            horarios = New DataTable
            basedatos.Fill(horarios)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        cadenaconexion.Close()
        DataGridView1.DataSource = horarios


        'Open "c:\consulta.txt" For Output As #1

        'Print #1, consulta

        'Close #1


        '  MsgBox(consulta)

        'DataEnvironment1.Connection1.Execute(consulta)

        'ent1.Value = Format(ent1.Value, "03/01/2000 HH:MM")
        'DataEnvironment1.rscom.MoveLast()

        '    DataEnvironment1.rscom.AddNew
        '    Dim consulta As String
        '    consulta = "insert into empleado (id,numero,nombre) values (1,2322,adrian)"

        'DataEnvironment1.rsCommand1.AddNew()
        ''este comentario es una prueba

    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        fechaini.Text = Convert.ToDateTime(DataGridView1.CurrentRow.Cells(1).Value()).Date
        fechafin.Text = Convert.ToDateTime(DataGridView1.CurrentRow.Cells(2).Value()).Date
    End Sub

    
   
    Private Sub DataGridView2_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEnter
        txtnombre.Text = Convert.ToString(DataGridView2.CurrentRow.Cells(2).Value) & " " & Convert.ToString(DataGridView2.CurrentRow.Cells(3).Value)
    End Sub
End Class
