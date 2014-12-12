Imports System.Data.OleDb

Public Class Form4
    Dim basedatos As OleDbDataAdapter
    Dim Conexion, selectsemanas, selectempleados, selecthorario, selecthorario2, selecteventos As String
    Dim cadenaconexion As OleDbConnection
    Dim horarios, horarios2, horarios3 As DataTable
    Dim semanas, empleados, eventos As DataTable
    Dim comando As OleDbCommand
    Dim consulta, nomhorario, borrahorario As String
    Dim lunes, martes, miercoles, jueves, viernes, sabado, domingo As String
    Dim tlunes, tmartes, tmiercoles, tjueves, tviernes, tsabado, tdomingo, ttotals As String
    Dim tlunesb, tmartesb, tmiercolesb, tjuevesb, tviernesb, tsabadob, tdomingob As String
    Dim lunest, martest, miercolest, juevest, viernest, sabadot, domingot As Integer
    Dim ttotal As Integer
    Dim result As DialogResult
    Dim idempleado, numsemana As String
    Public fechaini As DateTime
    Public fechafin As DateTime

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
        Conexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Program Files (x86)\TimeWork Reloj Checador\fcf.mdb';" & _
                  "Persist Security Info=True;Jet OLEDB:Database Password=YaDRMlMPtdwYEAdtDNtL"

        selectempleados = "SELECT id, numero, nombre, apellidos FROM  empleado where activo=true order by numero"
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

        selectsemanas = "SELECT * FROM semanas WHERE YEAR(fechaini) = YEAR(Date())"
        basedatos = New OleDbDataAdapter(selectsemanas, Conexion)
        cadenaconexion = New OleDbConnection(Conexion)

        Try
            cadenaconexion.Open()
            semanas = New DataTable
            basedatos.Fill(semanas)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        cadenaconexion.Close()
        ComboBox1.DataSource = semanas
        DataGridView1.DataSource = semanas

    End Sub

    Private Sub DataGridView2_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellEnter
        txtnombre.Text = Convert.ToString(DataGridView2.CurrentRow.Cells(2).Value) & " " & Convert.ToString(DataGridView2.CurrentRow.Cells(3).Value)
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        Fecha1.Text = Convert.ToDateTime(DataGridView1.CurrentRow.Cells(1).Value())
        Fecha2.Text = Convert.ToDateTime(DataGridView1.CurrentRow.Cells(2).Value())
    End Sub

    Private Sub Salir_Click(sender As Object, e As EventArgs) Handles Salir.Click
        fechaini = Convert.ToDateTime(Fecha1.Text)
        fechafin = Convert.ToDateTime(Fecha2.Text)

        idempleado = ComboBox2.Text
        numsemana = ComboBox1.Text

        ''selecteventos = "SELECT CONVERT(DATETIME, [Entrada(0)], 108),8 as [Entrada Lunes], [Salida(0)], [SalidaBreak(0)], [RegresoBreak(0)], " & _
        ''    "[entrada(1)], [Salida(1)], [SalidaBreak(1)], [RegresoBreak(1)], " & _
        ''    "[entrada(2)], [Salida(2)], [SalidaBreak(2)], [RegresoBreak(2)], " & _
        ''    "[entrada(3)], [Salida(3)], [SalidaBreak(3)], [RegresoBreak(3)], " & _
        ''    "[entrada(4)], [Salida(4)], [SalidaBreak(4)], [RegresoBreak(4)], " & _
        ''    "[entrada(5)], [Salida(5)], [SalidaBreak(5)], [RegresoBreak(5)], " & _
        ''    "[entrada(6)], [Salida(6)], [SalidaBreak(6)], [RegresoBreak(6)] " & _
        ''    "FROM horario WHERE nombre=" & "'" & idempleado & " " & numsemana & "'"
        selecteventos = "SELECT RIGHT(CONVERT(VARCHAR, [Entrada(0)], 108),8) AS FECHA FROM HORARIO"
        ' 'idempleado, idhorario, Format(Entrada,'mm/dd/yyyy') AS [Entradas], " & _
        ''  "Format(Salida,'mm/dd/yyyy') AS [Salidas] FROM Evento WHERE idempleado = " & idempleado & _
        ''  " AND Entrada BETWEEN #" & fechaini.Month & "/" & fechaini.Day & "/" & fechaini.Year &
        ''  "# AND #" & fechafin.Month & "/" & fechafin.Day & "/" & fechafin.Year & "# order by Entrada ASC"


        basedatos = New OleDbDataAdapter(selecteventos, Conexion)

        Try
            cadenaconexion.Open()
            eventos = New DataTable
            MessageBox.Show(selecteventos)
            basedatos.Fill(eventos)
            MessageBox.Show(selecteventos)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        cadenaconexion.Close()
        DataGridView3.DataSource = eventos
        DataGridView3.Refresh()
    End Sub
End Class