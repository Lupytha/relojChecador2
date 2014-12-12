Imports System
Imports System.IO
Imports System.Text
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.TimeSpan
Imports System.Data.DataRow

Public Class Form2
    Public fechaini As DateTime
    Public fechafin As DateTime
    'Dim irrow As Integer
    Dim Excel As Object
    Dim libro As Object
    Dim hoja As Object
    Dim retardo As Integer
    Dim contador As Integer
    Dim sretardo As String
    Dim posn As Integer
    Dim auxpos As Integer
    Dim posr As Integer
    Dim idempleado As String
    Dim basedatos As OleDbDataAdapter
    Dim empleados As DataTable
    Dim selectEmpleados As String
    Dim Conexion As String
    Dim n As Integer
    Dim selectEventos As String
    Dim eventos As DataTable
    Dim cadenaconexion As OleDbConnection
    Dim fechin As String
    Dim fechfi As String
    Dim idhorario As String
    Dim selectHorario As String
    Dim horario As DataTable
    Dim faltas As Integer
    Dim faltasretardo As Decimal
    Dim aux As Integer
    Dim selectDia As String
    Dim FechaActual As DateTime
    Dim exd As Integer
    Dim entr As String
    Dim salida As String
    Dim tabladias As DataTable
    Dim auxsa As String
    Dim auxen As String
    Dim dia As Integer
    Dim Cs, Ce, Ceb, Csb As Integer
    Dim hen, hes, hsab, henb As String
    Dim resta As String
    Dim Fila As Integer
    Dim dias As Long
    Dim posf As Integer
    Dim auxcount As Integer
    Dim pr As New Printing.PrintDocument
    Dim impresora As String
    Dim oldprinter, newprinter As Printing.PrintDocument
    Dim FechaActual2 As Date
    Dim entr2 As Date
    Dim rowev As DataRow
   
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
        'TODO: esta línea de código carga datos en la tabla 'DataSet1.Empleados' Puede moverla o quitarla según sea necesario.
        ProgressBar1.Visible = False
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 15
        ProgressBar1.Value = 0

        'Fecha de inicio de consulta
        Fecha2.Value = Now()
        fechaini = Convert.ToDateTime(Fecha1.Value)
        fechafin = Convert.ToDateTime(Fecha2.Value)
        Conexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='C:\Program Files (x86)\TimeWork Reloj Checador\fcf.mdb';" & _
                    "Persist Security Info=True;Jet OLEDB:Database Password=YaDRMlMPtdwYEAdtDNtL"
        selectEmpleados = "SELECT id, nombre, numero, apellidos FROM  empleado where activo=true"
        basedatos = New OleDbDataAdapter(selectEmpleados, Conexion)
        cadenaconexion = New OleDbConnection(Conexion)

        Try
            cadenaconexion.Open()
            empleados = New DataTable
            basedatos.Fill(empleados)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        cadenaconexion.Close()
        DataGridView1.DataSource = empleados
    End Sub

    Private Sub BtnGenerar_Click(sender As Object, e As EventArgs) Handles BtnGenerar.Click
        Excel = CreateObject("Excel.Application")
        libro = Excel.workbooks.open("c:\reportefcf2.xls")
        Excel.visible = False
        hoja = Excel.activesheet
        ProgressBar1.Visible = True

        'irrow = 1
        contador = 0
        retardo = 0
        sretardo = ""
        hoja.cells(8, 5) = fechaini.Date
        hoja.cells(8, 7) = fechafin.Date

        dias = DateDiff(DateInterval.Day, fechaini, fechafin)
        posn = 11
        posr = 14
        fechafin = fechafin.AddDays(1)
        FechaActual = Format(fechaini.Date, "dd/MM/yyyy")

        If empleados.Rows.Count = 0 Then
            MessageBox.Show("No se encontro ningun dato en la consulta")
        Else
            For Each row As DataRow In empleados.Rows

                If empleados.Rows.Count > 0 Then
                    posn = posn + auxpos
                    posr = posr + auxpos
                End If

                If DateDiff(DateInterval.Day, fechaini, fechafin) > 7 Then
                    auxpos = 39
                Else
                    auxpos = 33
                End If

                hoja.cells(posn, 2) = "NOMBRE"
                hoja.cells(posn, 3) = CStr(row(1))
                hoja.cells(posn, 5) = CStr(row(3))
                hoja.cells(posn + 1, 2) = "NUMERO"
                hoja.cells(posn + 1, 3) = CStr(row(2))
                hoja.cells(posn, 2).Font.Bold = True
                hoja.cells(posn + 1, 2).Font.Bold = True

                idempleado = CStr(row("Id"))

                selectEventos = "SELECT idempleado, idhorario, Format(Entrada,'mm/dd/yyyy') AS [Entradas], " & _
                    "Format(Salida,'mm/dd/yyyy') AS [Salidas] FROM Evento WHERE idempleado = " & idempleado & _
                    "AND Entrada BETWEEN #" & fechaini.Month & "/" & fechaini.Day & "/" & fechaini.Year &
                    "# AND #" & fechafin.Month & "/" & fechafin.Day & "/" & fechafin.Year & "# order by Entrada ASC"

                basedatos = New OleDbDataAdapter(selectEventos, Conexion)

                Try
                    cadenaconexion.Open()
                    eventos = New DataTable
                    basedatos.Fill(eventos)
                    rowev = eventos.Rows(1)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
                cadenaconexion.Close()
                DataGridView2.DataSource = eventos
                DataGridView2.Refresh()
                If eventos.Rows.Count > 0 Then
                    idhorario = CStr(rowev("Idhorario"))
                    selectHorario = "Select * from horario where id=" & idhorario
                    basedatos = New OleDbDataAdapter(selectHorario, Conexion)

                    Try
                        cadenaconexion.Open()
                        horario = New DataTable
                        basedatos.Fill(horario)
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try
                    cadenaconexion.Close()
                    DataGridView3.DataSource = horario

                    aux = 5
                    faltas = 0
                    retardo = 0

                    hoja.cells(posr - 1, 2) = "FECHA"
                    hoja.cells(posr - 1, 3) = "ENTRADA"
                    hoja.cells(posr - 1, 4) = "HORARIO"
                    hoja.cells(posr - 1, 5) = "SALIDA"
                    hoja.cells(posr - 1, 6) = "HORARIO"
                    hoja.cells(posr - 1, 7) = "OBSEREVACION"
                    hoja.rows(posr - 1).Font.Bold = True

                    Fila = posr

                    For Each rowhor As DataRow In horario.Rows
                        For n As Integer = 0 To dias '+ 2 ' - 1
                            FechaActual = fechaini.AddDays(n)
                            selectDia = "Select idempleado, idhorario, Entrada, " & _
                            "Salida from evento where idempleado =" & idempleado & _
                            "and datevalue(entrada) like #" & _
                            FechaActual.Month & "/" & FechaActual.Day & "/" & FechaActual.Year & "# order by Entrada ASC"

                            basedatos = New OleDbDataAdapter(selectDia, Conexion)

                            Try
                                cadenaconexion.Open()
                                tabladias = New DataTable
                                basedatos.Fill(tabladias)
                            Catch ex As Exception
                                Console.WriteLine(ex.Message)
                            End Try
                            cadenaconexion.Close()
                            DataGridView4.DataSource = tabladias
                            DataGridView4.Refresh()

                            exd = tabladias.Rows.Count
                            contador = 0

                            If exd = 0 Then
                                dia = Weekday(FechaActual)

                                If dia = 1 Then
                                    Ce = 39
                                    Cs = 40
                                    Ceb = 41
                                    Csb = 42
                                Else
                                    Ce = 3 + ((dia - 2) * 6)
                                    Cs = 4 + ((dia - 2) * 6)
                                    Ceb = 5 + ((dia - 2) * 6)
                                    Csb = 6 + ((dia - 2) * 6)
                                End If

                                If Convert.ToString(rowhor(Ce)) <> "" Then
                                    hoja.cells(Fila, 2) = FechaActual.Date
                                    hoja.cells(Fila, 7) = "FALTA"
                                    faltas = faltas + 1
                                Else
                                    hoja.cells(Fila, 2) = FechaActual.Date
                                    hoja.cells(Fila, 7) = "DIA DE DESCANSO"
                                End If
                            End If

                            For Each rowdia As DataRow In tabladias.Rows
                                idhorario = rowdia("idhorario")
                                If exd > 0 Then

                                    entr = CStr(rowdia(2))
                                    If Convert.ToString(rowdia(3)) <> "" Then
                                        salida = CStr(rowdia(3))
                                        auxen = entr.Substring(11, 13)
                                        auxsa = salida.Substring(11, 13)
                                    End If

                                    dia = Weekday(entr)

                                    If dia = 1 Then
                                        Ce = 39
                                        Cs = 40
                                        Ceb = 41
                                        Csb = 42
                                    Else
                                        Ce = 3 + ((dia - 2) * 6)
                                        Cs = 4 + ((dia - 2) * 6)
                                        Ceb = 5 + ((dia - 2) * 6)
                                        Csb = 6 + ((dia - 2) * 6)
                                    End If

                                    If Convert.ToString(rowhor(Ceb)) = "" Then
                                        hes = rowhor(Cs).ToString
                                        hen = rowhor(Ce).ToString
                                        If hes <> "" Then
                                            hes = hes.Substring(11, 13)
                                        End If
                                        If hen <> "" Then
                                            hen = hen.Substring(11, 13)
                                        End If

                                        contador = 1
                                    Else
                                        hes = rowhor(Ceb).ToString
                                        hen = rowhor(Ce).ToString

                                        If hes <> "" Then
                                            hes = hes.Substring(11, 13)
                                        End If
                                        If hen <> "" Then
                                            hen = hen.Substring(11, 13)
                                        End If
                                        contador = contador + 1
                                        henb = Convert.ToString(rowhor(Ceb))
                                        hsab = Convert.ToString(rowhor(Csb))
                                        If exd = 1 Then
                                            retardo = retardo + 1
                                            hoja.cells(Fila + 1, 2) = FechaActual.Date
                                            hoja.cells(Fila + 1, 3) = "Sin Registro"
                                            hoja.cells(Fila + 1, 5) = "Sin Registro"
                                            hoja.cells(Fila + 1, 7) = "RETARDO (FALTA REGISTRO DE BREAK)"
                                        End If
                                    End If

                                    If entr = "" Then
                                        entr = "Sin registro            "
                                    End If

                                    entr2 = CDate(entr)

                                    If DateDiff(DateInterval.Day, FechaActual, fechafin) > 0 And contador = 1 Then
                                        If DateDiff(DateInterval.Day, entr2, FechaActual) = 0 Then
                                            If Convert.ToString(hen) <> "" Then
                                                hen = DateTime.Parse(CStr(rowhor(Ce)))
                                                resta = DateDiff(DateInterval.Minute, TimeValue(hen), TimeValue(entr2)) 'TimeValue(auxen))
                                            Else
                                                resta = 0
                                            End If
                                            If auxen > hen And resta > "00:05:00" Then
                                                retardo = retardo + 1
                                                hoja.cells(Fila, 7) = "RETARDO (ENTRADA TARDE)"
                                            End If

                                            If Convert.ToString(hen) = "" Then
                                                hoja.cells(Fila + 1, 7) = "DESCANSO LABORADO"
                                            End If

                                            If salida = "" Then
                                                hoja.cells(Fila + 1, 2) = FechaActual.Date
                                                hoja.cells(Fila + 1, 5) = "Sin registro"
                                                hoja.cells(Fila + 1, 7) = "RETARDO (FALTA REGISTRO DE SALIDA)"
                                                retardo = retardo + 1
                                            End If

                                            If rowhor(Ce).ToString <> "" Then
                                                hoja.cells(Fila, 2) = FechaActual.Date
                                                hoja.cells(Fila, 3) = entr.Substring(11, 13)
                                                hoja.cells(Fila, 4) = Convert.ToString(rowhor(Ce)).Substring(11, 13)
                                                hoja.cells(Fila, 5) = salida.Substring(11, 13)
                                                hoja.cells(Fila, 6) = hes
                                                'hoja.cells(Fila, 8) = idhorario
                                            End If
                                            posf = posr + dias + 1
                                            ' End If
                                        End If

                                        If contador = 1 And Convert.ToString(rowhor(Ceb)) <> "" Then
                                            hoja.cells(Fila + 1, 2) = FechaActual.Date
                                            If hoja.cells(Fila + 1, 5).ToString = "Sin Registro" Then
                                            Else
                                                Fila = Fila + 1
                                            End If
                                        End If
                                    End If

                                    If exd = 2 And contador = 2 Then
                                        entr = CStr(rowdia(2))
                                        auxen = entr.Substring(11, 13)
                                        If DateDiff(DateInterval.Day, FechaActual, fechafin) > 0 Then
                                            If DateDiff(DateInterval.Day, entr2, FechaActual) = 0 Then
                                                If Convert.ToString(hen) <> "" Then
                                                    If Convert.ToString(rowhor(Ce)) <> "" Then
                                                        hen = DateTime.Parse(CStr(rowhor(Ce)))
                                                        resta = DateDiff(DateInterval.Minute, TimeValue(hen), TimeValue(entr2)) 'TimeValue(auxen))
                                                    End If
                                                Else
                                                    resta = 0
                                                End If

                                                If dia = 1 Then
                                                    Ce = 39
                                                    Cs = 40
                                                    Ceb = 41
                                                    Csb = 42
                                                Else
                                                    Ce = 3 + ((dia - 2) * 6)
                                                    Cs = 4 + ((dia - 2) * 6)
                                                    Ceb = 5 + ((dia - 2) * 6)
                                                    Csb = 6 + ((dia - 2) * 6)
                                                End If

                                                hsab = Convert.ToString(rowhor(Csb))
                                                hes = Convert.ToString(rowhor(Cs))

                                                If Convert.ToString(rowdia(3)) <> "" Then
                                                    salida = CStr(rowdia(3))
                                                    auxen = entr.Substring(11, 13)
                                                End If

                                                If henb <> "" And hsab <> "" Then
                                                    resta = DateDiff(DateInterval.Minute, TimeValue(hsab), TimeValue(auxen))
                                                    If auxen > hsab And resta > "00:05:00" Then
                                                        retardo = retardo + 1
                                                        hoja.cells(Fila, 7) = "RETARDO (ENTRADA DE BREAK TARDE)"
                                                    End If
                                                    If salida = "" Then
                                                        hoja.cells(Fila + 1, 2) = FechaActual.Date
                                                        hoja.cells(Fila + 1, 3) = "Sin registro"
                                                        hoja.cells(Fila + 1, 5) = "Sin registro"
                                                        hoja.cells(Fila + 1, 7) = "RETARDO"
                                                        hoja.cells(Fila, 7) = "RETARDO (FALTA REGISTRO DE SALIDA)"
                                                        retardo = retardo + 1
                                                    End If
                                                    hoja.cells(Fila, 2) = FechaActual.Date 'entr.Substring(0, 10)
                                                    hoja.cells(Fila, 3) = entr.Substring(11, 13)
                                                    hoja.cells(Fila, 4) = hsab.Substring(11, 13)
                                                    hoja.cells(Fila, 5) = salida.Substring(11, 13)
                                                    hoja.cells(Fila, 6) = hes.Substring(11, 13)
                                                    'hoja.cells(Fila, 8) = idhorario
                                                End If
                                                hoja.cells(Fila, 2).Interior.ColorIndex = 24
                                                hoja.cells(Fila, 3).Interior.ColorIndex = 24
                                                hoja.cells(Fila, 4).Interior.ColorIndex = 24
                                                hoja.cells(Fila, 5).Interior.ColorIndex = 24
                                                hoja.cells(Fila, 6).Interior.ColorIndex = 24
                                                hoja.cells(Fila, 7).Interior.ColorIndex = 24
                                            End If
                                        End If
                                    End If
                                End If
                                ' Si no hay registros en el dia, hay que verificar si es dia de descanso o falta 
                                If ProgressBar1.Value < ProgressBar1.Maximum Then
                                    ProgressBar1.Value += 1
                                    If ProgressBar1.Value = ProgressBar1.Maximum Then
                                        ProgressBar1.Value = 14
                                        ProgressBar1.Value = 0
                                    End If
                                End If
                            Next
                            Fila = Fila + 1
                        Next
                    Next
                End If
                faltasretardo = retardo / 3
                hoja.cells(1 + Fila, 2) = "Faltas"
                hoja.cells(1 + Fila, 3) = faltas
                hoja.cells(1 + Fila + 1, 2) = "Retardos"
                hoja.cells(1 + Fila + 1, 3) = retardo
                hoja.cells(1 + Fila + 2, 2) = "Faltas por retardos: "
                hoja.cells(1 + Fila + 2, 3) = Decimal.Truncate(faltasretardo)
                hoja.cells(1 + Fila + 3, 2) = "Faltas a descontar: "
                hoja.cells(1 + Fila + 3, 3) = faltas + Decimal.Truncate(faltasretardo)
                hoja.cells(1 + Fila + 3, 5) = "FIRMA"
                hoja.cells(1 + Fila + 3, 5).Font.Bold = True
                hoja.cells(1 + Fila + 3, 6) = "                                                      "
                hoja.cells(1 + Fila + 3, 6).Font.Underline = True
            Next
        End If
        ProgressBar1.Value = 15
        MessageBox.Show("REPORTE GENERADO")
        hoja.rows(1).Font.Bold = True
        Excel.Visible = True
        'hoja.printout()
        auxcount = 0
        ' impresora = PrinterSettings.InstalledPrinters
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Fecha1_ValueChanged_1(sender As Object, e As EventArgs) Handles Fecha1.ValueChanged
        fechaini = Convert.ToDateTime(Fecha1.Value)
    End Sub

    Private Sub Fecha2_ValueChanged(sender As Object, e As EventArgs) Handles Fecha2.ValueChanged
        fechafin = Convert.ToDateTime(Fecha2.Value)
    End Sub

End Class