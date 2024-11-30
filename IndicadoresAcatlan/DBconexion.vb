'antes de ejecutar se debera configurar de manera
'adecuada la conexion con la base de datos
'de lo contrario marca error.
'importaciones para la conexion con la base
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO

'importaciones para generar reportes formato excel
Imports OfficeOpenXml
Imports OfficeOpenXml.Style



Public Class DBconexion

    'variables para los parametros que usan las funciones
    Dim log As Double
    Dim causa As String
    Dim plnAcn As String
    Dim proy As Double


    Dim conexion As New SqlConnection()
    Dim direccion As ConnectionStringSettings
    Dim comando As SqlCommand

    Sub Conectar()

        'obtener nombre de la conexion desde el app config.
        direccion = ConfigurationManager.ConnectionStrings("nombre de la conexion")
        conexion.ConnectionString = direccion.ConnectionString
        conexion.Open()
        MsgBox("CONEXION EXITOSA")

    End Sub

    Sub Desconectar()

        conexion.Close()

    End Sub

    'metodo que almacena en sql la informacion ingresada por el usuario.
    Sub Insertar(logro, causa, plnAcn)

        Conectar()
        comando = New SqlCommand("insert into database values(" + logro + "," + causa + "," + plnAcn + ")", conexion)
        comando.ExecuteNonQuery()
        Desconectar()

    End Sub

    'metodo que valida el valor de la casilla <logro>
    'ingresado por el usuario.
    Sub Validacion(logro, proy)

        'si el porcentaje del logro con respecto a la proyeccion
        'es mayor a 80, las casillas de causa y plan de accion
        'no se habilitan para su escritura.
        If (logro * 100) / proy >= 80 Then

            'txtTarget1 ytxtTarget2 hacen referencia a los nombres
            'de las cajas de texto para causa y plan.
            'asegurarse de usar los nombres correctos.
            txtTarget1.Enabled = False
            txtTarget2.Enabled = False

            'si el porcentaje del logro con respecto a la proyeccion
            'es menor a 80, las casillas de causa y plan de accion
            'se habilitan para su escritura.
        ElseIf (logro * 100) / proy < 80 Then

            'txtTarget1 ytxtTarget2 hacen referencia a los nombres
            'de las cajas de texto para causa y plan.
            'asegurarse de usar los nombres correctos.
            txtTarget1.Enabled = True
            txtTarget2.Enabled = True

        End If
    End Sub


    'metodo que valida el valor de la casilla <logro>
    'ingresado por el usuario y dependiendo del porcentaje
    'pinta las casillas de color verde, amarillo o rojo.
    Sub Semaforos(logro, proy)

        'si el porcentaje del logro con respecto a la proyeccion
        'es mayor o igual a 80, la casilla de logro se pinta en verde.
        If (logro * 100) / proy >= 80 Then

            'txtlogro hace referencia al nombre
            'de la caja de texto para logro.
            'asegurarse de usar los nombres correctos.
            txtlogro.BackColor = System.Drawing.Color.FromArgb(0, 200, 0)


            'si el porcentaje del logro con respecto a la proyeccion
            'es menor a 80 y mayor a 75, la casilla de logro se pinta en amarillo.
        ElseIf (logro * 100) / proy < 80 And (logro * 100) / proy > 75 Then

            'txtlogro hace referencia al nombre
            'de la caja de texto para logro.
            'asegurarse de usar los nombres correctos.
            txtlogro.BackColor = System.Drawing.Color.FromArgb(0, 200, 200)

            'si el porcentaje del logro con respecto a la proyeccion
            'es menor o igual a 75, la casilla de logro se pinta en rojo.
        ElseIf (logro * 100) / proy <= 75 Then

            'txtlogro hace referencia al nombre
            'de la caja de texto para logro.
            'asegurarse de usar los nombres correctos.
            txtlogro.BackColor = System.Drawing.Color.FromArgb(200, 0, 0)


        End If
    End Sub

    Public Function Obtenerdatos() As DataTable

        Dim conectString As String = "Data Source=servidor;Initial Catalog=baseDatos;Integrated Security=True"
        Dim query As String = "SELECT * From tabla"

        Dim tabla As New DataTable()

        Using conexion As New SqlConnection(conectString)
            Using comand As New SqlCommand(query, conexion)

                conexion.Open()

                Using reader As SqlDataReader = comand.ExecuteReader

                    tabla.Load(reader)

                End Using

            End Using

        End Using

        Return tabla

    End Function


    'metodo que exporta a un archivo de excel los datos que se visualizan en pantalla 
    Public Sub ExportarExcel()

        Dim i, j As Integer
        Dim datos As DataTable = ObtenerDatos()

        'configurar EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial


        Using package As New ExcelPackage()

            Dim hoja As ExcelWorksheet = package.Workbook.Worksheets.Add("ReporteIndicadores")

            'Escribir encabezados del reporte
            For i = 1 To datos.Columns.Count()

                hoja.Cells(1, i).Value = datos.Columns(i - 1).ColumnName
                hoja.Cells(1, i).Style.Font.Bold = True
                hoja.Cells(1, i).Style.Border.BorderAround(ExcelBorderStyle.Thin)

            Next

            'Migrar datos de la base al reporte 
            For i = 0 To datos.Rows.Count - 1
                For j = 0 To datos.Columns.Count - 1

                    hoja.Cells(i + 2, j + 1).Value = datos.Rows(i)(j)
                    hoja.Cells(i + 2, j + 1).Style.Border.BorderAround(ExcelBorderStyle.Thin)

                Next
            Next

            'Guarda el reporte en la ruta indicada
            Dim rutaArchivo As String = "C:\Reportes\Reporte.xlsx"
            Directory.CreateDirectory(Path.GetDirectoryName(rutaArchivo))
            File.WriteAllBytes(rutaArchivo, package.GetAsByteArray())

            MessageBox.Show($"Reporte generado con éxito: {rutaArchivo}")

        End Using

    End Sub


End Class
