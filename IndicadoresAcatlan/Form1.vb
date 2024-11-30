'podemos crear tantas variables como cajas de texto dentro de nuestro formulario
'para acceder y manipular de forma dinamica cada valor, caja de texto etc, y adaptar los metodos
Public Class Form1
    'variables para almacenar los valores
    'de las cajas de texto.
    Dim log As Double
    Dim causa As String
    Dim planactn As String
    Dim proy As Double

    'variable que se puede utilizar como
    'indicador de cajas de texto y eventualmente
    'enviarse como parametro de la funcion de semaforos
    Dim cajaTexto As TextBox

    Dim dbconx As New DBconexion()

    'se asignan los valores de las cajas
    'a las variables.
    'asegurarse de que se utilizan los nombres correctos de las
    'cajas de texto en el formulario.
    log = logro.Text
    causa = causa.Text
    planactn = plan.Text
    proy = proy.Text

    
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    'evento que manda a llamar a la funcion <validar> una vez que el usuario 
    'ingresa informacion en la casilla de <logro>.
    'asegurarse de que el evento se usa con el nombre de la casilla <logro>.
    Private Sub txtTargetControl(sender As Object, e As EventArgs) Handles txtControl.TextChanged

        dbconx.Validacion(log, proy)

    End Sub

    'evento que manda a llamar a la funcion <semaforo> una vez que el usuario 
    'ingresa informacion en la casilla de <logro>.
    'asegurarse de que el evento se usa con el nombre de la casilla <logro>.
    Private Sub txtColor(sender As Object, e As EventArgs) Handles txtControl.TextChanged

        'se puede modificar la funcion, para que reciba un parametro mas
        'de tipo caja de texto 
        dbconx.Semaforos(log, proy)

    End Sub


    'evento que manda a llamar a la funcion <insertar> una vez que el usuario 
    'hace click en el boton de guargar informacion.
    'asegurarse de que el evento se usa con el nombre del boton que guarda los datos.
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnInsertar.Click

        'asegurarse de que el metodo se ejecute con el mismo numero de variables 
        'que cajas de texto en el formulario.
        dbconx.Insertar(log, causa, planactn)

    End Sub

    'evento que manda a llamar a la funcion <Exportar> una vez que el usuario hace click
    'en el boton para exportar el reporte
    'asegurarse de que el evento se usa con el nombre del boton que exporta los datos.
    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click

        dbconx.ExportarExcel()


    End Sub


End Class
