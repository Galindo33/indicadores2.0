Public Class DataBase
    'variables para almacenar y manipular la informacion
    'de la base de datos
    Dim ind_logro As Double
    Dim ind_causa As String
    Dim ind_plnAcn As String

    'propiedades para acceder y cambiar los valores
    'del logro dentro de la base de datos
    Public Property Logro() As Double
        Get
            Return Me.ind_logro
        End Get
        Set(value As Double)
            Me.ind_logro = value
        End Set
    End Property

    'propiedades para acceder y cambiar los valores
    'de la causa dentro de la base de datos
    Public Property Causa() As String
        Get
            Return Me.ind_causa
        End Get
        Set(value As String)
            Me.ind_causa = value
        End Set
    End Property

    'propiedades para acceder y cambiar los valores
    'del plan de accion dentro de la base de datos
    Public Property PlnAcn() As String
        Get
            Return Me.ind_plnAcn
        End Get
        Set(value As String)
            Me.ind_plnAcn = value
        End Set
    End Property


End Class
