Public Class Produto

    Private _codigo As String
    Private _nome As String
    Private _precodecompra As Single
    Private _precodevenda As Single
    Private _Nexemplaressup As Integer
    Private _Nexemplarestotal As Integer
    Private _validade As Date

    Public Sub New(ByVal codigo As String, ByVal nome As String, ByVal precocompra As Single, ByVal precovenda As Single, ByVal nexemplaressup As Integer, ByVal numexemptotal As Integer, ByVal validade As Date)

        _codigo = codigo
        _nome = nome
        _precodecompra = precocompra
        _precodevenda = precovenda
        _Nexemplaressup = nexemplaressup
        _Nexemplarestotal = Nexemplarestotal
        _validade = validade

    End Sub

    Public Property Codigo As String
        Get
            Return _codigo
        End Get
        Set(value As String)
            _codigo = value
        End Set
    End Property

    Public Property Nome As String
        Get
            Return _nome
        End Get
        Set(value As String)
            _nome = value
        End Set
    End Property

    Public Property Precodecompra As Single
        Get
            Return _precodecompra
        End Get
        Set(value As Single)
            _precodecompra = value
        End Set
    End Property

    Public Property Precodevenda As Single
        Get
            Return _precodevenda
        End Get
        Set(value As Single)
            _precodevenda = value
        End Set
    End Property

    Public Property Nexemplaressup As Integer
        Get
            Return _Nexemplaressup
        End Get
        Set(value As Integer)
            _Nexemplaressup = value
        End Set
    End Property

    Public Property Nexemplarestotal As Integer
        Get
            Return _Nexemplarestotal
        End Get
        Set(value As Integer)
            _Nexemplarestotal = value
        End Set
    End Property

    Public Property Validade As Date
        Get
            Return _validade
        End Get
        Set(value As Date)
            _validade = value
        End Set
    End Property

    Sub AtivarDesconto()

        If DateDiff("d", Validade, Date.Now) <= 7 And DateDiff("d", Validade, Date.Now) > 3 Then

            Me.Precodevenda = Me.Precodevenda - Me.Precodevenda * 0.25

        ElseIf DateDiff("d", Validade, Date.Now) <= 3 And DateDiff("d", Validade, Date.Now) >= 0 Then

            Me.Precodevenda = Me.Precodevenda - Me.Precodevenda * 0.5

        End If

    End Sub

    Function ImprimirSTR() As String

        ImprimirSTR = Me.Codigo & " " & Me.Nome & " " & Me.Precodevenda & " " & Me.Validade
        Return ImprimirSTR
    End Function

    Function Dias(ByVal data As Date) As Integer


    End Function



End Class
