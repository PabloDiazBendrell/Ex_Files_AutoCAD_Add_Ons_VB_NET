
Namespace Objects
    Public Class LayerObject
        Inherits BaseObject

        Private _isFrz As Boolean = False
        Public Property IsFrozen As Boolean
            Get
                Return _isFrz
            End Get
            Set(value As Boolean)
                _isFrz = value
            End Set
        End Property
    End Class
End Namespace

