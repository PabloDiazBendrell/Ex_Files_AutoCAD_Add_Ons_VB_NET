Imports Autodesk.AutoCAD.DatabaseServices

Namespace Objects
    Public Class BaseObject

        Private _bName As String = ""
        Public Property Name As String
            Get
                Return _bName
            End Get
            Set(value As String)
                _bName = value
            End Set
        End Property

        Private _bId As objectid = Nothing
        Public Property BaseObjectId As ObjectId
            Get
                Return _bId
            End Get
            Set(value As ObjectId)
                _bId = value
            End Set
        End Property

        Private _isSel As Boolean = False
        Public Property IsSelected As Boolean
            Get
                Return _isSel
            End Get
            Set(value As Boolean)
                _isSel = value
            End Set
        End Property
    End Class
End Namespace

