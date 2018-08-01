Imports System.ComponentModel
Imports Autodesk.AutoCAD.DatabaseServices

Namespace Objects
    Public Class BaseObject
        Implements ComponentModel.INotifyPropertyChanged

        Private _bName As String = ""
        Public Property Name As String
            Get
                Return _bName
            End Get
            Set(value As String)
                _bName = value
                RaisePropertyChanged("Name")
            End Set
        End Property

        Private _bId As ObjectId = Nothing
        Public Property BaseObjectId As ObjectId
            Get
                Return _bId
            End Get
            Set(value As ObjectId)
                _bId = value
                RaisePropertyChanged("BaseObjectId")
            End Set
        End Property

        Private _isSel As Boolean = False
        Public Property IsSelected As Boolean
            Get
                Return _isSel
            End Get
            Set(value As Boolean)
                _isSel = value
                RaisePropertyChanged("IsSelected")
            End Set
        End Property

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        Public Sub RaisePropertyChanged(ByVal propertyName As String)
            RaiseEvent PropertyChanged(Me, New ComponentModel.PropertyChangedEventArgs(propertyName))
        End Sub

    End Class
End Namespace

