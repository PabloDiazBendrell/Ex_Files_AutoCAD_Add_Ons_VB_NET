Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices

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
                RaisePropertyChanged("IsFrozen")
            End Set
        End Property
    End Class

    Public Class LayerObjectCollection
        Inherits ObjectModel.ObservableCollection(Of LayerObject)

        Public Sub GetCollectionFromDrawing(DwgDatabase As Database)
            Using trans As Transaction = DwgDatabase.TransactionManager.StartTransaction
                Try
                    Dim lyrTbl As LayerTable = trans.GetObject(DwgDatabase.LayerTableId, OpenMode.ForRead)
                    Dim lyrTblRec As LayerTableRecord
                    Dim lyrObj As LayerObject

                    For Each lyrId As ObjectId In lyrTbl
                        lyrTblRec = trans.GetObject(lyrId, OpenMode.ForRead)
                        lyrObj = New LayerObject
                        With lyrObj
                            .BaseObjectId = lyrId
                            .IsFrozen = lyrTblRec.IsFrozen
                            .Name = lyrTblRec.Name
                        End With
                        Me.Add(lyrObj)
                    Next
                    trans.Commit()
                Catch ex As Exception
                    Application.ShowAlertDialog("Error in LayerObjectCollection->GetCollectionFromDrawing->Transaction" &
                                                vbLf & ex.Message)
                End Try
            End Using
        End Sub


    End Class
End Namespace

