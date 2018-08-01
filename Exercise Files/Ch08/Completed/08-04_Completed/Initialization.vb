Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Initialization
    Implements IExtensionApplication

    <CommandMethod("LIChangeLayer")> Public Sub cmdChangeLayers()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Dim prSelOpts As New PromptSelectionOptions
        With prSelOpts
            .AllowDuplicates = False
            .MessageForAdding = vbLf & "Select objects to change layer: "
            .MessageForRemoval = vbLf & "Remove objects from selection: "
            .RejectObjectsFromNonCurrentSpace = True
            .RejectObjectsOnLockedLayers = True
            .RejectPaperspaceViewport = True
        End With

        Dim prSelRes As PromptSelectionResult = ed.GetSelection(prSelOpts)
        If prSelRes.Status <> PromptStatus.OK Then Exit Sub

        Dim lyrs As New Objects.LayerObjectCollection
        lyrs.GetCollectionFromDrawing(HostApplicationServices.WorkingDatabase)

        Dim winLyr As New winSelectLayer(lyrs)
        If Application.ShowModalWindow(winLyr) <> True Then Exit Sub

        Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
            Try
                Dim ent As Entity

                For Each selObj As SelectedObject In prSelRes.Value
                    ent = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite)
                    ent.LayerId = winLyr.cboLayers.SelectedValue
                Next
                trans.Commit()
            Catch ex As Exception
                Application.ShowAlertDialog("Error in LIChangeLayer->Transcation" & vbLf & ex.Message)
            End Try
        End Using

        ed.WriteMessage(vbLf & prSelRes.Value.Count & " objects moved to " & winLyr.cboLayers.SelectedItem.Name & " layer.")
    End Sub

#Region "Support Functions"
    Private Function IsSavedFile() As Boolean
        If Application.GetSystemVariable("DWGTITLED") <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function IsLayerExists(CurrentTransaction As Transaction, Layers As LayerTable, LayerName As String) As Boolean
        Dim lyrObj As LayerTableRecord

        For Each lyrId As ObjectId In Layers
            lyrObj = CurrentTransaction.GetObject(lyrId, OpenMode.ForRead)
            If lyrObj.Name = LayerName Then Return True
        Next

        Return False
    End Function

#End Region

#Region "Extension Routines"
    Public Sub Initialize() Implements IExtensionApplication.Initialize

    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate

    End Sub

#End Region

End Class
