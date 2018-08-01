Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Initialization
    Implements IExtensionApplication

    <CommandMethod("LIGetPoint")> Public Sub cmdGetPoint()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Dim prPtOpts As New PromptPointOptions(vbLf & "Pick a point:")

        With prPtOpts
            .AllowArbitraryInput = False
            .AllowNone = True
        End With

        Dim prPtRes1 As PromptPointResult = ed.GetPoint(prPtOpts)
        If prPtRes1.Status <> PromptStatus.OK Then Exit Sub

        With prPtOpts
            .BasePoint = prPtRes1.Value
            .UseBasePoint = True
        End With

        Dim prPtRes2 As PromptPointResult = ed.GetPoint(prPtOpts)
        If prPtRes2.Status <> PromptStatus.OK Then Exit Sub

        ed.WriteMessage(vbLf & prPtRes2.Value.ToString)


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
