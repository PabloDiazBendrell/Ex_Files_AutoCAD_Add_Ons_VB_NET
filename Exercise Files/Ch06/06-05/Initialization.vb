Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Initialization
    Implements IExtensionApplication

    <CommandMethod("LIGetKeyword")> Public Sub cmdGetKWord()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Dim prDistOpts As New PromptDistanceOptions(vbLf & "Specify a distance:")
        With prDistOpts
            .AllowArbitraryInput = False
            .AllowNegative = False
            .AllowNone = True
            .AllowZero = False
            .DefaultValue = 1
            .Only2d = True
            .UseDefaultValue = True

        End With

        Dim prDistRes As PromptDoubleResult = ed.GetDistance(prDistOpts)
        If prDistRes.Status <> PromptStatus.OK Then Exit Sub

        ed.WriteMessage(vbLf & "Distance is " & Math.Round(prDistRes.Value, HostApplicationServices.WorkingDatabase.Luprec, MidpointRounding.AwayFromZero))
    End Sub


#Region "Commands"
    <CommandMethod("LIGetDistance")> Public Sub cmdGetDist()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Dim prDistOpts As New PromptDistanceOptions(vbLf & "Specify a distance:")
        With prDistOpts
            .AllowArbitraryInput = False
            .AllowNegative = False
            .AllowNone = True
            .AllowZero = False
            .DefaultValue = 1
            .Only2d = True
            .UseDefaultValue = True
        End With

        Dim prDistRes As PromptDoubleResult = ed.GetDistance(prDistOpts)
        If prDistRes.Status <> PromptStatus.OK Then Exit Sub

        ed.WriteMessage(vbLf & "Distance is " & prDistRes.Value.ToString)
        ed.WriteMessage(vbLf & "Distance is " & Math.Round(prDistRes.Value, HostApplicationServices.WorkingDatabase.Luprec, MidpointRounding.AwayFromZero))

    End Sub

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

    <CommandMethod("LISelectEntity")> Public Sub cmdSelEnt()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Dim prSelOpts As New PromptEntityOptions(vbLf & "Select a line:")
        With prSelOpts
            .AllowNone = True
            .AllowObjectOnLockedLayer = True
            .SetRejectMessage(vbLf & "Selected object has to be a line.")
            .AddAllowedClass(GetType(Line), True)
        End With

        Dim prEntRes As PromptEntityResult = ed.GetEntity(prSelOpts)
        If prEntRes.Status <> PromptStatus.OK Then Exit Sub


        Using db As Database = HostApplicationServices.WorkingDatabase
            Try
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Try
                        Dim objId As ObjectId = prEntRes.ObjectId
                        Dim lObj As Line = trans.GetObject(objId, OpenMode.ForRead)

                        ed.WriteMessage(vbLf & "Line distance is " & Math.Round(lObj.Length, db.Luprec, MidpointRounding.AwayFromZero))

                        trans.Commit()
                    Catch ex As Exception
                        Application.ShowAlertDialog("Error in LISelectEntity->Transaction" & vbLf & ex.Message)
                    End Try
                End Using
            Catch ex As Exception
                Application.ShowAlertDialog("Error in LISelectEntity->database" & vbLf & ex.Message)
            End Try
        End Using

    End Sub

    <CommandMethod("LISelectSet")> Public Sub cmdSelSet()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Dim prSelOpts As New PromptSelectionOptions
        With prSelOpts
            .AllowDuplicates = False
            .MessageForAdding = vbLf & "Select lines to count length of <exit>: "
            .MessageForRemoval = vbLf & "Select lines to remove from selection: "
            .RejectObjectsFromNonCurrentSpace = True
            .RejectObjectsOnLockedLayers = False
            .RejectPaperspaceViewport = True
        End With

        Dim tValues(0) As TypedValue
        tValues(0) = New TypedValue(DxfCode.Start, "LINE")
        Dim selFil As New SelectionFilter(tValues)

        Dim prSelRes As PromptSelectionResult = ed.GetSelection(prSelOpts, selFil)
        If prSelRes.Status <> PromptStatus.OK Then Exit Sub


        Using db As Database = HostApplicationServices.WorkingDatabase
            Try
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Try
                        Dim selSet As SelectionSet = prSelRes.Value
                        Dim lObj As Line
                        Dim lenTotal As Double = 0

                        For Each selObj As SelectedObject In selSet
                            lObj = trans.GetObject(selObj.ObjectId, OpenMode.ForRead)
                            lenTotal += lObj.Length
                        Next

                        ed.WriteMessage(vbLf & "Total length of selected objects is " & Math.Round(lenTotal, db.Luprec, MidpointRounding.AwayFromZero))

                        trans.Commit()
                    Catch ex As Exception
                        Application.ShowAlertDialog("Error in LISelectSet->Transaction" & vbLf & ex.Message)
                    End Try
                End Using
            Catch ex As Exception
                Application.ShowAlertDialog("Error in LISelectSet->database" & vbLf & ex.Message)
            End Try
        End Using

    End Sub
#End Region

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
