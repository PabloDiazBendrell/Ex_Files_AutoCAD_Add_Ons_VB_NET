﻿Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Public Class Initialization
    Implements IExtensionApplication

    <CommandMethod("LISelectEntity")> Public Sub cmdSelEnt()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor


        Using db As Database = HostApplicationServices.WorkingDatabase
            Try
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Try
                        Dim objId As ObjectId
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
