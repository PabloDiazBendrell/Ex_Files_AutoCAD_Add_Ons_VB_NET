Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices

Public Class Initialization
    Implements IExtensionApplication

#Region "Commands"
    <CommandMethod("MyFirstCommand")> Public Sub cmdMyFirst()
        Dim ed As Editor
        Dim doc As Document
        doc = Application.DocumentManager.MdiActiveDocument
        ed = doc.Editor

        If IsSavedFile() = True Then
            ed.WriteMessage(vbLf & "File is an existing file")
        Else
            ed.WriteMessage(vbLf & "File is a new drawing")
        End If

    End Sub

    <CommandMethod("LIGetVersion")> Public Sub cmdAcadVersion()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Select Case Application.DocumentManager.MdiActiveDocument.Database.LastSavedAsVersion
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1032
                ed.WriteMessage(vbLf & "AutoCAD 2018")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027
                ed.WriteMessage(vbLf & "AutoCAD 2013")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1024
                ed.WriteMessage(vbLf & "AutoCAD 2010")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1021
                ed.WriteMessage(vbLf & "AutoCAD 2007")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1800, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1800a
                ed.WriteMessage(vbLf & "AutoCAD 2004")
            Case Else
                ed.WriteMessage(vbLf & "Uknown or too old")
        End Select

    End Sub


#End Region


    Private Function IsSavedFile() As Boolean
        If Application.GetSystemVariable("DWGTITLED") <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function

#Region "Extension Routines"
    Public Sub Initialize() Implements IExtensionApplication.Initialize

    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate

    End Sub

#End Region
End Class
