Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices

Public Class Initialization
    Implements IExtensionApplication

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

    Public Sub Initialize() Implements IExtensionApplication.Initialize


    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate

    End Sub

    Private Function IsSavedFile() As Boolean
        If Application.GetSystemVariable("DWGTITLED") <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
