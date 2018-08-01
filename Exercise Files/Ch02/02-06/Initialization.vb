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
        ed.WriteMessage(vbLf & "Hello World.")

    End Sub

    Public Sub Initialize() Implements IExtensionApplication.Initialize


    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate

    End Sub
End Class
