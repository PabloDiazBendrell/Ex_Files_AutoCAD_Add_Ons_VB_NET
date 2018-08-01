Imports Autodesk.AutoCAD.Runtime

Public Class Initialization
    Implements IExtensionApplication

    <CommandMethod("MyFirstCommand")> Public Sub cmdMyFirst()

    End Sub

    Public Sub Initialize() Implements IExtensionApplication.Initialize

    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate

    End Sub
End Class
