Imports System.Windows

Public Class winSelectLayer
    Public Sub New(Layers As Objects.LayerObjectCollection)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.DataContext = Layers
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub btnOk_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btnOk.Click
        Me.DialogResult = True
        Me.Close()
    End Sub
End Class
